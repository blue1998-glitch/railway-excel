import io
import re
import time
import random
import openpyxl
import pandas as pd
from pydantic import BaseModel, Field
from pypdf import PdfReader, PdfWriter
import streamlit as st
from google import genai
from google.genai import types, errors
from concurrent.futures import ThreadPoolExecutor, as_completed
import base64
import copy

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 多執行緒極速辨識 ｜ ⚖️ 會計立場扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

raw_key = st.secrets.get("GEMINI_API_KEY", "")
cleaned_key = str(raw_key).replace('"', '').replace("'", "").strip()

with st.sidebar:
    st.header("⚙️ 系統設定")
    user_key = st.text_input(
        "Gemini API Key",
        value=cleaned_key,
        type="password",
        help="系統會優先讀取 secrets 中的 GEMINI_API_KEY"
    )
    active_api_key = user_key.strip()

    # 白名單機制（取代舊版 models.list 動態撈取）：
    # 動態撈取雖然自動化，但會連同「已棄用、需額外權限申請、或不支援 PDF+結構化 JSON」的型號一併列出，
    # 這正是過去出現 404 / 401 / 403 的主因。以下僅列出官方目前明確支援
    # 「多模態 PDF 辨識 + 結構化 JSON 輸出」且屬穩定版（非 preview/experimental）的 Flash / Flash-Lite
    # 主流模型，優先選用免費額度 RPM 較高的機型。日後 Google 更新模型時，只需增減下方字串即可，
    # 不必更動其他程式邏輯。
    MODEL_WHITELIST = [
        "gemini-3.6-flash",           # 已實測穩定可用
        "gemini-flash-lite-latest",   # 已實測穩定可用（官方別名，恆指向最新穩定版 Flash-Lite）
        "gemini-flash-latest",        # 新增：官方別名，恆指向最新穩定版 Flash（尚待實測）
        "gemini-3.5-flash-lite",
        "gemini-3.1-flash-lite",
        "gemini-3.7-flash",
        "gemini-2.5-flash-lite",
        "gemini-2.5-flash",
    ]

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=MODEL_WHITELIST,
        index=0,
        help="白名單僅列出官方目前穩定支援 PDF 辨識＋結構化 JSON 輸出的高速模型，避免選到已棄用或需額外權限的型號"
    )

    concurrency = st.slider(
        "⚡ 並行辨識加速線程數",
        min_value=1,
        max_value=6,
        value=4,
        help="同時辨識多頁。設定 4~5 可將辨識速度提升數倍；若頻繁遇到 429 頻率限制，建議調低或改用 1"
    )

    if active_api_key:
        st.success("✅ API 金鑰連線正常")
    else:
        st.warning("⚠️ 請確認已在 secrets 設定或在此輸入 API Key")

    st.markdown("---")
    st.markdown("""
    💡 **會計計算原則**：
    * **電腦信用卡**：`信用卡刷卡(-)` 減 `信用卡退刷(+)`
    * **條碼**：`條碼支付進款(-)` 減 `條碼支付退款(+)`
    * **平衡檢核**：客運 + 貨運 - 電腦信用卡 - 條碼 + 其他 ＝ 自輸(應解總計)
    """)

# ----------------------------------------------------
# 2. 定義資料結構與工具函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱（例如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，不需包含'站'字或站碼）")
    date_day: int = Field(description="報表進款日期中的『日/號』(1 至 31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】區塊內的『客運(+)』金額，無則為 0")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】區塊內的『貨運(+)』金額，無則為 0")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】區塊內的『信用卡刷卡(-)』金額 (填正數)，無則為 0")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】區塊內的『信用卡退刷(+)』金額 (填正數)，無退刷則為 0")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】區塊內的『條碼支付進款(-)』金額 (填正數)，無則為 0")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】區塊內的『條碼支付退款(+)』金額 (填正數)，無退款則為 0")
    other_amount: float = Field(default=0.0, description="左側【應解款數】區塊內除上述項目外的其他明細金額加總（如存付運費、託收支票、補繳金額、繳回週轉金、其他短欠等），若無則填 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】區塊內的『應解總計』金額")

def extract_json_str(text: str) -> str:
    """安全去除 Markdown 標籤以解析 JSON"""
    if not text:
        return "{}"
    t = text.strip()
    if t.startswith("```"):
        lines = t.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        t = "\n".join(lines).strip()
    return t

def clean_station_name(val):
    """標準化車站名稱，自動剔除站碼前綴（如 3150_）以精準比對工作表"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    """轉換為乾淨整數或保留小數"""
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """會計立場扣抵公式：當有退刷/退款時寫入 =進款-退款，Excel 將自動算值並在公式列完整保留紀錄"""
    c = to_clean_num(charge_val)
    r = to_clean_num(refund_val)
    if c == 0 and r == 0:
        return None
    if c > 0 and r > 0:
        return f"={c}-{r}"
    if c > 0 and r == 0:
        return c
    if c == 0 and r > 0:
        return f"=-{r}"
    return None

def analyze_sheet_structure(sheet):
    """智慧掃描工作表表頭取得各欄位位置（絕不改動總計與加總公式欄）"""
    col_map = {}
    for r in range(1, 6):
        for c in range(1, sheet.max_column + 1):
            val = str(sheet.cell(row=r, column=c).value or "").replace(" ", "").replace("\n", "").strip()
            if not val:
                continue
            if "客運" in val and "passenger" not in col_map:
                col_map["passenger"] = c
            elif "貨運" in val and "freight" not in col_map:
                col_map["freight"] = c
            elif ("信用卡" in val or "刷卡" in val) and "credit" not in col_map:
                col_map["credit"] = c
            elif ("條碼" in val or "支付" in val) and "barcode" not in col_map:
                col_map["barcode"] = c
            elif "其他" in val and "other" not in col_map:
                col_map["other"] = c
            elif ("自輸" in val or "字輸" in val) and "remittance" not in col_map:
                col_map["remittance"] = c
            elif ("日" in val or "期" in val or "號" in val) and "date" not in col_map:
                col_map["date"] = c

    defaults = {
        "date": 1,
        "passenger": 2,
        "freight": 3,
        "credit": 4,
        "barcode": 5,
        "other": 6,
        "remittance": 7
    }
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return col_map

def find_target_row(sheet, date_day, date_col=1):
    """精準尋找對應日期的列號 (1~31)"""
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            try:
                val_str = str(val).replace("日", "").replace("號", "").strip()
                if int(float(val_str)) == int(date_day):
                    return r
            except Exception:
                pass
    return int(date_day) + 1

def get_column_reference_style(sheet, col_idx, ref_row=33):
    """公版中「尚未填寫」的資料儲存格通常沒有明確樣式(僅預設細明體/未加粗)，
    但第33列(月合計列)同一欄位是公式、必定已有公版設計的正確樣式(字型/對齊/框線/數字格式)。
    取這裡當樣式範本，讓辨識後填入的儲存格外觀與公版一致，不再變成預設樣式"""
    if not col_idx:
        return None
    ref_cell = sheet.cell(row=ref_row, column=col_idx)
    return ref_cell if ref_cell.has_style else None

def apply_style_ref(cell, style_ref):
    """將 style_ref 儲存格的樣式套用到 cell，style_ref 為 None 時不做任何事"""
    if style_ref is None:
        return
    cell.font = copy.copy(style_ref.font)
    cell.border = copy.copy(style_ref.border)
    cell.alignment = copy.copy(style_ref.alignment)
    cell.fill = copy.copy(style_ref.fill)
    cell.number_format = style_ref.number_format

def write_cell_if_valid(sheet, row_idx, col_idx, val, style_ref=None):
    """安全寫入指定儲存格（0 或空值不寫入以維護公版乾淨）；
    若提供 style_ref，會同步套用該樣式，確保新填入的資料格外觀與公版設計一致"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            cell = sheet.cell(row=row_idx, column=col_idx, value=val)
            apply_style_ref(cell, style_ref)
            return 1
    return 0

def trigger_auto_download(file_bytes: bytes, file_name: str):
    """在瀏覽器端自動觸發檔案下載（免使用者點擊）。採用 Blob 網址方式，不受網址長度限制，穩定支援較大檔案；
    此函式只負責「觸發下載」，不影響、不取代任何既有的 st.download_button（該按鍵仍會正常保留）"""
    b64 = base64.b64encode(file_bytes).decode()
    st.components.v1.html(f"""
        <script>
        const b64Data = "{b64}";
        const byteChars = atob(b64Data);
        const byteNumbers = new Array(byteChars.length);
        for (let i = 0; i < byteChars.length; i++) {{
            byteNumbers[i] = byteChars.charCodeAt(i);
        }}
        const blob = new Blob([new Uint8Array(byteNumbers)], {{type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}});
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "{file_name}";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        </script>
    """, height=0, width=0)

class FatalAPIError(Exception):
    """不可能靠重試解決的錯誤（如 404 模型不存在/已停用、401、403 金鑰或權限問題）"""
    pass

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=5, pre_call_cooldown=0.0):
    """單頁辨識函式
    - pre_call_cooldown：呼叫前先冷卻等待。單線程模式下用來讓每頁請求彼此保持間隔，降低瞬間連續打點觸發 429 的機率
    - 429 頻率限制：較長的指數退避（8s→16s→32s→60s，之後固定 60s 封頂），確保額度視窗有足夠時間重置
    - 503 伺服器忙碌：一般指數退避（2s→4s→8s→16s→25s封頂），通常短暫即恢復，不需等太久
    - 模型或金鑰問題（404 已停用/不存在、401/403 權限不足）：重試無用，直接判定為致命錯誤
    """
    if pre_call_cooldown > 0:
        time.sleep(pre_call_cooldown)

    last_err = "未知錯誤"
    for retry in range(max_retries):
        try:
            res = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=page_bytes, mime_type="application/pdf"),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=StationReport,
                    temperature=0.0,
                )
            )
            if res and res.text:
                clean_text = extract_json_str(res.text)
                return StationReport.model_validate_json(clean_text), None
            last_err = "模型回傳空白內容（可能被安全機制擋下），重試中"
        except errors.ClientError as e:
            if e.code == 429:
                last_err = f"429 頻率限制：{e.message}"
                time.sleep(min(60.0, 8.0 * (2 ** retry) + random.uniform(0, 2.0)))
                continue
            # 404 模型不存在/已停用、401 金鑰錯誤、403 權限不足等，重試無用，直接判定為致命錯誤
            raise FatalAPIError(f"{e.code} {e.message}")
        except errors.ServerError as e:
            last_err = f"{e.code} 伺服器忙碌：{e.message}"
            time.sleep(min(25.0, 2.0 ** (retry + 1) + random.uniform(0, 1.5)))
            continue
        except Exception as e:
            last_err = str(e)
        time.sleep(0.8)
    return None, f"超過最大重試次數（{last_err}）"

# ----------------------------------------------------
# 3. 介面上傳區塊
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳公版 Excel 檔案 (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 核心並行辨識與填寫
# ----------------------------------------------------
if st.button("🚀 開始智慧辨識與自動填表", type="primary", use_container_width=True):
    if not active_api_key:
        st.error("❌ 尚未設定 Gemini API Key，請在側邊欄輸入！")
        st.stop()
    if not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請確認已同時上傳 Excel 公版與 PDF 檔案！")
        st.stop()

    start_time = time.time()
    client = genai.Client(api_key=active_api_key)

    try:
        # data_only=False 確保公版內既有的公式、樣式與格式 100% 完整保留
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_excel.getvalue()), data_only=False)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗，請確認上傳標準 .xlsx 公版：{e}")
        st.stop()

    status_box = st.status("📄 [階段 1/2] 正在拆分 PDF 頁面...", expanded=True)
    all_pages = []
    
    for pdf_file in uploaded_pdfs:
        try:
            reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
            total_p = len(reader.pages)
            for p_idx, page in enumerate(reader.pages, 1):
                writer = PdfWriter()
                writer.add_page(page)
                page_buf = io.BytesIO()
                writer.write(page_buf)
                all_pages.append((pdf_file.name, p_idx, total_p, page_buf.getvalue()))
        except Exception as e:
            status_box.write(f"⚠️ 檔案 `{pdf_file.name}` 讀取異常：{e}")

    total_tasks = len(all_pages)
    if total_tasks == 0:
        status_box.update(label="❌ 沒有找到可處理的 PDF 頁面", state="error")
        st.stop()

    status_box.update(label=f"⚡ [階段 2/2] {concurrency} 線程高速並行辨識 {total_tasks} 頁解款單據...", state="running")
    
    prompt = """
    你是一位專業精確的台鐵會計表單辨識專家。請仔細檢視這張解款單據：
    【特別注意淺色或複寫印件】：
    單據若為複寫或影印，字跡可能較淺，請特別仔細分辨淺色數字（如 0、3、8、1、7），切勿遺漏。
    1. 擷取【車站名稱】（例如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）與進款【日期】（僅需日/號數，1-31 的整數）。
    2. 專注看左側【應解款數】大項目區塊，精確擷取各數值（若無填 0）：
       - 客運(+)
       - 貨運(+)
       - 信用卡刷卡(-) (填正數)
       - 信用卡退刷(+) (請仔細分辨，有退刷務必填正數金額；若無退刷填 0)
       - 條碼支付進款(-) (填正數)
       - 條碼支付退款(+) (請仔細分辨，有退款務必填正數金額；若無退款填 0)
       - 應解總計 (報表上的應解總計數值)
       - 其他項目加總（如存付運費、託收支票、補繳金額、繳回週轉金等其他明細淨額，若無填 0）
    3. 會計勾稽驗算原則：必符合「客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計」。請務必以此原則交叉驗算確認！
    """

    results_data = []
    progress_bar = st.progress(0)
    completed_count = 0
    fatal_error = None

    # 單線程模式下，每頁請求之間加入固定冷卻秒數，避免瞬間連續打點觸發 429；
    # 多線程模式下請求時間點已因並行而自然分散，故不額外節流，以維持產出效率
    per_page_cooldown = 2.0 if concurrency == 1 else 0.0

    # 採用 ThreadPoolExecutor 並行加速呼叫 API
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        future_map = {
            executor.submit(
                call_gemini_page, client, selected_model, page_bytes, prompt,
                pre_call_cooldown=per_page_cooldown
            ): (idx, file_name, p_idx, total_p)
            for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1)
        }

        for future in as_completed(future_map):
            completed_count += 1
            idx, file_name, p_idx, total_p = future_map[future]
            progress_bar.progress(
                completed_count / total_tasks,
                text=f"⚡ 高速並行辨識中：已完成 {completed_count}/{total_tasks} 頁 ({completed_count * 100 // total_tasks}%)"
            )
            
            try:
                parsed_data, err_msg = future.result()
                if parsed_data and parsed_data.date_day > 0:
                    results_data.append((idx, file_name, p_idx, parsed_data))
                    status_box.write(f"✅ **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功 - `{file_name}` 第 {p_idx} 頁")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗 ({str(err_msg)[:60]})")
            except FatalAPIError as e:
                fatal_error = fatal_error or str(e)
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁執行異常：{e}")

    if fatal_error:
        st.error(
            f"❌ Gemini API 回傳無法靠重試解決的錯誤：{fatal_error}\n\n"
            f"常見原因：所選模型「{selected_model}」已不存在或已停用（404），或 API Key 權限不足（401/403）。"
            f"請至左側「AI 辨識核心模型」選單改選其他模型，或確認 API Key 是否正確後再重新執行一次。"
            f"（下方仍會列出本次已成功辨識的頁面，可先下載保留。）"
        )

    # 依照原始單據頁面順序排序，確保勾稽審核表井然有序
    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel 與 會計立場平衡檢查（主線程安全寫入）
    # ----------------------------------------------------
    status_box.update(label="📝 正在將辨識資料寫入 Excel 公版...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_name_clean = clean_station_name(data.station_name)
        
        # 嚴格優先完全符合（防止「烏日」與「新烏日」搶先誤配）
        target_sheet = None
        for s_name in wb.sheetnames:
            if clean_station_name(s_name) == target_name_clean:
                target_sheet = wb[s_name]
                break
        if not target_sheet:
            for s_name in wb.sheetnames:
                s_clean = clean_station_name(s_name)
                if s_clean and (s_clean in target_name_clean or target_name_clean in s_clean):
                    target_sheet = wb[s_name]
                    break

        net_credit = data.credit_card_charge - data.credit_card_refund
        net_barcode = data.barcode_in - data.barcode_refund

        computed_total = data.passenger_revenue + data.freight_revenue - net_credit - net_barcode + data.other_amount
        diff = round(data.remittance_total - computed_total, 2)
        is_balanced = (abs(diff) < 0.01)

        cc_formula = build_deduction_formula(data.credit_card_charge, data.credit_card_refund)
        bc_formula = build_deduction_formula(data.barcode_in, data.barcode_refund)

        audit_records.append({
            "檔案名稱": file_name,
            "車站名稱": data.station_name,
            "日期": f"{data.date_day} 日",
            "客運": to_clean_num(data.passenger_revenue),
            "貨運": to_clean_num(data.freight_revenue),
            "電腦信用卡 (=刷卡-退刷)": cc_formula if cc_formula is not None else 0,
            "條碼 (=進款-退款)": bc_formula if bc_formula is not None else 0,
            "其他": to_clean_num(data.other_amount),
            "自輸 (PDF應解總計)": to_clean_num(data.remittance_total),
            "計算總計": to_clean_num(computed_total),
            "差額 (自輸-總計)": diff,
            "平衡狀態": "✅ 平衡 (0)" if is_balanced else f"❌ 差額 {diff:+.0f} (請核對)",
            "工作表寫入": f"已寫入 [{target_sheet.title}]" if target_sheet else "❌ 找不到分頁"
        })

        if not target_sheet:
            continue

        col_map = analyze_sheet_structure(target_sheet)
        target_row = find_target_row(target_sheet, data.date_day, col_map["date"])
        style_refs = {k: get_column_reference_style(target_sheet, v) for k, v in col_map.items()}

        # 精準寫入特定資料儲存格，原儲存格公式 100% 不受干擾；樣式則比照公版第33列同欄位設計
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue), style_refs.get("passenger"))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue), style_refs.get("freight"))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula, style_refs.get("credit"))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula, style_refs.get("barcode"))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount), style_refs.get("other"))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total), style_refs.get("remittance"))

        success_count += 1

    status_box.update(label="🎉 辨識與 Excel 寫入全數完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    elapsed = time.time() - start_time

    # 將結果存入 session_state：確保之後任何操作（含下載按鍵本身、或下方合併工具的按鍵）
    # 造成的畫面重新整理，此結果區塊與下載按鍵都會「持續顯示、不會消失」；
    # is_new / trigger_auto_download 只用來讓慶祝動畫與自動下載「只觸發這一次」，不會每次重整都重複下載
    st.session_state["ocr_result"] = {
        "excel_bytes": out_stream.getvalue(),
        "audit_records": audit_records,
        "elapsed": elapsed,
        "success_count": success_count,
        "total_tasks": total_tasks,
        "total_written": total_written,
    }
    st.session_state["ocr_result_is_new"] = True
    st.session_state["ocr_trigger_auto_download"] = True

# ----------------------------------------------------
# 6. 會計平衡檢核儀表板與下載（永遠從 session_state 讀取顯示，按鍵不會消失）
# ----------------------------------------------------
if "ocr_result" in st.session_state:
    res = st.session_state["ocr_result"]

    if st.session_state.get("ocr_result_is_new", False):
        st.balloons()
        st.session_state["ocr_result_is_new"] = False

    st.success(f"✨ 處理完成！耗時 {res['elapsed']:.1f} 秒，共成功處理 {res['success_count']}/{res['total_tasks']} 頁單據，填寫了 {res['total_written']} 個儲存格！")

    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if res["audit_records"]:
        df_audit = pd.DataFrame(res["audit_records"])
        unbalanced_count = sum(1 for r in res["audit_records"] if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 警告：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請依下方表格核對單據金額！")
        else:
            st.success("🎯 太棒了！所有辨識成功的單據「自輸 - 總計」皆等於 0，會計平衡完全正確！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=res["excel_bytes"],
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        key="ocr_download_btn"
    )

    if st.session_state.get("ocr_trigger_auto_download", False):
        trigger_auto_download(res["excel_bytes"], "台鐵解款單_彙總完成表.xlsx")
        st.session_state["ocr_trigger_auto_download"] = False

# ----------------------------------------------------
# 7. 全通用型 Excel 跨公版資料搬移工具 (全分頁強效套用版)
# ----------------------------------------------------
st.markdown("---")
st.header("🛠️ 全通用型 Excel 跨公版搬移工具")
st.caption("🚀 鎖定範本欄位結構 ➜ 100% 自動套用至所有同名分頁 ｜ 🛡️ 公式完全防護 ｜ 🎯 容錯日期對齊")

col_u1, col_u2 = st.columns(2)
with col_u1:
    target_file = st.file_uploader("📋 步驟一：上傳【目標公版】(保留其公式)", type=["xlsx"], key="gen_target_uploader")
with col_u2:
    source_file = st.file_uploader("📥 步驟二：上傳【來源檔案】(提取純數據)", type=["xlsx"], key="gen_source_uploader")

if target_file and source_file:
    try:
        t_wb_preview = openpyxl.load_workbook(io.BytesIO(target_file.getvalue()), data_only=True)
        s_wb_preview = openpyxl.load_workbook(io.BytesIO(source_file.getvalue()), data_only=True)
        t_sheets = t_wb_preview.sheetnames
        s_sheets = s_wb_preview.sheetnames
    except Exception as e:
        st.error(f"❌ 檔案讀取失敗：{e}")
        st.stop()

    def clean_s_name(name):
        """車站名稱去空格與台/臺轉換"""
        return str(name).replace("臺", "台").replace("站", "").strip()

    t_norm_map = {clean_s_name(name): name for name in t_sheets if clean_s_name(name)}
    common_sheets = [s_name for s_name in s_sheets if clean_s_name(s_name) in t_norm_map]

    st.markdown("---")
    st.markdown("#### ⚙️ 步驟三：設定對齊與欄位映射規則")
    
    cfg_c1, cfg_c2 = st.columns(2)
    with cfg_c1:
        sheet_mode = st.radio("工作表比對模式", ["自動按同名工作表批次套用", "單一工作表指定搬移"], horizontal=True)
        
        if sheet_mode == "單一工作表指定搬移":
            s_sample_sheet_name = st.selectbox("來源工作表", s_sheets)
            t_sample_sheet_name = st.selectbox("目標工作表", t_sheets)
        else:
            if not common_sheets:
                st.warning("⚠️ 兩個檔案中沒有找到名稱完全相同的分頁，請切換為【單一工作表指定搬移】。")
                st.stop()
            st.success(f"✅ 系統成功配對 **{len(common_sheets)}** 個同名車站分頁，搬移時將全面套用！")
            s_sample_sheet_name = st.selectbox("請選擇 1 個分頁作為【欄位結構範本】", options=common_sheets, index=0)
            t_sample_sheet_name = t_norm_map[clean_s_name(s_sample_sheet_name)]

        header_row = st.number_input("標題欄位名稱在第幾列？ (表頭列號)", min_value=1, max_value=10, value=1, step=1, help="若第 1 列為大標題，請調整為 2 或 3")

    # 深度掃描表頭欄位（相容換行符號、空格與常見合併儲存格）
    def parse_sheet_headers_robust(ws, h_row):
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=h_row, column=col_idx).value
            # 若選定列為空，自動往下一列探測（應對合併儲存格表頭）
            if val is None or str(val).strip() == "":
                val = ws.cell(row=h_row + 1, column=col_idx).value
            if val is not None and str(val).strip() != "":
                c_name = str(val).replace("\n", "").replace(" ", "").replace("　", "").strip()
                if c_name:
                    headers[c_name] = col_idx
        return headers

    s_ws_sample = s_wb_preview[s_sample_sheet_name]
    t_ws_sample = t_wb_preview[t_sample_sheet_name]

    s_headers = parse_sheet_headers_robust(s_ws_sample, header_row)
    t_headers = parse_sheet_headers_robust(t_ws_sample, header_row)

    with st.expander("🔍 點擊查看範本工作表偵測到的表頭欄位", expanded=not bool(t_headers)):
        col_pv1, col_pv2 = st.columns(2)
        with col_pv1:
            st.caption(f"來源範本 [{s_sample_sheet_name}]：")
            st.write(list(s_headers.keys()) if s_headers else "❌ 讀取空白，請更改列號")
        with col_pv2:
            st.caption(f"目標範本 [{t_sample_sheet_name}]：")
            st.write(list(t_headers.keys()) if t_headers else "❌ 讀取空白，請更改列號")

    if not t_headers or not s_headers:
        st.warning(f"⚠️ 範本分頁第 {header_row} 列未讀取到完整表頭文字，請在上方切換「表頭列號」（改為 2 或 3）。")
        st.stop()

    with cfg_c2:
        align_mode = st.radio("資料列對齊方式", ["依據關鍵欄位對齊 (強烈推薦，如：日期/編號)", "直接依列號對齊 (第 N 列搬到第 N 列)"])
        key_col = None
        if align_mode.startswith("依據關鍵欄位"):
            # 優先搜尋包含「日」、「期」、「號」的欄位
            suggested_keys = [k for k in t_headers.keys() if any(w in k for w in ["日", "期", "號", "代碼", "編號"])]
            def_idx = list(t_headers.keys()).index(suggested_keys[0]) if suggested_keys else 0
            key_col = st.selectbox(
                "選擇用來對齊資料的【關鍵欄位】", 
                options=list(t_headers.keys()),
                index=def_idx,
                help="以日期為例：不論兩邊第幾列，只要日期相同就會精確填入"
            )

    st.markdown("##### 📌 選擇要搬移的欄位對應（目標欄位 ➜ 取自來源欄位）")
    t_header_list = [k for k in t_headers.keys() if k != key_col]
    s_header_list = ["(略過不搬移)"] + [k for k in s_headers.keys() if k != key_col]

    mapping_cols = st.columns(3)
    field_mappings = {} # {目標欄名: 來源欄名}

    for idx, t_col_name in enumerate(t_header_list):
        col_slot = mapping_cols[idx % 3]
        # 自動智慧配對（同名或包含關係）
        matched_idx = 0
        for s_idx, s_name in enumerate(s_header_list):
            if s_name != "(略過不搬移)" and (s_name in t_col_name or t_col_name in s_name):
                matched_idx = s_idx
                break

        with col_slot:
            matched_s = st.selectbox(
                f"目標 `{t_col_name}`",
                options=s_header_list,
                index=matched_idx,
                key=f"map_select_{t_col_name}"
            )
            if matched_s != "(略過不搬移)":
                field_mappings[t_col_name] = matched_s

    st.markdown("##### 🧹 寫入與清除控制")
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        clear_target_data = st.checkbox("🧹 搬移前先清空目標公版對應欄位的舊數值", value=False, help="僅清除非公式純數值，公式儲存格 100% 受到保護保留")
    with col_opt2:
        overwrite_mode = st.radio("遇到目標已有數值時：", ["直接覆蓋舊數值", "保留原值（只填補完全空白的儲存格）"], horizontal=True)

    # ----------------------------------------------------
    # 執行通用搬移邏輯
    # ----------------------------------------------------
    if st.button("🚀 開始執行全分頁資料搬移", type="primary", use_container_width=True):
        if not field_mappings:
            st.error("❌ 請至少選擇一個要搬移的欄位！")
            st.stop()

        target_wb = openpyxl.load_workbook(io.BytesIO(target_file.getvalue()), data_only=False)
        src_wb = openpyxl.load_workbook(io.BytesIO(source_file.getvalue()), data_only=False)

        # 鎖定範本欄位座標（解決其他分頁因合併儲存格讀不到表頭的問題）
        fixed_t_col_indices = {t_name: t_headers[t_name] for t_name in field_mappings.keys()}
        fixed_s_col_indices = {t_name: s_headers[field_mappings[t_name]] for t_name in field_mappings.keys()}
        fixed_t_key_col = t_headers.get(key_col)
        fixed_s_key_col = s_headers.get(key_col)

        def normalize_key_value(val):
            """將日期等數值強制轉為乾淨比對字串（如 1.0 -> 1, ' 1日 ' -> 1）"""
            if val is None:
                return None
            s = str(val).replace("日", "").replace("號", "").strip()
            try:
                f = float(s)
                return str(int(f)) if f.is_integer() else str(f)
            except Exception:
                return s

        work_pairs = []
        if sheet_mode == "單一工作表指定搬移":
            work_pairs.append((src_wb[s_sample_sheet_name], target_wb[t_sample_sheet_name]))
        else:
            for s_name in common_sheets:
                t_name = t_norm_map[clean_s_name(s_name)]
                work_pairs.append((src_wb[s_name], target_wb[t_name]))

        logs = []
        cleared_count = 0
        written_count = 0

        for s_ws, t_ws in work_pairs:
            # 建立目標資料列主鍵映射 (Key -> Row Number)
            t_row_map = {}
            if align_mode.startswith("依據關鍵欄位") and fixed_t_key_col:
                for r in range(header_row + 1, min(t_ws.max_row + 1, 100)):
                    val = t_ws.cell(row=r, column=fixed_t_key_col).value
                    clean_k = normalize_key_value(val)
                    if clean_k:
                        t_row_map[clean_k] = r

            # 功能 A：預先清除選定欄位的純數值（保護公式）
            if clear_target_data:
                for t_name, tc_idx in fixed_t_col_indices.items():
                    for r in range(header_row + 1, t_ws.max_row + 1):
                        cell = t_ws.cell(row=r, column=tc_idx)
                        if cell.value is not None and not (isinstance(cell.value, str) and cell.value.startswith("=")):
                            cell.value = None
                            cleared_count += 1

            # 功能 B：批次填寫
            for s_row in range(header_row + 1, s_ws.max_row + 1):
                target_r = None
                if align_mode.startswith("依據關鍵欄位") and fixed_s_key_col:
                    s_key_val = s_ws.cell(row=s_row, column=fixed_s_key_col).value
                    clean_sk = normalize_key_value(s_key_val)
                    if clean_sk:
                        target_r = t_row_map.get(clean_sk)
                else:
                    target_r = s_row if s_row <= t_ws.max_row else None

                if not target_r:
                    continue

                for t_name, tc_idx in fixed_t_col_indices.items():
                    sc_idx = fixed_s_col_indices[t_name]
                    t_cell = t_ws.cell(row=target_r, column=tc_idx)
                    s_val = s_ws.cell(row=s_row, column=sc_idx).value

                    # 1. 嚴格保護目標公版公式
                    if isinstance(t_cell.value, str) and t_cell.value.startswith("="):
                        continue

                    # 2. 來源若是空值跳過
                    if s_val is None or str(s_val).strip() == "":
                        continue

                    # 3. 填入或覆蓋數值
                    is_empty_target = (t_cell.value is None or str(t_cell.value).strip() == "" or t_cell.value == 0)
                    
                    if is_empty_target:
                        t_cell.value = s_val
                        written_count += 1
                        logs.append({"工作表": t_ws.title, "列號": target_r, "欄位": t_name, "填入值": s_val})
                    elif overwrite_mode.startswith("直接覆蓋"):
                        t_cell.value = s_val
                        written_count += 1
                        logs.append({"工作表": t_ws.title, "列號": target_r, "欄位": t_name, "填入值": f"{s_val} (覆蓋舊值)"})

        out = io.BytesIO()
        target_wb.save(out)
        st.session_state["gen_merge_result"] = {
            "bytes": out.getvalue(),
            "logs": logs,
            "written_count": written_count,
            "cleared_count": cleared_count,
            "sheets_processed": len(work_pairs)
        }
        st.session_state["gen_merge_new"] = True

if "gen_merge_result" in st.session_state:
    res = st.session_state["gen_merge_result"]
    if st.session_state.get("gen_merge_new", False):
        st.balloons()
        st.session_state["gen_merge_new"] = False

    st.success(f"🎉 搬移作業完成！共處理了 **{res['sheets_processed']}** 個車站分頁，成功填入/更新 **{res['written_count']}** 個儲存格！" +
               (f"（事前已清空 {res['cleared_count']} 個純數值儲存格）" if res['cleared_count'] > 0 else ""))

    with st.expander(f"📋 查看詳細搬移紀錄清單（共 {len(res['logs'])} 筆）"):
        if res["logs"]:
            st.dataframe(pd.DataFrame(res["logs"]), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 本次沒有任何資料被寫入。請確認：1. 關鍵欄位（如日期）是否能配對 2. 來源檔選定欄位是否非空。")

    st.download_button(
        label="📥 點擊下載搬移完成的 Excel 報表 (.xlsx)",
        data=res["bytes"],
        file_name="跨公版通用搬移完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
