import io
import re
import time
import openpyxl
import pandas as pd
from pydantic import BaseModel, Field
from pypdf import PdfReader, PdfWriter
import streamlit as st
from google import genai
from google.genai import types, errors

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 單線程極速辨識 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🛡️ 防429頻率限制")

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
    
    EXCLUDE_KEYWORDS = ("embedding", "image", "audio", "tts", "live", "aqa", "vision", "transcribe", "robotics")

    def _model_rank(name):
        n = name.lower()
        tier = 0 if ("flash" in n and "lite" not in n) else (1 if "flash" in n else (2 if "pro" in n else 3))
        unstable = 1 if ("preview" in n or "exp" in n) else 0
        nums = re.findall(r'\d+(?:\.\d+)?', n)
        version = float(nums[0]) if nums else 0.0
        return (tier, unstable, -version)

    available_models = []
    if active_api_key:
        try:
            temp_client = genai.Client(api_key=active_api_key)
            for m in temp_client.models.list():
                m_name = getattr(m, "name", "").replace("models/", "")
                name_l = m_name.lower()
                actions = getattr(m, "supported_actions", None)
                if "gemini" not in name_l or (actions and "generateContent" not in actions):
                    continue
                if any(k in name_l for k in EXCLUDE_KEYWORDS):
                    continue
                available_models.append(m_name)
        except Exception:
            pass

    if not available_models:
        available_models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    else:
        available_models = sorted(set(available_models), key=_model_rank)

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=available_models,
        index=0,
        help="系統已自動篩選出可用模型，預設推薦高速 Flash 模型"
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
    station_name: str = Field(description="車站名稱（如豐富、苗栗、銅鑼、三義、后里、豐原、台中、彰化、員林等，不需站字或站碼）")
    date_day: int = Field(default=0, description="報表進款日期中的『日/號』(1 至 31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="客運(+)金額，無則為 0")
    freight_revenue: float = Field(default=0.0, description="貨運(+)金額，無則為 0")
    credit_card_charge: float = Field(default=0.0, description="信用卡刷卡(-)金額，填正數，無則為 0")
    credit_card_refund: float = Field(default=0.0, description="信用卡退刷(+)金額，填正數，無退刷則為 0")
    barcode_in: float = Field(default=0.0, description="條碼支付進款(-)金額，填正數，無則為 0")
    barcode_refund: float = Field(default=0.0, description="條碼支付退款(+)金額，填正數，無退款則為 0")
    other_amount: float = Field(default=0.0, description="其他項目淨額加總，無則為 0")
    remittance_total: float = Field(default=0.0, description="應解總計金額")

def extract_json_str(text: str) -> str:
    """去除 Markdown 標記以安全解析 JSON"""
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
    """標準化車站名稱，剔除站碼與雜字"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("火車站", "").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """有退刷/退款時寫入 =進款-退款，無退款則直接填數值"""
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
    """掃描工作表表頭取得欄位座標"""
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
            elif ("自輸" in val or "字輸" in val or "應解" in val or "總計" in val) and "remittance" not in col_map:
                col_map["remittance"] = c
            elif ("日" in val or "期" in val or "號" in val) and "date" not in col_map:
                col_map["date"] = c

    defaults = {"date": 1, "passenger": 2, "freight": 3, "credit": 4, "barcode": 5, "other": 6, "remittance": 7}
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return col_map

def find_target_row(sheet, date_day, date_col=1):
    """定位對應日期的列號 (1~31)"""
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            try:
                val_str = str(val).replace("日", "").replace("號", "").strip()
                if int(float(val_str)) == int(date_day):
                    return r
            except Exception:
                pass
    # 備援：若指定欄位未對到，向前 3 欄進行全面搜尋
    for c in range(1, 4):
        for r in range(1, 45):
            val = sheet.cell(row=r, column=c).value
            if val is not None:
                try:
                    val_str = str(val).replace("日", "").replace("號", "").strip()
                    if int(float(val_str)) == int(date_day):
                        return r
                except Exception:
                    pass
    return int(date_day) + 1

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    """安全寫入非空數值，不覆蓋空白格式"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page_single(client, model_name, page_bytes, prompt, status_box, max_retries=5):
    """單線程辨識函式：遇到 429 採用階梯冷卻等待，保證成功重試"""
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
        except errors.ClientError as e:
            if e.code == 429:
                wait_sec = 8 * (retry + 1)
                status_box.write(f"⏳ 遭遇 API 頻率限制 (429)，自動冷卻 {wait_sec} 秒後重試 (第 {retry+1}/{max_retries} 次)...")
                time.sleep(wait_sec)
                continue
            return None, f"API 錯誤 ({e.code})：{e.message}"
        except errors.ServerError as e:
            wait_sec = 5 * (retry + 1)
            status_box.write(f"⚠️ 伺服器忙碌 ({e.code})，等待 {wait_sec} 秒後重試...")
            time.sleep(wait_sec)
            continue
        except Exception as e:
            time.sleep(2.0)
            if retry == max_retries - 1:
                return None, str(e)
    return None, "超過最大重試次數，請稍候再試"

# ----------------------------------------------------
# 3. 介面上傳區塊
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳公版 Excel 檔案 (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 單線程高速辨識與即時填寫
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
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_excel.getvalue()), data_only=False)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗，請確認上傳標準 .xlsx 檔案：{e}")
        st.stop()

    status_box = st.status("📄 [階段 1/2] 正在拆分 PDF 單據頁面...", expanded=True)
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

    status_box.update(label=f"⚡ [階段 2/2] 單線程高速逐頁辨識中 (共 {total_tasks} 頁)...", state="running")
    
    prompt = """
    你是一位專業精確的台鐵會計表單辨識專家。請仔細檢視這張解款單據：
    【特別注意淺色或複寫印件】：
    單據若為複寫或影印，字跡可能較淺，請特別仔細分辨淺色數字（如 0、3、8、1、7），切勿遺漏。
    1. 擷取【車站名稱】（如豐富、苗栗、銅鑼、三義、后里、豐原、台中、彰化、員林等，去除站碼與'站'字）與進款【日期】（僅需日/號數，1-31 的整數）。
    2. 專注看左側【應解款數】大項目區塊，精確擷取各數值（若無填 0）：
       - 客運(+)
       - 貨運(+)
       - 信用卡刷卡(-) (填正數)
       - 信用卡退刷(+) (若無退刷填 0)
       - 條碼支付進款(-) (填正數)
       - 條碼支付退款(+) (若無退款填 0)
       - 應解總計 (報表上的應解總計數值)
       - 其他項目加總（如存付運費、託收支票、補繳金額、繳回週轉金等其他明細淨額，若無填 0）
    3. 會計勾稽驗算原則：必符合「客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計」。請務必以此原則交叉驗算確認！
    """

    progress_bar = st.progress(0)
    total_written = 0
    success_count = 0
    audit_records = []

    # 單線程逐頁執行，每頁辨識完立即填寫
    for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1):
        progress_bar.progress(
            (idx - 1) / total_tasks,
            text=f"🔍 正在辨識第 {idx}/{total_tasks} 頁：`{file_name}` (第 {p_idx} 頁)..."
        )
        
        data, err_msg = call_gemini_page_single(client, selected_model, page_bytes, prompt, status_box)

        if data and data.date_day > 0:
            target_name_clean = clean_station_name(data.station_name)
            
            # 優先完全符合工作表名稱
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

            if target_sheet:
                col_map = analyze_sheet_structure(target_sheet)
                target_row = find_target_row(target_sheet, data.date_day, col_map["date"])

                # 精準回填指定儲存格
                page_cells = 0
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue))
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue))
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
                page_cells += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))

                total_written += page_cells
                success_count += 1
                status_box.write(f"✅ 第 {idx}/{total_tasks} 頁：**{data.station_name}**（{data.date_day} 日）辨識成功，填寫 {page_cells} 格 ➜ `{target_sheet.title}`")
            else:
                status_box.write(f"⚠️ 第 {idx}/{total_tasks} 頁：**{data.station_name}** 辨識成功，但在 Excel 中找不到對應車站工作表")
        else:
            status_box.write(f"❌ 第 {idx}/{total_tasks} 頁辨識失敗（`{file_name}` 第 {p_idx} 頁）：{err_msg}")

        # 頁與頁之間安全微幅節流，維持在安全請求速率之內
        if idx < total_tasks:
            time.sleep(2.0)

    progress_bar.progress(1.0, text="🎉 所有頁面處理完成！")
    status_box.update(label="🎉 辨識與 Excel 寫入全數完成！", state="complete")

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，共成功處理 {success_count}/{total_tasks} 頁單據，填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 5. 會計平衡檢核儀表板與下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 警告：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請依下方表格核對單據金額！")
        else:
            st.success("🎯 太棒了！所有辨識成功的單據「自輸 - 總計」皆等於 0，會計平衡完全正確！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
