import io
import os
import time
import random
import openpyxl
import pandas as pd
from PIL import Image
from pydantic import BaseModel, Field
from pypdf import PdfReader, PdfWriter
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("⚡ 多檔批次極速版 ｜ 🔍 高清原生影像直出 ｜ 🧮 扣抵公式保留 ｜ 🎯 零死角跨日彙總")

raw_key = st.secrets.get("GEMINI_API_KEY", "")
cleaned_key = str(raw_key).replace('"', '').replace("'", "").strip()

with st.sidebar:
    st.header("⚙️ 系統效能設定")
    user_key = st.text_input(
        "Gemini API Key",
        value=cleaned_key,
        type="password",
        help="系統會優先讀取 secrets 中的 GEMINI_API_KEY"
    )
    active_api_key = user_key.strip()
    
    # 僅提供官方已上線之標準模型
    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=["gemini-2.0-flash", "gemini-1.5-flash"],
        index=0,
        help="推薦 gemini-2.0-flash，辨識速度最快且對表格與淺色字跡判讀力最強"
    )

    max_workers = st.slider(
        "並行執行緒數",
        min_value=1,
        max_value=4,
        value=2,
        help="建議設為 2，既能大幅加速又能穩定防止 Google 頻率限制 (429)"
    )

    if active_api_key:
        st.success("✅ API 金鑰已就緒")
    else:
        st.warning("⚠️ 請確認已在 secrets 設定或在此輸入 API Key")

    st.markdown("---")
    st.markdown("""
    💡 **會計計算原則**：
    * **電腦信用卡**：`信用卡刷卡(-)` 減 `信用卡退刷(+)`
    * **條碼**：`條碼支付進款(-)` 減 `條碼支付退款(+)`
    * **平衡驗證**：客運 + 貨運 - 電腦信用卡 - 條碼 + 其他 ＝ 自輸(應解總計)
    """)

# ----------------------------------------------------
# 2. 定義資料結構與輔助函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱（例如：埔心、楊梅、富岡、北湖、湖口、新豐、竹北、北新竹、新竹、香山、三姓橋、千甲、新莊、竹中、六家、竹東、內灣、竹南、大山、後龍、白沙屯、新埔、通霄、苑裡、日南、大甲、台中港、清水、沙鹿、龍井、大肚、追分等，不含'站'字）")
    date_day: int = Field(description="報表日期中的『日/號』(1 至 31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】內的『客運(+)』金額，無則為 0")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】內的『貨運(+)』金額，無則為 0")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】內的『信用卡刷卡(-)』金額 (填正數)，無則為 0")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】內的『信用卡退刷(+)』金額 (填正數)，無則為 0")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】內的『條碼支付進款(-)』金額 (填正數)，無則為 0")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】內的『條碼支付退款(+)』金額 (填正數)，無則為 0")
    other_amount: float = Field(default=0.0, description="左側【應解款數】內除上述外的獨立明細淨額（如存付運費、託收支票、短欠、補繳等），若無則填 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】內的『應解總計』金額")

def extract_json_str(text: str) -> str:
    """去除 Markdown 標籤以安全解析 JSON"""
    if not text:
        return "{}"
    t = text.strip()
    if "```json" in t:
        t = t.split("```json")[1].split("```")[0].strip()
    elif "```" in t:
        t = t.split("```")[1].split("```")[0].strip()
    start = t.find("{")
    end = t.rfind("}")
    if start != -1 and end != -1 and end > start:
        t = t[start:end+1]
    return t

def clean_station_name(val):
    """標準化車站名稱"""
    if not val:
        return ""
    return str(val).replace("臺", "台").replace("站", "").replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    """將數值轉為整數或小數"""
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """生成會計相減公式"""
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

def find_matching_sheet(wb, station_name):
    """精準搜尋符合的車站分頁（優先完全精確相符，避免新竹被北新竹覆蓋）"""
    clean_target = clean_station_name(station_name)
    if not clean_target:
        return None
    # 步驟 1：完全精確相符
    for s_name in wb.sheetnames:
        if clean_station_name(s_name) == clean_target:
            return wb[s_name]
    # 步驟 2：模糊包含比對
    candidates = []
    for s_name in wb.sheetnames:
        s_clean = clean_station_name(s_name)
        if s_clean and (s_clean in clean_target or clean_target in s_clean):
            candidates.append(s_name)
    if candidates:
        candidates.sort(key=lambda x: abs(len(clean_station_name(x)) - len(clean_target)))
        return wb[candidates[0]]
    return None

def extract_page_image_payload(page):
    """從 PDF 頁面抽取高解析度 JPEG 原圖（純 Python，免 Poppler）"""
    if len(page.images) > 0:
        raw_bytes = page.images[0].data
        try:
            pil_img = Image.open(io.BytesIO(raw_bytes))
            # 若圖檔尺寸超大，微幅縮小至 1800px 以極速傳輸，同時保持印刷筆劃極致清晰
            if max(pil_img.size) > 1800:
                ratio = 1800 / max(pil_img.size)
                new_size = (int(pil_img.size[0] * ratio), int(pil_img.size[1] * ratio))
                resized = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                resized.save(buf, format="JPEG", quality=85)
                return ("image/jpeg", buf.getvalue())
            return ("image/jpeg", raw_bytes)
        except Exception:
            return ("image/jpeg", raw_bytes)
    else:
        # 若為文字型 PDF 則直接輸出單頁 PDF 串流
        writer = PdfWriter()
        writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        return ("application/pdf", buf.getvalue())

def load_excel_preserving_all(file_bytes, filename):
    """載入 Excel 檔案並完整保留公式結構"""
    if filename.lower().endswith(".xls"):
        import xlrd
        xls_book = xlrd.open_workbook(file_contents=file_bytes, formatting_info=False)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        for sheet_name in xls_book.sheet_names():
            xls_sheet = xls_book.sheet_by_name(sheet_name)
            xlsx_sheet = wb.create_sheet(title=sheet_name)
            for r in range(xls_sheet.nrows):
                for c in range(xls_sheet.ncols):
                    cell = xls_sheet.cell(r, c)
                    val = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            val = xlrd.xldate.xldate_as_datetime(val, xls_book.datemode)
                        except Exception:
                            pass
                    if val != "" and val is not None:
                        xlsx_sheet.cell(row=r + 1, column=c + 1, value=val)
            
            header_row = [str(xlsx_sheet.cell(1, c).value or "") for c in range(1, 10)]
            if any("客運" in h for h in header_row):
                for day_r in range(2, 33):
                    xlsx_sheet.cell(row=day_r, column=8, value=f"=B{day_r}+C{day_r}-D{day_r}-E{day_r}+F{day_r}")
                    xlsx_sheet.cell(row=day_r, column=9, value=f"=G{day_r}-H{day_r}")
                xlsx_sheet.cell(row=33, column=1, value="合計")
                for col_i in range(2, 10):
                    col_letter = openpyxl.utils.get_column_letter(col_i)
                    xlsx_sheet.cell(row=33, column=col_i, value=f"=SUM({col_letter}2:{col_letter}32)")
        return wb
    else:
        return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)

def analyze_sheet_structure(sheet):
    """分析工作表欄位位置"""
    col_map = {}
    for r in range(1, 8):
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
    """搜尋日期所在列數"""
    for r in range(1, 50):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            try:
                val_str = str(val).replace("日", "").replace("號", "").strip()
                if int(float(val_str)) == int(date_day):
                    return r
            except Exception:
                pass
    return int(date_day) + 1

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    """安全寫入非 0 儲存格"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

# ----------------------------------------------------
# 3. 單頁背景獨立辨識函式
# ----------------------------------------------------
def process_single_page_worker(task, api_key, model_name, prompt):
    file_name, p_idx, total_p, mime_type, payload_bytes = task
    client = genai.Client(api_key=api_key)
    
    # 輕微錯開執行緒，避免瞬間觸發 API 限速
    time.sleep(random.uniform(0.3, 1.2))

    last_err = ""
    # 針對同一模型原地重試，遇 429 退避等待
    for attempt in range(1, 6):
        try:
            res = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=payload_bytes, mime_type=mime_type),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=StationReport,
                    temperature=0.0,
                )
            )
            if res and res.text:
                clean_json = extract_json_str(res.text)
                data = StationReport.model_validate_json(clean_json)
                if data and data.date_day > 0:
                    return (file_name, p_idx, data, None)
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "RESOURCE_EXHAUSTED" in last_err:
                sleep_time = 4.0 * attempt + random.uniform(1.0, 2.0)
                time.sleep(sleep_time)
                continue
            time.sleep(1.5)
            
    return (file_name, p_idx, None, last_err)

# ----------------------------------------------------
# 4. 介面上傳區塊
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳 Excel 公版 (.xlsx 或 .xls)", type=["xlsx", "xls"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選同時處理)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 5. 核心處理引擎
# ----------------------------------------------------
if st.button("🚀 開始極速自動辨識與填表", type="primary", use_container_width=True):
    if not active_api_key:
        st.error("❌ 尚未設定 Gemini API Key！")
        st.stop()
    if not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請確認已上傳 Excel 公版與 PDF 檔案！")
        st.stop()

    start_time = time.time()
    client = genai.Client(api_key=active_api_key)

    # 階段 1：執行前 1 秒連線快篩
    status_box = st.status("📄 [階段 1/3] 正在驗證模型與拆分 PDF 頁面...", expanded=True)
    
    verified_model = selected_model
    try:
        test_res = client.models.generate_content(
            model=selected_model,
            contents="hello",
            config=types.GenerateContentConfig(max_output_tokens=2)
        )
    except Exception as e:
        status_box.write(f"⚠️ 模型 `{selected_model}` 回應異常，自動切換至備用主力 `gemini-1.5-flash`...")
        verified_model = "gemini-1.5-flash"

    status_box.write(f"🎯 核心模型已鎖定：`{verified_model}`")

    # 載入 Excel
    try:
        wb = load_excel_preserving_all(uploaded_excel.getvalue(), uploaded_excel.name)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗：{e}")
        st.stop()

    # 高速抽取 PDF 掃描圖檔
    all_pages = []
    for pdf_file in uploaded_pdfs:
        try:
            reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
            total_p = len(reader.pages)
            status_box.write(f"🔍 讀入檔案 `{pdf_file.name}`（共 {total_p} 頁）")
            for p_idx, page in enumerate(reader.pages, 1):
                mime_type, payload_bytes = extract_page_image_payload(page)
                all_pages.append((pdf_file.name, p_idx, total_p, mime_type, payload_bytes))
        except Exception as e:
            status_box.write(f"⚠️ 檔案 `{pdf_file.name}` 讀取異常：{e}")

    total_tasks = len(all_pages)
    if total_tasks == 0:
        status_box.update(label="❌ 沒有找到可處理的 PDF 頁面", state="error")
        st.stop()

    status_box.update(label=f"🚀 [階段 2/3] 啟動 {max_workers} 執行緒並行辨識全數 {total_tasks} 頁單據...", state="running")

    prompt = """
    你是一位具備會計勾稽專業的台鐵表單辨識專家。請仔細辨識這張站務解款單據影像：
    1. 【基本資料】：
       - 車站名稱：請精準擷取單據頂部的車站名稱（例如：埔心、楊梅、富岡、北湖、湖口、新豐、竹北、北新竹、新竹、香山、三姓橋、千甲、新莊、竹中、六家、竹東、內灣、竹南、大山、後龍、白沙屯、新埔、通霄、苑裡、日南、大甲、台中港、清水、沙鹿、龍井、大肚、追分等），勿含「站」字。
       - 日期：請擷取「進款日期」的『日/號數』（1至31整數）。
    2. 【淺色印痕強化】：
       - 單據為影印或複寫掃描，部分數字可能較淡或斷線，請特別留意高位數與小數點，勿將 3 誤判為 8、勿將 0 誤判為 6。
    3. 【應解款數金額明細擷取（若無填 0）】：
       - 客運(+)
       - 貨運(+)
       - 信用卡刷卡(-)（填正數）
       - 信用卡退刷(+)（填正數）
       - 條碼支付進款(-)（填正數）
       - 條碼支付退款(+)（填正數）
       - 應解總計（單據上的應解總計數字）
       - 其他項目淨額（若應解款數內有其他獨立明細如「存付運費」、「託收支票」、「補繳金額」或「短欠」等，請依正負計算淨額；無則填 0）
    4. 【會計平衡嚴格自檢】：
       應解總計 = 客運(+) + 貨運(+) - (信用卡刷卡(-) - 信用卡退刷(+)) - (條碼支付進款(-) - 條碼支付退款(+)) + 其他項目
       若初次計算與【應解總計】不符，代表有模糊位數讀錯，請立即重新核對算式中每個數字直到相符！
    """

    results_data = []
    failed_tasks = []
    progress_bar = st.progress(0)
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {
            executor.submit(process_single_page_worker, task, active_api_key, verified_model, prompt): task
            for task in all_pages
        }
        
        for future in as_completed(future_to_task):
            file_name, p_idx, data, err = future.result()
            completed_count += 1
            progress_bar.progress(
                completed_count / total_tasks,
                text=f"🚀 辨識中：已完成 {completed_count}/{total_tasks} 頁 ({int(completed_count/total_tasks*100)}%)"
            )
            
            if data and data.date_day > 0:
                results_data.append((file_name, p_idx, data))
                status_box.write(f"⚡ [{completed_count:02d}/{total_tasks:02d}] **{data.station_name}**（{data.date_day} 日）辨識成功")
            else:
                failed_tasks.append((file_name, p_idx, err))
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁：辨識未果 ({err[:80] if err else '解析空白'})")

    progress_bar.empty()

    if not results_data:
        status_box.update(label="❌ 辨識失敗：未成功辨識任何單據！", state="error")
        st.error("⚠️ 未能成功擷取到單據資料。以下為詳細錯誤原因，請核對 API 金鑰：")
        for fn, p, err in failed_tasks[:10]:
            st.code(f"{fn} 第 {p} 頁: {err}")
        st.stop()

    status_box.update(label=f"📝 [階段 3/3] 正在將 {len(results_data)} 筆單據填入 Excel 對應分頁與日期...", state="running")

    # 依檔案與頁碼排序
    results_data.sort(key=lambda x: (x[0], x[1]))

    # ----------------------------------------------------
    # 6. 回填 Excel 與 會計立場平衡檢查
    # ----------------------------------------------------
    total_written = 0
    success_count = 0
    audit_records = []

    for file_name, p_idx, data in results_data:
        target_sheet = find_matching_sheet(wb, data.station_name)

        net_credit = data.credit_card_charge - data.credit_card_refund
        net_barcode = data.barcode_in - data.barcode_refund

        # 會計平衡計算：總計 = 客運 + 貨運 - 電腦信用卡 - 條碼 + 其他
        computed_total = data.passenger_revenue + data.freight_revenue - net_credit - net_barcode + data.other_amount
        diff = round(data.remittance_total - computed_total, 2)
        is_balanced = (abs(diff) < 0.01)

        audit_records.append({
            "檔案名稱": file_name,
            "頁數": f"第 {p_idx} 頁",
            "車站名稱": data.station_name,
            "日期": f"{data.date_day} 日",
            "客運": to_clean_num(data.passenger_revenue),
            "貨運": to_clean_num(data.freight_revenue),
            "電腦信用卡 (=刷卡-退刷)": f"{to_clean_num(data.credit_card_charge)} - {to_clean_num(data.credit_card_refund)}" if data.credit_card_refund > 0 else to_clean_num(data.credit_card_charge),
            "條碼 (=進款-退款)": f"{to_clean_num(data.barcode_in)} - {to_clean_num(data.barcode_refund)}" if data.barcode_refund > 0 else to_clean_num(data.barcode_in),
            "其他": to_clean_num(data.other_amount),
            "自輸 (PDF應解總計)": to_clean_num(data.remittance_total),
            "計算總計": to_clean_num(computed_total),
            "差額 (自輸-總計)": diff,
            "平衡狀態": "✅ 平衡 (0)" if is_balanced else f"❌ 差額 {diff:+.0f}",
            "工作表寫入": f"已寫入 [{target_sheet.title}]" if target_sheet else "❌ 找不到分頁"
        })

        if not target_sheet:
            continue

        col_map = analyze_sheet_structure(target_sheet)
        target_row = find_target_row(target_sheet, data.date_day, col_map["date"])

        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue))
        
        cc_formula = build_deduction_formula(data.credit_card_charge, data.credit_card_refund)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)

        bc_formula = build_deduction_formula(data.barcode_in, data.barcode_refund)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)

        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))

        success_count += 1

    status_box.update(label="🎉 全部辨識與 Excel 自動寫入完成！", state="complete")

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒！共成功處理 {success_count}/{total_tasks} 頁單據，填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 7. 會計平衡檢核儀表板
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 提示：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」有差額，請對照下方清單確認原始單據是否有特殊折讓或備註！")
        else:
            st.success("🎯 太棒了！全數單據「自輸 - 總計」皆精確等於 0，完全平衡！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
