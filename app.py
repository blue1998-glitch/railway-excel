import io
import os
import time
import openpyxl
import pandas as pd
from pydantic import BaseModel, Field
from pypdf import PdfReader, PdfWriter
import streamlit as st
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 支援 .xls/.xlsx ｜ 🔍 自動探測可用模型 ｜ 🧮 自動相減公式 ｜ ⚖️ 自輸總計平衡檢查")

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
    
    # 動態取得此金鑰可用的模型清單
    available_models = []
    if active_api_key:
        try:
            temp_client = genai.Client(api_key=active_api_key)
            for m in temp_client.models.list():
                m_name = getattr(m, "name", "").replace("models/", "")
                if "gemini" in m_name.lower():
                    available_models.append(m_name)
        except Exception:
            pass

    # 若自動查詢失敗的保底清單
    if not available_models:
        available_models = [
            "gemini-1.5-flash",
            "gemini-1.5-flash-latest",
            "gemini-2.0-flash",
            "gemini-1.5-pro",
            "gemini-2.0-flash-exp"
        ]

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=available_models,
        index=0,
        help="系統已自動連線篩選出您金鑰目前支援的模型"
    )

    if active_api_key:
        st.success("✅ API 金鑰連線正常")
    else:
        st.warning("⚠️ 請確認已在 secrets 設定或在此輸入 API Key")

    st.markdown("---")
    st.markdown("""
    💡 **操作流程**：
    1. 上傳 Excel 公版（支援 `.xlsx` 或 `.xls`）
    2. 批次上傳掃描 PDF 解款單
    3. 點擊「開始辨識與填表」
    4. 檢視**平衡勾稽核對表**（確認 自輸 - 總計 = 0）
    5. 下載填寫完成的 Excel 報表
    """)

# ----------------------------------------------------
# 2. 定義資料結構與輔助函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱（例如：豐富、苗栗、銅鑼、三義、新竹、竹北、楊梅等，不需包含'站'字）")
    date_day: int = Field(description="報表日期中的『日/號』(1 至 31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】區塊內的『客運』金額，無則為 0")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】區塊內的『貨運』金額，無則為 0")
    credit_card_pos: float = Field(default=0.0, description="左側【應解款數】區塊內的『信用卡刷卡(+)』金額 (正數)，無則為 0")
    credit_card_neg: float = Field(default=0.0, description="左側【應解款數】區塊內的『信用卡刷卡(-)』金額 (填正數)，無則為 0")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】區塊內的『條碼支付進款』金額 (正數)，無則為 0")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】區塊內的『條碼支付退款』金額 (填正數)，無則為 0")
    other_amount: float = Field(default=0.0, description="左側【應解款數】區塊內除上述項目外的其他明細金額加總，無則為 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】區塊內的『應解總計』金額")

def extract_json_str(text: str) -> str:
    """安全解析模型回傳的 JSON 字串，去除 Markdown 標籤"""
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
    """標準化車站名稱以利比對"""
    if not val:
        return ""
    return str(val).replace("臺", "台").replace("站", "").replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    """將浮點數轉換為乾淨整數或保留小數"""
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def load_excel_workbook(file_bytes, filename):
    """支援 .xlsx 與舊版 .xls 格式轉換為 openpyxl 活頁簿"""
    if filename.lower().endswith(".xls"):
        try:
            import xlrd
        except ImportError:
            raise ImportError("系統尚未安裝 xlrd 套件，請在 requirements.txt 中加入 xlrd。")

        xls_book = xlrd.open_workbook(file_contents=file_bytes)
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
        return wb
    else:
        return openpyxl.load_workbook(io.BytesIO(file_bytes))

def build_subtraction_value_or_formula(pos_val, neg_val):
    """建構自動相減公式"""
    p = to_clean_num(pos_val)
    n = to_clean_num(neg_val)
    if p == 0 and n == 0:
        return None
    if p > 0 and n > 0:
        return f"={p}-{n}"
    if n > 0 and p == 0:
        return f"=-{n}"
    return p

def analyze_sheet_structure(sheet):
    """智慧掃描 Excel 工作表表頭，動態抓取欄位索引"""
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
            elif ("自輸" in val or "字輸" in val or "應解總計" in val or "應解" in val) and "remittance" not in col_map:
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
    """在工作表的日期欄中找到對應日期的列號 (1~31)"""
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
    """寫入儲存格函式：數值為 0 或空值時不填入"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

# ----------------------------------------------------
# 3. 檔案上傳介面
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳 Excel 公版檔案 (.xlsx 或 .xls)", type=["xlsx", "xls"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 核心處理邏輯
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
        wb = load_excel_workbook(uploaded_excel.getvalue(), uploaded_excel.name)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗：{e}")
        st.stop()

    status_box = st.status("📄 [階段 1/2] 正在讀取與拆分 PDF 頁面...", expanded=True)
    all_pages = []
    
    for pdf_file in uploaded_pdfs:
        status_box.write(f"🔍 讀取檔案 `{pdf_file.name}` ...")
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
            status_box.write(f"⚠️ 檔案 `{pdf_file.name}` 解析異常：{e}")

    total_tasks = len(all_pages)
    if total_tasks == 0:
        status_box.update(label="❌ 沒有找到可處理的 PDF 頁面", state="error")
        st.stop()

    status_box.update(label=f"🚀 [階段 2/2] 正在辨識 {total_tasks} 頁解款單據並填寫 Excel...", state="running")
    
    prompt = """
    你是一位專業精確的台鐵會計表單辨識專家。請仔細檢視這張解款單據：
    1. 擷取【車站名稱】（如：豐富、苗栗、三義、新竹等）與報表【日期】（僅需日/號數，1-31 的整數）。
    2. 專注看左側【應解款數】大項目區塊，精確擷取各數值（若無或空白填 0）：
       - 客運
       - 貨運
       - 信用卡刷卡(+) (填正數)
       - 信用卡刷卡(-) (填正數，不要填負號)
       - 條碼支付進款 (填正數)
       - 條碼支付退款 (填正數，不要填負號)
       - 其他項目總和（若應解款數大項內還有其他獨立款項明細則加總，否則填 0）
       - 應解總計
    3. 特別注意：若有信用卡刷卡(-)或條碼支付退款，切勿重複計入「其他項目」！
    """

    # 組合嘗試順序：使用者選擇的模型優先，接著自動嘗試可用的其他備援模型
    models_to_try = [selected_model] + [m for m in available_models if m != selected_model]

    results_data = []
    audit_records = []
    progress_bar = st.progress(0)
    
    for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1):
        progress_bar.progress(idx / total_tasks, text=f"正在辨識第 {idx}/{total_tasks} 頁...")
        
        parsed_data = None
        err_msg = ""
        
        for model_name in models_to_try:
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
                    parsed_data = StationReport.model_validate_json(clean_text)
                    break
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                    time.sleep(3.0)
                continue

        if parsed_data and parsed_data.date_day > 0:
            results_data.append((file_name, p_idx, parsed_data))
            status_box.write(f"✅ [{idx:02d}/{total_tasks:02d}] **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功")
        else:
            status_box.write(f"⚠️ [{idx:02d}/{total_tasks:02d}] `{file_name}` 第 {p_idx} 頁：辨識失敗 ({err_msg[:80]}...)")
        
        time.sleep(0.5)

    # ----------------------------------------------------
    # 5. 回填 Excel 與 平衡檢查
    # ----------------------------------------------------
    total_written = 0
    success_count = 0

    for file_name, p_idx, data in results_data:
        target_name_clean = clean_station_name(data.station_name)
        
        target_sheet = None
        for s_name in wb.sheetnames:
            s_clean = clean_station_name(s_name)
            if s_clean == target_name_clean or s_clean in target_name_clean or target_name_clean in s_clean:
                target_sheet = wb[s_name]
                break

        net_credit = data.credit_card_pos - data.credit_card_neg
        net_barcode = data.barcode_in - data.barcode_refund
        computed_total = data.passenger_revenue + data.freight_revenue + net_credit + net_barcode + data.other_amount
        diff = round(data.remittance_total - computed_total, 2)
        is_balanced = (abs(diff) < 0.01)

        audit_records.append({
            "檔案名稱": file_name,
            "車站名稱": data.station_name,
            "日期": f"{data.date_day} 日",
            "客運": to_clean_num(data.passenger_revenue),
            "貨運": to_clean_num(data.freight_revenue),
            "信用卡刷卡": f"{to_clean_num(data.credit_card_pos)} - {to_clean_num(data.credit_card_neg)}" if data.credit_card_neg > 0 else to_clean_num(data.credit_card_pos),
            "條碼支付": f"{to_clean_num(data.barcode_in)} - {to_clean_num(data.barcode_refund)}" if data.barcode_refund > 0 else to_clean_num(data.barcode_in),
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

        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue))
        
        cc_formula = build_subtraction_value_or_formula(data.credit_card_pos, data.credit_card_neg)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)

        bc_formula = build_subtraction_value_or_formula(data.barcode_in, data.barcode_refund)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)

        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))

        success_count += 1

    status_box.update(label="🎉 辨識與 Excel 寫入完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，共成功處理 {success_count}/{total_tasks} 頁單據，填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核儀表板 (自輸 - 總計 = 0 檢查)
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 警告：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請依下方表格核對單據金額是否有誤！")
        else:
            st.success("🎯 太棒了！所有辨識成功的單據「自輸 - 總計」皆等於 0，會計平衡完全正確！")

        st.dataframe(
            df_audit,
            use_container_width=True,
            hide_index=True
        )

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    
