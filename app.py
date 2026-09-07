import io
import os
import re
import time
import openpyxl
import pandas as pd
from pydantic import BaseModel, Field
from pypdf import PdfReader, PdfWriter
import streamlit as st
from google import genai
from google.genai import types
from concurrent.futures import ThreadPoolExecutor, as_completed

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 高精度並行辨識 ｜ ⚖️ 會計立場扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

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
    
    # 支援最新高精度模型
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

    if not available_models:
        available_models = [
            "gemini-2.5-flash",
            "gemini-2.0-flash",
            "gemini-1.5-flash",
            "gemini-2.5-pro"
        ]
    else:
        # 將 flash 高速與 2.5 系列排在前面
        flash_models = [m for m in available_models if "flash" in m.lower()]
        other_models = [m for m in available_models if "flash" not in m.lower()]
        available_models = flash_models + other_models

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=available_models,
        index=0,
        help="推薦使用 flash 系列，速度快且對表格辨識精準"
    )

    concurrency = st.slider(
        "⚡ 並行辨識線程數",
        min_value=1,
        max_value=5,
        value=3,
        help="建議設定 2~3，兼具高速且不易觸發 API 頻率限制"
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
# 2. 資料結構與精準比對函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站中文名稱（例如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼、英文與'站'字）")
    date_day: int = Field(description="進款或解款日期中的『日/號數』(1 至 31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="『客運收入』或『客運(+)』金額，多數單據皆有此數值，勿漏填，無則為 0")
    freight_revenue: float = Field(default=0.0, description="『貨運收入』或『貨運(+)』金額，無則為 0")
    credit_card_charge: float = Field(default=0.0, description="『信用卡刷卡(-)』進款金額 (填正數)，無則為 0")
    credit_card_refund: float = Field(default=0.0, description="『信用卡退刷(+)』金額 (填正數)，無則為 0")
    barcode_in: float = Field(default=0.0, description="『條碼支付進款(-)』或行動支付金額 (填正數)，無則為 0")
    barcode_refund: float = Field(default=0.0, description="『條碼支付退款(+)』金額 (填正數)，無則為 0")
    other_amount: float = Field(default=0.0, description="存付運費、託收支票、補繳金額等其他項目淨額，無則為 0")
    remittance_total: float = Field(default=0.0, description="『應解總計』、『應解款總額』或『本期應解款額』金額")

def extract_json_str(text: str) -> str:
    """去除 Markdown 標籤以安全解析 JSON"""
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
    """標準化車站名稱：徹底去除站碼前綴（如 3150-、3150_）與無關符號"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^[0-9A-Za-z\-_ ]+', '', s)
    s = re.sub(r'[\s\-_()（）]', '', s)
    return s.strip()

def match_target_sheet(wb, raw_station_name):
    """精準對應工作表，防止新烏日與烏日混淆"""
    target = clean_station_name(raw_station_name)
    if not target:
        return None
    # 優先完全相符
    for name in wb.sheetnames:
        if clean_station_name(name) == target:
            return wb[name]
    # 次之依名稱長度由長到短包含匹配（避免短站名搶先誤配）
    sorted_sheets = sorted(wb.sheetnames, key=lambda x: len(clean_station_name(x)), reverse=True)
    for name in sorted_sheets:
        c_name = clean_station_name(name)
        if c_name and (c_name in target or target in c_name):
            return wb[name]
    return None

def to_clean_num(val):
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
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
    """定位表頭列並解析欄位位置，避免抓到製表日期"""
    header_row = 3
    for r in range(1, 8):
        row_str = "".join(str(sheet.cell(row=r, column=c).value or "") for c in range(1, sheet.max_column + 1))
        if "客運" in row_str or "信用卡" in row_str or "貨運" in row_str or "應解" in row_str:
            header_row = r
            break

    col_map = {}
    for c in range(1, sheet.max_column + 1):
        val = str(sheet.cell(row=header_row, column=c).value or "").replace(" ", "").replace("\n", "").strip()
        if not val:
            continue
        if "客運" in val and "passenger" not in col_map:
            col_map["passenger"] = c
        elif "貨運" in val and "freight" not in col_map:
            col_map["freight"] = c
        elif any(k in val for k in ["信用卡", "刷卡"]) and "credit" not in col_map:
            col_map["credit"] = c
        elif any(k in val for k in ["條碼", "支付"]) and "barcode" not in col_map:
            col_map["barcode"] = c
        elif "其他" in val and "other" not in col_map:
            col_map["other"] = c
        elif any(k in val for k in ["自輸", "字輸", "應解", "解款", "總計"]) and "remittance" not in col_map:
            col_map["remittance"] = c
        elif val in ["日", "日期", "日次", "號"] and "date" not in col_map:
            col_map["date"] = c

    defaults = {"date": 1, "passenger": 2, "freight": 3, "credit": 4, "barcode": 5, "other": 6, "remittance": 7}
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return header_row, col_map

def find_target_row(sheet, date_day, date_col, header_row):
    """在對應的日期欄中尋找精確列號"""
    for r in range(header_row + 1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            try:
                val_str = str(val).replace("日", "").replace("號", "").strip()
                if int(float(val_str)) == int(date_day):
                    return r
            except Exception:
                pass
    # 若在指定欄未搜尋到，嘗試第 1 欄
    if date_col != 1:
        for r in range(header_row + 1, 45):
            val = sheet.cell(row=r, column=1).value
            if val is not None:
                try:
                    val_str = str(val).replace("日", "").replace("號", "").strip()
                    if int(float(val_str)) == int(date_day):
                        return r
                except Exception:
                    pass
    return header_row + int(date_day)

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=4):
    """單頁呼叫並內建指數退避重試，防止 429 掉頁"""
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
        except Exception as e:
            err = str(e)
            if "429" in err or "RESOURCE_EXHAUSTED" in err:
                time.sleep(3.0 * (retry + 1))
                continue
            elif retry == max_retries - 1:
                return None, err
            time.sleep(1.5)
    return None, "超過重試上限"

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
        st.error("❌ 請先在側邊欄輸入 Gemini API Key！")
        st.stop()
    if not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請同時上傳 Excel 公版與 PDF 解款單！")
        st.stop()

    start_time = time.time()
    client = genai.Client(api_key=active_api_key)

    try:
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_excel.getvalue()), data_only=False)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗：{e}")
        st.stop()

    status_box = st.status("📄 [階段 1/2] 正在拆分單據頁面...", expanded=True)
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
            status_box.write(f"⚠️ 檔案 `{pdf_file.name}` 拆分異常：{e}")

    total_tasks = len(all_pages)
    if total_tasks == 0:
        status_box.update(label="❌ 未找到有效 PDF 頁面", state="error")
        st.stop()

    status_box.update(label=f"⚡ [階段 2/2] 啟動 {concurrency} 線程高速辨識 {total_tasks} 頁單據...", state="running")
    
    prompt = """
    你是一位專業的台鐵會計解款單據辨識專家。請仔細辨識這張報表：
    【重要辨識準則】：
    1. 車站與日期：
       - 擷取報表表頭的【車站名稱】（如台中、豐原、彰化、苗栗、員林、田中、二水等，去除站碼與'站'字）。
       - 擷取進款或解款日期的【日】（號數，1 至 31 的整數）。
    2. 數值擷取（字跡若較淺或為複寫件，請仔細分辨 0、3、8、1、7 等數字，切勿遺漏）：
       - 客運收入：對應「客運(+)」或「客運收入」，絕大多數單據皆有客運金額，務必精確填寫。
       - 貨運收入：對應「貨運(+)」或「貨運收入」，無則填 0。
       - 信用卡刷卡：對應「信用卡刷卡(-)」進款金額，填正數，無則填 0。
       - 信用卡退刷：對應「信用卡退刷(+)」金額，填正數，無則填 0。
       - 條碼支付進款：對應「條碼支付進款(-)」或行動支付進款，填正數，無則填 0。
       - 條碼支付退款：對應「條碼支付退款(+)」，填正數，無則填 0。
       - 其他項目：存付運費、代收、補繳等其他雜項淨額，無則填 0。
       - 應解總計：對應「應解總計」或「應解款總額」數值。
    3. 驗算校對：請依「客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計」進行核對確認。
    """

    results_data = []
    progress_bar = st.progress(0)
    completed_count = 0

    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        future_map = {
            executor.submit(call_gemini_page, client, selected_model, page_bytes, prompt): (idx, file_name, p_idx, total_p)
            for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1)
        }

        for future in as_completed(future_map):
            completed_count += 1
            idx, file_name, p_idx, total_p = future_map[future]
            progress_bar.progress(
                completed_count / total_tasks,
                text=f"辨識進度：{completed_count}/{total_tasks} 頁 ({completed_count * 100 // total_tasks}%)"
            )
            
            try:
                parsed_data, err_msg = future.result()
                if parsed_data and parsed_data.date_day > 0:
                    results_data.append((idx, file_name, p_idx, parsed_data))
                    status_box.write(f"✅ `{file_name}` 第 {p_idx} 頁：**{parsed_data.station_name}**（{parsed_data.date_day}日）辨識完成")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗：{str(err_msg)[:60]}")
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁異常：{e}")

    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel
    # ----------------------------------------------------
    status_box.update(label="📝 正在回填 Excel 儲存格...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_sheet = match_target_sheet(wb, data.station_name)

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
            "電腦信用卡": cc_formula if cc_formula is not None else 0,
            "條碼支付": bc_formula if bc_formula is not None else 0,
            "其他": to_clean_num(data.other_amount),
            "自輸(應解總計)": to_clean_num(data.remittance_total),
            "計算總計": to_clean_num(computed_total),
            "差額": diff,
            "平衡狀態": "✅ 平衡" if is_balanced else f"❌ 差額 {diff:+.0f}",
            "寫入工作表": target_sheet.title if target_sheet else "❌ 未找到對應工作表"
        })

        if not target_sheet:
            continue

        header_row, col_map = analyze_sheet_structure(target_sheet)
        target_row = find_target_row(target_sheet, data.date_day, col_map["date"], header_row)

        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))

        success_count += 1

    status_box.update(label="🎉 處理完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，成功處理 {success_count}/{total_tasks} 頁，共填入 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核表與下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡檢核表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 共有 {unbalanced_count} 筆單據「自輸 - 總計」不為 0，請參閱下方核對：")
        else:
            st.success("🎯 全部單據勾稽正確，差額皆為 0！")
        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 下載完成之 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
