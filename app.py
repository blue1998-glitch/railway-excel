import io
import os
import re
import time
from typing import Any, Union
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
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 穩定高速辨識 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

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
    
    available_models = []
    if active_api_key:
        try:
            temp_client = genai.Client(api_key=active_api_key)
            for m in temp_client.models.list():
                m_name = getattr(m, "name", "").replace("models/", "")
                if "gemini" in m_name.lower() and "2.5" not in m_name:
                    available_models.append(m_name)
        except Exception:
            pass

    if not available_models:
        available_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    else:
        flash_models = [m for m in available_models if "flash" in m.lower()]
        other_models = [m for m in available_models if "flash" not in m.lower()]
        available_models = flash_models + other_models

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=available_models,
        index=0,
        help="建議選擇 flash 系列模型，速度快且辨識精準"
    )

    concurrency = st.slider(
        "⚡ 並行辨識線程數",
        min_value=1,
        max_value=5,
        value=3,
        help="推薦設為 3。兼顧極速辨識與 API 速率上限，避免觸發 429 限流"
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
# 2. 定義資料結構與工具函式（強化容錯性）
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(default="", description="車站名稱，例如豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，不需包含'站'字或站碼")
    date_day: Any = Field(default=0, description="進款日期的『日/號』(1 至 31 的整數數字，若被遮擋可參考繳款編號如第28號即為28)")
    passenger_revenue: Any = Field(default=0, description="客運(+)金額")
    freight_revenue: Any = Field(default=0, description="貨運(+)金額")
    credit_card_charge: Any = Field(default=0, description="信用卡刷卡(-)金額（正數）")
    credit_card_refund: Any = Field(default=0, description="信用卡退刷(+)金額（正數）")
    barcode_in: Any = Field(default=0, description="條碼支付進款(-)金額（正數）")
    barcode_refund: Any = Field(default=0, description="條碼支付退款(+)金額（正數）")
    other_amount: Any = Field(default=0, description="其他項目淨額（存付運費、補繳金額、繳回週轉金、繳回零用金、繳回找零金、託收支票等明細之合計）")
    remittance_total: Any = Field(default=0, description="應解總計金額")

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

def clean_amount(val) -> float:
    """強化清洗：過濾千分位逗號、括號負數與非數字字元，徹底避免數值報錯"""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("，", "").replace(" ", "").replace("$", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        nums = re.findall(r"[-+]?\d*\.?\d+", s)
        return float(nums[0]) if nums else 0.0

def to_clean_num(val):
    """轉為乾淨整數或保留小數"""
    f_val = clean_amount(val)
    return int(f_val) if f_val.is_integer() else f_val

def clean_station_name(val):
    """標準化車站名稱，自動剔除站碼前綴（如 3150_）以精準比對工作表"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def build_deduction_formula(charge_val, refund_val):
    """會計立場扣抵公式：當有退刷/退款時寫入 =進款-退款"""
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
    """精準鎖定表頭行，防止被試算表頂部的列印日期帶偏"""
    col_map = {}
    header_row = None
    for r in range(1, 8):
        for c in range(1, sheet.max_column + 1):
            val = str(sheet.cell(row=r, column=c).value or "").replace(" ", "").replace("\n", "").strip()
            if "客運" in val:
                header_row = r
                break
        if header_row:
            break

    search_rows = [header_row] if header_row else range(1, 6)
    for r in search_rows:
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

    defaults = {"date": 1, "passenger": 2, "freight": 3, "credit": 4, "barcode": 5, "other": 6, "remittance": 7}
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return col_map

def find_target_row(sheet, date_day, date_col=1):
    """三層防禦尋找對應日期的列號 (1~31)"""
    target = int(clean_amount(date_day))
    if target <= 0:
        return None
    # 1. 優先從偵測到的 date_col 尋找
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            try:
                s = str(val).replace("日", "").replace("號", "").strip()
                if int(float(s)) == target:
                    return r
            except Exception:
                pass
    # 2. 備用：前三欄全域掃描
    for c in [1, 2, 3]:
        for r in range(1, 45):
            val = sheet.cell(row=r, column=c).value
            if val is not None:
                try:
                    s = str(val).replace("日", "").replace("號", "").strip()
                    if int(float(s)) == target:
                        return r
                except Exception:
                    pass
    # 3. 預設推算位置
    return target + 3

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    """安全寫入指定儲存格（0 或空值不覆寫）"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx and row_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=5):
    """單頁辨識函式，內建指數退避重試機制徹底克服 429 限流"""
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
                # 遇到頻率限制，採指數退避等待
                time.sleep(2.5 * (2 ** retry))
                continue
            elif retry == max_retries - 1:
                return None, err
            time.sleep(1.5)
    return None, "超過最大重試次數"

# ----------------------------------------------------
# 3. 介面上傳區塊
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳公版 Excel 檔案 (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 核心辨識與填寫流程
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

    status_box.update(label=f"⚡ [階段 2/2] {concurrency} 線程穩定辨識 {total_tasks} 頁解款單據...", state="running")
    
    prompt = """
    你是一位專業精確的台鐵會計表單辨識專家。請仔細檢視這張解款單據：
    1. 擷取【車站名稱】（如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）。
    2. 擷取進款【日期】（僅需 1-31 的整數數字。若進款日期被印章遮擋，請參考繳款編號如『第28號』即為 28，或列印時間）。
    3. 專注看左側【應解款數】大項目區塊，精確擷取各數值（若無填 0，請直接填數字）：
       - 客運(+)
       - 貨運(+)
       - 信用卡刷卡(-) (填正數)
       - 信用卡退刷(+) (若有退刷務必填正數金額，無則填 0)
       - 條碼支付進款(-) (填正數)
       - 條碼支付退款(+) (若有退款務必填正數金額，無則填 0)
       - 其他項目加總 (包含存付運費、補繳金額、繳回週轉金、繳回零用金、找零金、託收支票等其他項目的淨額合計)
       - 應解總計 (報表上的應解總計數值)
    4. 會計勾稽驗算原則：必符合「客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計」。
    """

    results_data = []
    progress_bar = st.progress(0)
    completed_count = 0

    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        future_map = {}
        for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1):
            future = executor.submit(call_gemini_page, client, selected_model, page_bytes, prompt)
            future_map[future] = (idx, file_name, p_idx, total_p)
            time.sleep(0.3)  # 平滑微幅間隔，避免瞬間並發衝擊

        for future in as_completed(future_map):
            completed_count += 1
            idx, file_name, p_idx, total_p = future_map[future]
            progress_bar.progress(
                completed_count / total_tasks,
                text=f"⚡ 辨識處理中：已完成 {completed_count}/{total_tasks} 頁 ({completed_count * 100 // total_tasks}%)"
            )
            
            try:
                parsed_data, err_msg = future.result()
                day_val = int(clean_amount(parsed_data.date_day)) if parsed_data else 0
                if parsed_data and day_val > 0:
                    results_data.append((idx, file_name, p_idx, parsed_data))
                    status_box.write(f"✅ **{parsed_data.station_name}**（{day_val} 日）辨識成功 - `{file_name}` 第 {p_idx} 頁")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗：{str(err_msg)[:60]}")
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁執行異常：{e}")

    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel 與 會計平衡檢查
    # ----------------------------------------------------
    status_box.update(label="📝 正在將辨識資料寫入 Excel 公版...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_name_clean = clean_station_name(data.station_name)
        
        # 精準配對工作表（特別保護烏日與新烏日不互搶）
        target_sheet = None
        for s_name in wb.sheetnames:
            if clean_station_name(s_name) == target_name_clean:
                target_sheet = wb[s_name]
                break
        if not target_sheet:
            for s_name in wb.sheetnames:
                s_clean = clean_station_name(s_name)
                if not s_clean:
                    continue
                # 防止烏日與新烏日互相誤配
                if ("新烏日" in target_name_clean and s_clean == "烏日") or (target_name_clean == "烏日" and "新烏日" in s_clean):
                    continue
                if s_clean in target_name_clean or target_name_clean in s_clean:
                    target_sheet = wb[s_name]
                    break

        passenger = clean_amount(data.passenger_revenue)
        freight = clean_amount(data.freight_revenue)
        card_charge = clean_amount(data.credit_card_charge)
        card_refund = clean_amount(data.credit_card_refund)
        barcode_in = clean_amount(data.barcode_in)
        barcode_refund = clean_amount(data.barcode_refund)
        other = clean_amount(data.other_amount)
        remittance = clean_amount(data.remittance_total)
        day_num = int(clean_amount(data.date_day))

        net_credit = card_charge - card_refund
        net_barcode = barcode_in - barcode_refund
        computed_total = passenger + freight - net_credit - net_barcode + other
        diff = round(remittance - computed_total, 2)
        is_balanced = (abs(diff) < 0.01)

        cc_formula = build_deduction_formula(card_charge, card_refund)
        bc_formula = build_deduction_formula(barcode_in, barcode_refund)

        audit_records.append({
            "檔案名稱": file_name,
            "車站名稱": data.station_name,
            "日期": f"{day_num} 日",
            "客運": to_clean_num(passenger),
            "貨運": to_clean_num(freight),
            "電腦信用卡 (=刷卡-退刷)": cc_formula if cc_formula is not None else 0,
            "條碼 (=進款-退款)": bc_formula if bc_formula is not None else 0,
            "其他": to_clean_num(other),
            "自輸 (應解總計)": to_clean_num(remittance),
            "計算總計": to_clean_num(computed_total),
            "差額": diff,
            "平衡狀態": "✅ 平衡 (0)" if is_balanced else f"❌ 差額 {diff:+.0f}",
            "工作表狀態": f"寫入 [{target_sheet.title}]" if target_sheet else "❌ 找不到分頁"
        })

        if not target_sheet:
            continue

        col_map = analyze_sheet_structure(target_sheet)
        target_row = find_target_row(target_sheet, day_num, col_map["date"])

        if target_row:
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(passenger))
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(freight))
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(other))
            total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(remittance))
            success_count += 1

    status_box.update(label="🎉 辨識與 Excel 寫入全數完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，共成功處理 {success_count}/{total_tasks} 頁單據，成功填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核儀表板與下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.warning(f"⚠️ 注意：共有 **{unbalanced_count}** 筆單據計算有差額，請依下方表格核對單據！")
        else:
            st.success("🎯 太棒了！所有單據會計勾稽驗算皆完全平衡 (差額 0)！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
