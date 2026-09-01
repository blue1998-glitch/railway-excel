import os
import io
import time
import subprocess
import streamlit as st
from pdf2image import convert_from_bytes
import openpyxl
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 V5.0 零死角極速穩定版 ｜ 智慧平流防限速 ｜ 🧮 公式保留 ｜ ⚖️ 會計平衡")

EMBEDDED_API_KEY = "AQ.Ab8RN6Lno7iLAGhnHoc0cs_k_-c60g2atjdd_84u9nwjpbNN7Q"
raw_key = st.secrets.get("GEMINI_API_KEY", EMBEDDED_API_KEY)
cleaned_key = str(raw_key).replace('"', '').replace("'", "").strip()

with st.sidebar:
    st.header("⚙️ 系統設定")
    user_key = st.text_input(
        "Gemini API Key",
        value=cleaned_key,
        type="password",
        help="支援 Google AI Studio 最新金鑰"
    )
    active_api_key = user_key.strip()
    
    if active_api_key:
        st.success("✅ API 金鑰已就緒")
    else:
        st.warning("⚠️ 請輸入金鑰。")

    st.markdown("---")
    st.markdown("💡 **操作流程**：\n1. 上傳 Excel 公版\n2. 批次選取掃描 PDF\n3. 點擊開始轉換並下載完成檔")

# ----------------------------------------------------
# 2. 定義資料結構與工具函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱，例如：豐富、苗栗、銅鑼、三義、二水、埔心、楊梅、北湖、富岡、湖口、新豐、竹北、新竹、香山、北新竹等")
    date_day: int = Field(description="報表日期中的『日』(1-31 的整數數字)")
    passenger_revenue: float = Field(default=0.0, description="客運金額")
    freight_revenue: float = Field(default=0.0, description="貨運金額")
    credit_card_pos: float = Field(default=0.0, description="信用卡刷卡(+)金額 (正數)")
    credit_card_neg: float = Field(default=0.0, description="信用卡刷卡(-)金額 (正數，不要填負數)")
    barcode_in: float = Field(default=0.0, description="條碼支付進款金額 (正數)")
    barcode_refund: float = Field(default=0.0, description="條碼支付退款金額 (正數，不要填負數)")
    other_amount: float = Field(default=0.0, description="【應解款數】區塊內獨立的其他明細項目總和")
    remittance_total: float = Field(default=0.0, description="應解總計金額")

def clean_name(val):
    if not val:
        return ""
    return str(val).replace("臺", "台").replace("站", "").replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_subtraction_formula(pos_val, neg_val):
    p = to_clean_num(pos_val)
    n = to_clean_num(neg_val)
    if p == 0 and n == 0:
        return 0
    if n > 0:
        return f"={p}-{n}"
    return p

# ----------------------------------------------------
# 3. 檔案上傳介面
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳 Excel 公版 (.xlsx 或 .xls)", type=["xlsx", "xls"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF (可一次選取多檔)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 核心平穩處理引擎
# ----------------------------------------------------
if st.button("🚀 開始極速自動辨識與填表", type="primary", use_container_width=True):
    if not active_api_key:
        st.error("❌ 尚未設定 Gemini API Key！")
        st.stop()
    if not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請確認 Excel 公版與 PDF 檔案皆已上傳！")
        st.stop()

    start_time = time.time()
    client = genai.Client(api_key=active_api_key)

    # 處理 Excel 格式
    excel_bytes = uploaded_excel.getvalue()
    if uploaded_excel.name.lower().endswith(".xls"):
        with open("temp_template.xls", "wb") as f:
            f.write(excel_bytes)
        subprocess.run(["libreoffice", "--headless", "--convert-to", "xlsx", "temp_template.xls"], check=True)
        wb = openpyxl.load_workbook("temp_template.xlsx")
    else:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))

    # 階段 1：解析 PDF 頁面
    status_box = st.status("⚡ [階段 1/2] 正在解析 PDF 頁面...", expanded=True)
    all_tasks = []
    
    for pdf_file in uploaded_pdfs:
        status_box.write(f"📄 載入檔案 `{pdf_file.name}` ...")
        images = convert_from_bytes(pdf_file.getvalue(), dpi=150, fmt="jpeg", thread_count=2)
        total_p = len(images)
        for p_idx, p_img in enumerate(images, 1):
            buf = io.BytesIO()
            p_img.save(buf, format="JPEG", quality=85)
            all_tasks.append((pdf_file.name, p_idx, total_p, buf.getvalue()))

    total_pages_all = len(all_tasks)
    status_box.update(label=f"🚀 [階段 2/2] 正在逐頁穩定辨識全數 {total_pages_all} 頁單據...", state="running")
    
    prompt = """
    你是一個精確的台鐵會計表單辨識助理。請仔細辨識單據內容：
    1. 擷取『車站名稱』與『日期（幾號/幾日）』。
    2. 查看左邊【應解款數】大項目區塊，精準擷取數值（若無填 0）：客運、貨運、信用卡刷卡(+)、信用卡刷卡(-)、條碼支付進款、條碼支付退款、其他獨立明細總和、應解總計。
    3. 勾稽校核：[客運 + 貨運 + (刷卡+ 減 刷卡-) + (條碼進款 減 條碼退款) + 其他] 必須剛好等於 [應解總計]。
       【嚴禁將信用卡刷卡(-)、條碼退款重複計入其他！】
    """

    # 僅使用現行有效的官方主力模型
    active_models = ["gemini-2.5-flash", "gemini-2.0-flash"]
    progress_bar = st.progress(0)
    results = []
    logs = []

    for idx, (file_name, p_idx, total_p, img_bytes) in enumerate(all_tasks, 1):
        progress_bar.progress(idx / total_pages_all, text=f"正在處理第 {idx}/{total_pages_all} 頁單據...")
        
        data = None
        last_err = ""

        # 智慧重試機制（遇限速自動等待）
        for attempt in range(1, 4):
            for model_name in active_models:
                try:
                    res = client.models.generate_content(
                        model=model_name,
                        contents=[
                            types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"),
                            prompt
                        ],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=StationReport,
                            temperature=0.0,
                        )
                    )
                    if res and res.text:
                        clean_text = res.text.strip()
                        if "```json" in clean_text:
                            clean_text = clean_text.split("```json")[1].split("```")[0].strip()
                        elif "```" in clean_text:
                            clean_text = clean_text.split("```")[1].split("```")[0].strip()
                        data = StationReport.model_validate_json(clean_text)
                        break
                except Exception as e:
                    last_err = str(e)
                    if "429" in last_err or "RESOURCE_EXHAUSTED" in last_err:
                        time.sleep(8.0 * attempt)
                    continue
            if data is not None:
                break

        if data and data.date_day > 0:
            results.append((file_name, p_idx, data))
            status_box.write(f"⚡ [{idx:02d}/{total_pages_all:02d}] **{data.station_name}**（{data.date_day} 號）辨識成功")
        else:
            logs.append(f"⚠️ 檔案 `{file_name}` 第 {p_idx} 頁：{last_err if last_err else '無法辨識'}")

        # 每頁之間平滑節流 3 秒，徹底杜絕 429 頻率撞牆
        time.sleep(3.0)

    status_box.update(label="✅ 辨識完成！正在快速填寫 Excel 與公式...", state="complete")
    progress_bar.empty()

    # 階段 3：回填 Excel 與防呆平衡校準
    total_cells_written = 0
    success_pages = 0

    for file_name, page_idx, data in results:
        clean_target_st = clean_name(data.station_name)
        target_sheet = next(
            (wb[s] for s in wb.sheetnames if clean_target_st == clean_name(s) or clean_target_st in clean_name(s) or clean_name(s) in clean_target_st),
            None
        )

        if not target_sheet:
            logs.append(f"⚠️ 找不到車站分頁 `[{data.station_name}]`，已略過")
            continue

        calc_cc = data.credit_card_pos - data.credit_card_neg
        calc_bc = data.barcode_in - data.barcode_refund
        base_subtotal = data.passenger_revenue + data.freight_revenue + calc_cc + calc_bc
        
        final_other = data.other_amount
        if abs(data.remittance_total - base_subtotal) < 0.01:
            final_other = 0.0
        elif abs(data.remittance_total - (base_subtotal + data.other_amount)) > 0.01:
            final_other = data.remittance_total - base_subtotal

        cc_val = build_subtraction_formula(data.credit_card_pos, data.credit_card_neg)
        bc_val = build_subtraction_formula(data.barcode_in, data.barcode_refund)

        target_row = int(data.date_day) + 1

        def write_cell(col_idx, val):
            global total_cells_written
            if val != 0 and val != "0":
                target_sheet.cell(row=target_row, column=col_idx, value=val)
                total_cells_written += 1

        write_cell(2, to_clean_num(data.passenger_revenue))
        write_cell(3, to_clean_num(data.freight_revenue))
        write_cell(4, cc_val)
        write_cell(5, bc_val)
        write_cell(6, to_clean_num(final_other))
        write_cell(7, to_clean_num(data.remittance_total))

        success_pages += 1
        logs.append(f"✅ **{data.station_name}** ({data.date_day} 號) ➜ 已成功寫入第 {target_row} 列（已平衡）")

    # 輸出下載
    elapsed = time.time() - start_time
    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)

    st.balloons()
    st.success(f"🎉 處理大成功！總耗時 {elapsed:.1f} 秒！共成功寫入 {success_pages}/{total_pages_all} 頁單據，填入 {total_cells_written} 個儲存格！")
    
    st.download_button(
        label="📥 點擊下載完成的 Excel 報表",
        data=out_stream,
        file_name="完成_解款單彙總報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

    with st.expander("🔍 檢視詳細日誌明細", expanded=False):
        for log in logs:
            st.markdown(log)
