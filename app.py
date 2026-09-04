import io
import json
import re
import streamlit as st
import openpyxl
import google.generativeai as genai

st.set_page_config(page_title="台鐵解款單彙總填表工具", page_icon="🚆", layout="wide")
st.title("🚆 台鐵站務解款單 ➔ 公版彙總自動填報工具")

# --- 側邊欄：API 金鑰設定 ---
with st.sidebar:
    st.header("🔑 金鑰設定")
    api_key = st.text_input(
        "Google AI Studio / Gemini API Key",
        value=st.secrets.get("GEMINI_API_KEY", ""),
        type="password",
        help="可直接輸入金鑰，或在 Streamlit Cloud 的 Secrets 中設定 GEMINI_API_KEY"
    )
    model_name = st.selectbox("辨識模型", ["gemini-2.5-flash", "gemini-1.5-flash"], index=0)
    st.caption("提示：公版若放在 GitHub 專案根目錄，名稱為 `豐富至二水解款單(公版)2.xlsx` 即可免重複上傳。")

# --- 核心工具函式 ---
STATION_LIST = [
    "豐富", "苗栗", "銅鑼", "三義", "泰安", "后里", "豐原", "栗林",
    "潭子", "頭家厝", "松竹", "太原", "精武", "臺中", "五權", "大慶",
    "烏日", "新烏日", "成功", "彰化", "花壇", "大村", "員林", "社頭", "田中", "二水"
]

def clean_station_name(text):
    """將文字正規化為標準站名，特別防範『新烏日』與『烏日』的混淆"""
    if not text:
        return None
    s = re.sub(r'[\d\s_站]', '', str(text))
    if "新烏日" in s:
        return "新烏日"
    if "烏日" in s:
        return "烏日"
    if "頭家" in s:
        return "頭家厝"
    if "台中" in s or "臺中" in s:
        return "臺中"
    for st_name in STATION_LIST:
        if st_name in s:
            return st_name
    return None

def analyze_pdf_with_gemini(pdf_bytes, key, model_version):
    """呼叫 Gemini 進行解款單的高精度結構化擷取"""
    genai.configure(api_key=key)
    model = genai.GenerativeModel(
        model_name=model_version,
        generation_config={"response_mime_type": "application/json"}
    )
    
    prompt = """
    你是一個專門辨識台灣鐵路「站務解款單」的會計專家。請審視此 PDF 每頁的解款單，輸出 JSON Array。
    每個車站包含以下欄位（所有金額請輸出整數，不要帶逗號）：
    - "station_name": 車站名稱（如 豐富、苗栗、新烏日、烏日、臺中、彰化 等）
    - "date": 進款日期（如 "2026-08-28"）
    - "passenger": 客運(+) 金額
    - "credit_card_in": 信用卡刷卡(-) 金額（若無填0）
    - "credit_card_refund": 信用卡退刷(+) 金額（若無填0，請務必精準辨識）
    - "barcode_in": 條碼支付進款(-) 金額（若無填0）
    - "barcode_refund": 條碼支付退款(+) 金額（若無填0，請務必精準辨識）
    - "revolving_fund": 繳回週轉金(+) 金額（若無填0）
    - "check_amount": 託收支票(-) 金額（若無填0）
    - "supplement": 補繳金額(+) 金額（若無填0）
    - 現金金額(amt)："cash_2000_amt", "cash_1000_amt", "cash_500_amt", "cash_200_amt", "cash_100_amt", "cash_50_amt", "cash_20_amt", "cash_10_amt", "cash_5_amt", "cash_1_amt"
    - 現金張數(cnt)："cash_2000_cnt", "cash_1000_cnt", "cash_500_cnt", "cash_200_cnt", "cash_100_cnt", "cash_50_cnt", "cash_20_cnt", "cash_10_cnt", "cash_5_cnt", "cash_1_cnt"
    - "expected_total": 應解總計
    - "actual_total": 實解總計
    """
    
    cookie_part = {"mime_type": "application/pdf", "data": pdf_bytes}
    resp = model.generate_content([cookie_part, prompt])
    return json.loads(resp.text)

def locate_table_mapping(ws):
    """自動定位公版工作表中的站名列與項目欄位，絕不改動現有公式"""
    station_rows = {}
    col_map = {}
    
    # 1. 尋找各站對應的 Row
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            st_norm = clean_station_name(ws.cell(row=r, column=c).value)
            if st_norm and st_norm not in station_rows:
                station_rows[st_norm] = r
                
    # 2. 尋找表頭欄位（只掃描車站列上方的表頭區域）
    min_st_row = min(station_rows.values()) if station_rows else 5
    for r in range(1, min_st_row):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip().replace(" ", "").replace("\n", "")
            if not val:
                continue
            if "客運" in val:
                col_map.setdefault("passenger", c)
            elif any(x in val for x in ["電腦信用卡", "信用卡", "刷卡"]):
                col_map.setdefault("credit_card", c)
            elif any(x in val for x in ["條碼", "行動支付"]):
                col_map.setdefault("barcode", c)
            elif "週轉金" in val:
                col_map.setdefault("revolving", c)
            elif "支票" in val:
                col_map.setdefault("check", c)
            elif "補繳" in val:
                col_map.setdefault("supplement", c)
            # 現金面額對應
            elif any(x == val or x in val for x in ["2000", "2,000", "貳仟"]):
                col_map.setdefault("cash_2000", c)
            elif any(x == val or x in val for x in ["1000", "1,000", "壹仟", "千元"]):
                col_map.setdefault("cash_1000", c)
            elif any(x == val or x in val for x in ["500", "伍佰"]):
                col_map.setdefault("cash_500", c)
            elif any(x == val or x in val for x in ["200", "貳佰"]):
                col_map.setdefault("cash_200", c)
            elif any(x == val or x in val for x in ["100", "壹佰", "百元"]):
                col_map.setdefault("cash_100", c)
            elif any(x == val or x in val for x in ["50", "伍拾"]):
                col_map.setdefault("cash_50", c)
            elif any(x == val or x in val for x in ["20", "貳拾"]):
                col_map.setdefault("cash_20", c)
            elif any(x == val or x in val for x in ["10", "拾元", "十元"]):
                col_map.setdefault("cash_10", c)
            elif any(x == val or x in val for x in ["5", "伍元", "五元"]):
                col_map.setdefault("cash_5", c)
            elif any(x == val or x in val for x in ["1", "壹元", "一元"]):
                col_map.setdefault("cash_1", c)
            elif any(x in val for x in ["現金小計", "現金合計"]):
                col_map.setdefault("cash_total", c)
                
    return station_rows, col_map

def write_data_to_sheet(ws, data_items):
    """安全寫入數據，保留原儲存格樣式與自帶公式"""
    station_rows, col_map = locate_table_mapping(ws)
    
    def safe_write(r, c, val):
        if not c:
            return
        cell = ws.cell(row=r, column=c)
        # 若原本儲存格已含有公版自帶公式（例如 =SUM(...) 或總計加總），嚴禁覆蓋
        if isinstance(cell.value, str) and cell.value.strip().startswith("="):
            return
        cell.value = val

    for item in data_items:
        st_norm = clean_station_name(item.get("station_name"))
        if not st_norm or st_norm not in station_rows:
            continue
        r = station_rows[st_norm]

        # 1. 客運收入
        safe_write(r, col_map.get("passenger"), item.get("passenger", 0))

        # 2. 電腦信用卡：刷卡(-) 減 退刷(+)，有減法時保留 Excel 算式紀錄
        c_in = item.get("credit_card_in", 0)
        c_ref = item.get("credit_card_refund", 0)
        if c_ref > 0:
            card_val = f"={c_in}-{c_ref}"
        else:
            card_val = c_in if c_in > 0 else 0
        safe_write(r, col_map.get("credit_card"), card_val)

        # 3. 條碼支付：進款(-) 減 退款(+)，有減法時保留 Excel 算式紀錄
        b_in = item.get("barcode_in", 0)
        b_ref = item.get("barcode_refund", 0)
        if b_ref > 0:
            barcode_val = f"={b_in}-{b_ref}"
        else:
            barcode_val = b_in if b_in > 0 else 0
        safe_write(r, col_map.get("barcode"), barcode_val)

        # 4. 其他欄位（週轉金、支票、補繳）
        if item.get("revolving_fund", 0) > 0:
            safe_write(r, col_map.get("revolving"), item.get("revolving_fund"))
        if item.get("check_amount", 0) > 0:
            safe_write(r, col_map.get("check"), item.get("check_amount"))
        if item.get("supplement", 0) > 0:
            safe_write(r, col_map.get("supplement"), item.get("supplement"))

        # 5. 現金面額（自動判斷公版是乘面額張數還是直接填金額）
        use_count = False
        if "cash_total" in col_map:
            total_cell_val = str(ws.cell(row=r, column=col_map["cash_total"]).value or "")
            if "*1000" in total_cell_val or "* 1000" in total_cell_val:
                use_count = True

        for d in [2000, 1000, 500, 200, 100, 50, 20, 10, 5, 1]:
            col_idx = col_map.get(f"cash_{d}")
            if col_idx:
                v = item.get(f"cash_{d}_cnt", 0) if use_count else item.get(f"cash_{d}_amt", 0)
                safe_write(r, col_idx, v if v > 0 else 0)

# --- 主畫面佈局 ---
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("1. 公版 Excel 檔案")
    uploaded_template = st.file_uploader("上傳 `豐富至二水解款單(公版)2.xlsx`", type=["xlsx"])
    
with col_right:
    st.subheader("2. 站務解款單 PDF")
    uploaded_pdfs = st.file_uploader("上傳掃描 PDF（可單選或多選，支援多天批次）", type=["pdf"], accept_multiple_files=True)

# 執行按鈕
if st.button("🚀 開始辨識並自動填入公版", type="primary", use_container_width=True):
    if not api_key:
        st.error("請先在左側填入 Google AI Studio API Key！")
        st.stop()
    if not uploaded_template:
        st.error("請上傳公版 Excel 檔案！")
        st.stop()
    if not uploaded_pdfs:
        st.error("請至少上傳一份 PDF 解款單！")
        st.stop()

    # 載入公版，data_only=False 確保完整保留原有儲存格中的公式與樣式
    template_bytes = io.BytesIO(uploaded_template.read())
    wb = openpyxl.load_workbook(template_bytes, data_only=False)

    progress_bar = st.progress(0)
    status_text = st.empty()
    all_summary_data = []

    for idx, pdf_file in enumerate(uploaded_pdfs):
        status_text.info(f"正在辨識檔案：{pdf_file.name} ...")
        pdf_data = pdf_file.read()
        
        try:
            extracted_items = analyze_pdf_with_gemini(pdf_data, api_key, model_name)
        except Exception as e:
            st.error(f"辨識 {pdf_file.name} 時出錯：{str(e)}")
            continue

        # 自動依據進款日期比對工作表名稱（如 8.28, 28, 0828, 或預設使用選定工作表）
        target_ws = wb.active
        if extracted_items and "date" in extracted_items[0]:
            d_parts = extracted_items[0]["date"].split("-")
            if len(d_parts) == 3:
                m_str, d_str = str(int(d_parts[1])), str(int(d_parts[2]))
                for candidate in [f"{m_str}.{d_str}", d_str, f"{d_parts[1]}{d_parts[2]}", f"{m_str}月{d_str}日"]:
                    if candidate in wb.sheetnames:
                        target_ws = wb[candidate]
                        break

        # 填入公版
        write_data_to_sheet(target_ws, extracted_items)
        all_summary_data.extend(extracted_items)
        progress_bar.progress((idx + 1) / len(uploaded_pdfs))

    status_text.success("🎉 全部解款單辨識並填寫完成！原本公版格式與計算公式已全數保留。")

    # 產出下載檔案
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    st.download_button(
        label="📥 下載填寫完成的解款單 Excel",
        data=output_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # 預覽辨識明細
    if all_summary_data:
        with st.expander("🔍 點擊查看辨識結果與公式明細（核對用）", expanded=False):
            preview_rows = []
            for item in all_summary_data:
                c_in, c_ref = item.get("credit_card_in", 0), item.get("credit_card_refund", 0)
                b_in, b_ref = item.get("barcode_in", 0), item.get("barcode_refund", 0)
                preview_rows.append({
                    "日期": item.get("date"),
                    "車站": item.get("station_name"),
                    "客運(+)": item.get("passenger"),
                    "電腦信用卡儲存格公式": f"={c_in}-{c_ref}" if c_ref > 0 else c_in,
                    "條碼儲存格公式": f"={b_in}-{b_ref}" if b_ref > 0 else b_in,
                    "應解總計": item.get("expected_total"),
                    "實解總計": item.get("actual_total")
                })
            st.dataframe(preview_rows, use_container_width=True)
