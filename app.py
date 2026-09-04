import io
import json
import os
import re
import tempfile
import google.generativeai as genai
import openpyxl
import pandas as pd
import streamlit as st

# 頁面配置
st.set_page_config(page_title="台鐵解款單自動彙總系統", layout="wide")
st.title("🚆 台鐵站務解款單 自動彙總填報系統")

# 側邊欄設定
st.sidebar.header("🔑 金鑰與模型設定")
api_key = st.sidebar.text_input(
    "Google Gemini API Key",
    value=st.secrets.get("GEMINI_API_KEY", ""),
    type="password",
    help="請輸入 Google AI Studio 申請的 API Key",
)
model_name = st.sidebar.selectbox("Gemini 模型", ["gemini-2.5-flash", "gemini-1.5-flash"], index=0)

# 標準車站清單（豐富至二水）
STATIONS = [
    "豐富", "苗栗", "銅鑼", "三義", "泰安", "后里", "豐原", "栗林",
    "潭子", "頭家厝", "松竹", "太原", "精武", "臺中", "五權", "大慶",
    "烏日", "新烏日", "成功", "彰化", "花壇", "大村", "員林", "社頭",
    "田中", "二水"
]

AI_PROMPT = """
你是一個專業的台鐵會計單據辨識專家。請分析這份「站務解款單」掃描文件（每一頁為一個車站）。
請逐頁辨識，並輸出為標準 JSON 陣列格式。

注意細節：
1. 站名請去除代號與「站」字（例如「3150_豐富站」請提取「豐富」；「臺中」或「台中」請統一輸出為「臺中」）。
2. 日期請務必提取「進款日期」（格式 YYYY-MM-DD），不要抓列印時間或簽核時間。
3. 客運(+)金額請轉為正整數（若有客運[離線](+)請相加）。
4. 信用卡刷卡(-) 與 信用卡退刷(+) 必須分開提取為正整數（無退刷請填 0）。
5. 條碼支付進款(-) 與 條碼支付退款(+) 必須分開提取為正整數（無退款請填 0）。
6. 若有「繳回週轉金(+)」請提取，無則填 0。
7. 貨運(+)、存付運費(+)、託收支票、補繳金額等若有請提取。

輸出格式範例（純 JSON 陣列，不要多餘註解）：
[
  {
    "station_name": "豐富",
    "date": "2026-08-28",
    "passenger": 26491,
    "credit_charge": 2233,
    "credit_refund": 0,
    "barcode_charge": 287,
    "barcode_refund": 0,
    "revolving_fund": 0,
    "total_due": 23971,
    "total_actual": 23971
  }
]
"""

def extract_pdf_data(pdf_file, key, model_choice):
    """呼叫 Gemini 視覺 API 提取 PDF 內解款單數據"""
    genai.configure(api_key=key)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    try:
        uploaded_file = genai.upload_file(tmp_path, mime_type="application/pdf")
        model = genai.GenerativeModel(
            model_name=model_choice,
            generation_config={"response_mime_type": "application/json"}
        )
        response = model.generate_content([uploaded_file, AI_PROMPT])
        data = json.loads(response.text)
        return data
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

def find_target_sheet(wb, date_str):
    """根據進款日期配對工作表（支援 28、8.28、8月28日、0828 等命名）"""
    parts = date_str.split("-")
    if len(parts) == 3:
        month, day = str(int(parts[1])), str(int(parts[2]))
        day_pad = f"{int(day):02d}"
        candidates = [
            day, f"{day}日", f"{month}.{day}", f"{month}/{day}",
            f"{month}月{day}日", f"{int(month):02d}{day_pad}", date_str
        ]
        for cand in candidates:
            if cand in wb.sheetnames:
                return wb[cand]
        for name in wb.sheetnames:
            for cand in candidates:
                if cand in name:
                    return wb[name]
    return wb.active

def get_mapping(ws):
    """掃描公版定位各車站行數與科目欄位，不破壞原有儲存格"""
    station_rows = {}
    col_map = {}

    # 1. 定位車站行號
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 8)):
            val = str(ws.cell(row=r, column=c).value or "").strip().replace(" ", "")
            if not val:
                continue
            for st_name in STATIONS:
                matched = False
                if st_name == "臺中":
                    if "臺中" in val or "台中" in val:
                        matched = True
                elif st_name in val:
                    matched = True
                if matched and st_name not in station_rows:
                    station_rows[st_name] = r
                    break

    # 2. 定位標題欄號
    min_row = min(station_rows.values()) if station_rows else 6
    for r in range(1, min_row):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip().replace(" ", "").replace("\n", "")
            if "客運" in val and "客運" not in col_map:
                col_map["客運"] = c
            elif ("信用卡" in val or "電腦信用卡" in val) and "信用卡" not in col_map:
                col_map["信用卡"] = c
            elif "條碼" in val and "條碼" not in col_map:
                col_map["條碼"] = c
            elif ("週轉金" in val or "周轉金" in val) and "週轉金" not in col_map:
                col_map["週轉金"] = c

    return station_rows, col_map

# 檔案上傳介面
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader(
        "1. 上傳公版 Excel 範本 (.xlsx)",
        type=["xlsx"],
        help="請務必使用 .xlsx 格式以保留原表內的所有公式與排版格式！"
    )
with col2:
    pdf_files = st.file_uploader(
        "2. 上傳掃描解款單 PDF（可多選）",
        type=["pdf"],
        accept_multiple_files=True
    )

if st.button("🚀 開始辨識並寫入公版", type="primary"):
    if not api_key:
        st.error("請在側邊欄填入 Google Gemini API Key！")
        st.stop()
    if not template_file or not pdf_files:
        st.error("請同時上傳「公版範本」與「解款單 PDF」！")
        st.stop()

    # 載入公版，保留所有原本公式與樣式
    wb = openpyxl.load_workbook(template_file, data_only=False)

    all_results = []
    progress_bar = st.progress(0)

    for idx, pdf in enumerate(pdf_files):
        st.write(f"正在分析單據：**{pdf.name}**...")
        try:
            records = extract_pdf_data(pdf, api_key, model_name)
        except Exception as e:
            st.error(f"辨識 {pdf.name} 失敗：{e}")
            continue

        for item in records:
            st_name = item.get("station_name", "").replace("站", "").strip()
            date_str = item.get("date", "")
            passenger = item.get("passenger", 0)
            c_charge = item.get("credit_charge", 0)
            c_refund = item.get("credit_refund", 0)
            b_charge = item.get("barcode_charge", 0)
            b_refund = item.get("barcode_refund", 0)
            revolving = item.get("revolving_fund", 0)

            # 計算公式生成：有退刷/退款時保留運算式（例如 =100-10）
            if c_refund > 0:
                credit_val = f"={c_charge}-{c_refund}"
            elif c_charge > 0:
                credit_val = c_charge
            else:
                credit_val = 0

            if b_refund > 0:
                barcode_val = f"={b_charge}-{b_refund}"
            elif b_charge > 0:
                barcode_val = b_charge
            else:
                barcode_val = 0

            # 寫入 Excel（僅更動對應儲存格）
            ws = find_target_sheet(wb, date_str)
            station_rows, col_map = get_mapping(ws)

            # 容錯比對臺中
            matched_st = "臺中" if st_name in ["臺中", "台中"] else st_name
            row_idx = station_rows.get(matched_st)

            if row_idx:
                if "客運" in col_map:
                    ws.cell(row=row_idx, column=col_map["客運"], value=passenger)
                if "信用卡" in col_map:
                    ws.cell(row=row_idx, column=col_map["信用卡"], value=credit_val)
                if "條碼" in col_map:
                    ws.cell(row=row_idx, column=col_map["條碼"], value=barcode_val)
                if "週轉金" in col_map and revolving > 0:
                    ws.cell(row=row_idx, column=col_map["週轉金"], value=revolving)

            all_results.append({
                "來源檔案": pdf.name,
                "進款日期": date_str,
                "車站": st_name,
                "客運(+)": passenger,
                "電腦信用卡(算式/值)": credit_val,
                "條碼支付(算式/值)": barcode_val,
                "繳回週轉金": revolving,
                "狀態": "已填入" if row_idx else "找不到對應行"
            })

        progress_bar.progress((idx + 1) / len(pdf_files))

    # 輸出修改後的公版檔案
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    st.success("🎉 所有單據處理完成！公版原有之總計與公式已完整保留。")

    st.download_button(
        label="📥 下載彙總完成表 (保留公版公式與格式)",
        data=output_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 顯示辨識明細預覽
    if all_results:
        st.subheader("📋 辨識與寫入明細預覽")
        st.dataframe(pd.DataFrame(all_results), use_container_width=True)
