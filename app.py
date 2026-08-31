import os
import io
import time
import subprocess
import streamlit as st
import openpyxl
from pydantic import BaseModel, Field
from typing import List
from google import genai
from google.genai import types

# 網頁基本設定
st.set_page_config(page_title="台鐵解款單自動填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵解款單 ➜ Excel 智慧自動填表系統")

# 讀取 API 金鑰
api_key = st.secrets.get("GEMINI_API_KEY", "")
with st.sidebar:
    st.header("⚙️ 系統設定")
    if not api_key:
        api_key = st.text_input("請輸入 Gemini API Key", type="password")
    else:
        st.success("✅ 已自動載入雲端 API Key")

# 資料模型
class StationReport(BaseModel):
    page_number: int = Field(description="PDF 頁碼 (從 1 開始)")
    station_name: str = Field(description="車站名稱")
    date_day: int = Field(description="日期幾號 (1-31)")
    passenger_revenue: float = Field(default=0.0, description="客運金額")
    freight_revenue: float = Field(default=0.0, description="貨運金額")
    credit_card_pos: float = Field(default=0.0, description="信用卡刷卡(+)")
    credit_card_neg: float = Field(default=0.0, description="信用卡刷卡(-)")
    barcode_in: float = Field(default=0.0, description="條碼進款")
    barcode_refund: float = Field(default=0.0, description="條碼退款")
    other_amount: float = Field(default=0.0, description="其他獨立明細")
    remittance_total: float = Field(default=0.0, description="應解總計")

class MultiPageDocument(BaseModel):
    station_reports: List[StationReport] = Field(description="各頁車站清單")

def clean_name(val):
    if not val: return ""
    return str(val).replace("臺", "台").replace("站", "").replace(" ", "").strip()

def to_num(val):
    try:
        f = float(val)
        return int(f) if f.is_integer() else f
    except Exception:
        return 0

def build_formula(pos, neg):
    p, n = to_num(pos), to_num(neg)
    if p == 0 and n == 0: return 0
    if n > 0: return f"={p}-{n}"
    return p

# 介面上傳
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上傳 Excel 公版 (.xlsx/.xls)", type=["xlsx", "xls"])
with col2:
    uploaded_pdfs = st.file_uploader("2. 批次上傳掃描 PDF (可多選)", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 開始自動辨識與填表", type="primary", use_container_width=True):
    if not api_key:
        st.error("❌ 請提供 Gemini API Key！")
        st.stop()
    if not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請上傳 Excel 公版與至少一個 PDF 檔案！")
        st.stop()

    client = genai.Client(api_key=api_key)
    progress_bar = st.progress(0)
    status_text = st.empty()

    # 處理 Excel 格式
    excel_bytes = uploaded_excel.getvalue()
    if uploaded_excel.name.lower().endswith(".xls"):
        with open("temp_template.xls", "wb") as f:
            f.write(excel_bytes)
        subprocess.run(["libreoffice", "--headless", "--convert-to", "xlsx", "temp_template.xls"], check=True)
        wb = openpyxl.load_workbook("temp_template.xlsx")
    else:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))

    prompt = """
    你是一個精確的台鐵會計表單辨識助理。這份 PDF 包含多頁每日解款單。
    請逐頁辨識：頁碼(page_number)、車站名稱(station_name)、日期幾號(date_day)。
    左邊【應解款數】各項金額：客運、貨運、信用卡刷卡(+)、信用卡刷卡(-)、條碼進款、條碼退款、其他獨立明細、應解總計。
    注意：嚴禁將刷卡(-)或條碼退款重複加進其他！各項加總必須平衡。
    """

    models = ["gemini-2.5-flash", "gemini-3.6-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    total_files = len(uploaded_pdfs)
    success_pages = 0
    logs = []

    for f_idx, pdf in enumerate(uploaded_pdfs, 1):
        status_text.info(f"⏳ 正在分析第 {f_idx}/{total_files} 份檔案: `{pdf.name}`")
        temp_path = f"temp_upload_{f_idx}.pdf"
        with open(temp_path, "wb") as f:
            f.write(pdf.getvalue())

        try:
            cloud_file = client.files.upload(file=temp_path)
            doc_data = None
            for m in models:
                try:
                    res = client.models.generate_content(
                        model=m,
                        contents=[cloud_file, prompt],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=MultiPageDocument,
                            temperature=0.0
                        )
                    )
                    if res and res.text:
                        doc_data = MultiPageDocument.model_validate_json(res.text)
                        break
                except Exception:
                    time.sleep(2)
                    continue

            if not doc_data:
                logs.append(f"❌ `{pdf.name}` 辨識失敗")
                continue

            for rep in doc_data.station_reports:
                if rep.date_day <= 0: continue
                clean_target = clean_name(rep.station_name)
                sheet = next((wb[s] for s in wb.sheetnames if clean_target in clean_name(s) or clean_name(s) in clean_target), None)
                if not sheet:
                    logs.append(f"⚠️ 找不到車站分頁 `[{rep.station_name}]`")
                    continue

                # 平帳校核
                calc_cc = rep.credit_card_pos - rep.credit_card_neg
                calc_bc = rep.barcode_in - rep.barcode_refund
                base = rep.passenger_revenue + rep.freight_revenue + calc_cc + calc_bc
                
                final_other = rep.other_amount
                if abs(rep.remittance_total - base) < 0.01:
                    final_other = 0.0
                elif abs(rep.remittance_total - (base + rep.other_amount)) > 0.01:
                    final_other = rep.remittance_total - base

                row = int(rep.date_day) + 1
                def write_c(c, v):
                    if v != 0 and v != "0": sheet.cell(row=row, column=c, value=v)

                write_c(2, to_num(rep.passenger_revenue))
                write_c(3, to_num(rep.freight_revenue))
                write_c(4, build_formula(rep.credit_card_pos, rep.credit_card_neg))
                write_c(5, build_formula(rep.barcode_in, rep.barcode_refund))
                write_c(6, to_num(final_other))
                write_c(7, to_num(rep.remittance_total))

                success_pages += 1
                logs.append(f"✅ **{rep.station_name}** ({rep.date_day} 號) -> 寫入第 {row} 列（已平衡）")

        except Exception as e:
            logs.append(f"❌ 處理 `{pdf.name}` 出錯: {e}")

        progress_bar.progress(f_idx / total_files)

    status_text.empty()
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)

    st.success(f"🎉 填表完成！成功處理 {success_pages} 頁單據！")
    st.download_button(
        label="📥 點擊下載完成的 Excel 報表",
        data=out_stream,
        file_name="完成_解款單彙總報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

    with st.expander("🔍 檢視詳細日誌", expanded=True):
        for log in logs: st.markdown(log)
          
