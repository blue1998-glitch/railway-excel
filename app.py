import io
import os
import re
import time
import random
import openpyxl
import pandas as pd
from typing import Any, Union
from pydantic import BaseModel, Field, field_validator
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
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 穩健防限流辨識 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

raw_key = st.secrets.get("GEMINI_API_KEY", "")
cleaned_key = str(raw_key).replace('"', '').replace("'", "").strip()

with st.sidebar:
    st.header("⚙️ 系統設定")
    user_key = st.text_input(
        "Gemini API Key",
        value=cleaned_key,
        type="password",
        help="優先讀取 secrets 中的 GEMINI_API_KEY"
    )
    active_api_key = user_key.strip()
    
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
        available_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    else:
        flash_models = [m for m in available_models if "flash" in m.lower()]
        other_models = [m for m in available_models if "flash" not in m.lower()]
        available_models = flash_models + other_models

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=available_models,
        index=0,
        help="推薦使用 Flash 系列模型，辨識速度最快且準確"
    )

    concurrency = st.slider(
        "⚡ 並行辨識線程數",
        min_value=1,
        max_value=5,
        value=2,
        help="建議設定 2~3。適度並行可兼顧高速辨識與穩定度，避免觸發 API 429 頻率上限"
    )

    st.markdown("---")
    st.markdown("""
    💡 **會計計算原則**：
    * **電腦信用卡**：`信用卡刷卡(-)` 減 `信用卡退刷(+)`
    * **條碼**：`條碼支付進款(-)` 減 `條碼支付退款(+)`
    * **平衡檢核**：客運 + 貨運 - 電腦信用卡 - 條碼 + 其他 ＝ 自輸(應解總計)
    """)

# ----------------------------------------------------
# 2. 定義容錯資料結構與工具函式
# ----------------------------------------------------
def parse_clean_num(v: Any) -> float:
    """安全清洗千分位逗號與多餘字符轉為浮點數"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace(" ", "").replace("¥", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0

class StationReport(BaseModel):
    station_name: str = Field(default="", description="車站名稱（如豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）")
    date_day: int = Field(default=0, description="進款日期的『日/號』(1 至 31 整數)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】客運(+)金額")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】貨運(+)金額")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】信用卡刷卡(-)金額")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】信用卡退刷(+)金額")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】條碼支付進款(-)金額")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】條碼支付退款(+)金額")
    other_amount: float = Field(default=0.0, description="左側除上述外的其他項目加總（如存付運費、託收支票、補繳金額、繳回週轉金等淨額）")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】應解總計金額")

    @field_validator(
        "passenger_revenue", "freight_revenue", "credit_card_charge",
        "credit_card_refund", "barcode_in", "barcode_refund",
        "other_amount", "remittance_total",
        mode="before"
    )
    @classmethod
    def clean_amounts(cls, v):
        return parse_clean_num(v)

    @field_validator("date_day", mode="before")
    @classmethod
    def clean_day(cls, v):
        if isinstance(v, str):
            digits = re.findall(r'\d+', v)
            if digits:
                return int(digits[-1])
        try:
            return int(float(v))
        except Exception:
            return 0

def extract_json_str(text: str) -> str:
    """過濾 Markdown 標籤以取得標準 JSON"""
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
    """標準化車站名稱（台/臺互通、去除站碼）"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    try:
        f = float(val)
        return int(f) if f.is_integer() else f
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """扣抵公式：退刷/退款時保留 =刷卡-退刷"""
    c = to_clean_num(charge_val)
    r = to_clean_num(refund_val)
    if c == 0 and r == 0:
        return 0
    if c > 0 and r > 0:
        return f"={c}-{r}"
    if c > 0 and r == 0:
        return c
    if c == 0 and r > 0:
        return f"=-{r}"
    return 0

def get_target_sheet(wb, station_name):
    """精準配對工作表，長度優先防止『新烏日』誤配『烏日』"""
    t_clean = clean_station_name(station_name)
    if not t_clean:
        return None
    for name in wb.sheetnames:
        if clean_station_name(name) == t_clean:
            return wb[name]
    sorted_sheets = sorted(wb.sheetnames, key=lambda x: len(clean_station_name(x)), reverse=True)
    for name in sorted_sheets:
        s_clean = clean_station_name(name)
        if s_clean and (s_clean in t_clean or t_clean in s_clean):
            return wb[name]
    return None

def analyze_sheet_structure(sheet):
    """分析工作表各會計項目所在欄位"""
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
            elif ("自輸" in val or "字輸" in val or "應解總計" in val) and "remittance" not in col_map:
                col_map["remittance"] = c

    defaults = {"passenger": 2, "freight": 3, "credit": 4, "barcode": 5, "other": 6, "remittance": 7}
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return col_map

def find_target_row(sheet, date_day):
    """直接在前 3 欄定位 1~31 的日期列號，不受表頭其他文字干擾"""
    target = int(date_day)
    for c in range(1, 4):
        for r in range(1, 45):
            val = sheet.cell(row=r, column=c).value
            if val is not None:
                try:
                    c_val = str(val).replace("日", "").replace("號", "").strip()
                    if int(float(c_val)) == target:
                        return r
                except Exception:
                    continue
    return target + 1

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=5):
    """呼叫 API 辨識單頁，內建 5 次長退避與隨機抖動重試，抗 429 限流"""
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
                report = StationReport.model_validate_json(clean_text)
                if report.station_name and report.date_day > 0:
                    return report, None
                elif retry == max_retries - 1:
                    return None, "未能識別出有效車站或進款日期"
        except Exception as e:
            err = str(e)
            if "429" in err or "RESOURCE_EXHAUSTED" in err or "quota" in err.lower():
                # 指數退避加上隨機抖動，避免多線程同時重發碰撞
                sleep_sec = (2 ** retry) * 2.5 + random.uniform(0.5, 2.0)
                time.sleep(sleep_sec)
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
# 4. 核心辨識與填寫
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

    status_box.update(label=f"⚡ [階段 2/2] 啟動 {concurrency} 線程精準辨識 {total_tasks} 頁單據...", state="running")
    
    prompt = """
    你是一位專業嚴謹的台鐵會計表單辨識審核員。請仔細辨識這張台鐵站務解款單（23_R_20）：
    1. 車站名稱：位於左上方（如 3150_豐富站、3160_苗栗站、3300_臺中站等，僅需填寫站名如「豐富」、「苗栗」、「台中」）。
    2. 進款日期：位於上方或右上（如「進款日期:2026年08月28日」），擷取其日數（1~31 的整數）。
    3. 專注看左側【應解款數】大欄位（切勿看右側實解款數的鈔票張數）：
       - 客運(+) 金額
       - 貨運(+) 金額（若無填 0）
       - 信用卡刷卡(-) 金額（填正數，無則填 0）
       - 信用卡退刷(+) 金額（若有退刷務必填正數，無退刷填 0）
       - 條碼支付進款(-) 金額（填正數，無則填 0）
       - 條碼支付退款(+) 金額（若有退款務必填正數，無退款填 0）
       - 其他項目：存付運費(+)、託收支票(-)、補繳金額(+)、繳回週轉金(+)等所有非上述四項的其餘項目淨額加總
       - 應解總計：左側最下方的應解總計金額
    4. 會計平衡檢核式：
       客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計。
       請務必以此勾稽原則覆核數值是否完整正確！
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
                text=f"辨識處理中：已完成 {completed_count}/{total_tasks} 頁 ({completed_count * 100 // total_tasks}%)"
            )
            
            try:
                parsed_data, err_msg = future.result()
                if parsed_data and parsed_data.date_day > 0:
                    results_data.append((idx, file_name, p_idx, parsed_data))
                    status_box.write(f"✅ **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功 - `{file_name}` 第 {p_idx} 頁")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗：{err_msg}")
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁處理異常：{e}")

    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel 與會計平衡檢核
    # ----------------------------------------------------
    status_box.update(label="📝 正在將辨識結果精確寫入 Excel 公版儲存格...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_sheet = get_target_sheet(wb, data.station_name)

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
            "電腦信用卡": cc_formula,
            "條碼": bc_formula,
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
        target_row = find_target_row(target_sheet, data.date_day)

        # 寫入各欄位儲存格
        fields_to_write = [
            (col_map.get("passenger"), to_clean_num(data.passenger_revenue)),
            (col_map.get("freight"), to_clean_num(data.freight_revenue)),
            (col_map.get("credit"), cc_formula),
            (col_map.get("barcode"), bc_formula),
            (col_map.get("other"), to_clean_num(data.other_amount)),
            (col_map.get("remittance"), to_clean_num(data.remittance_total))
        ]

        for col_idx, val in fields_to_write:
            if col_idx and val is not None and val != "":
                target_sheet.cell(row=target_row, column=col_idx, value=val)
                total_written += 1

        success_count += 1

    status_box.update(label="🎉 辨識與 Excel 自動填表全數完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！總耗時 {elapsed:.1f} 秒，共成功處理 {success_count}/{total_tasks} 頁解款單，成功填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核儀表板與檔案下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.error(f"⚠️ 注意：共有 **{unbalanced_count}** 筆單據計算不平衡，請檢視下方表格確認單據細項！")
        else:
            st.success("🎯 太棒了！所有辨識成功的單據會計平衡檢核完全吻合（差額為 0）！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
                 )
    
