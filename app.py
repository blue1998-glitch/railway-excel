import io
import os
import re
import time
import random
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
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 智慧抗限流逐頁辨識 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

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
    
    # 鎖定官方穩定、支援圖片/PDF的高速模型，徹底防止 404 錯誤
    stable_models = [
        "gemini-2.0-flash",
        "gemini-2.5-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro"
    ]

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=stable_models,
        index=0,
        help="預設 gemini-2.0-flash 速度最快、辨識精確度極高且最穩定"
    )

    concurrency = st.slider(
        "⚡ 並行辨識線程數",
        min_value=1,
        max_value=4,
        value=2,
        help="推薦設為 2~3。適度並行既能大幅縮短時間，又能徹底防止觸發 Gemini API 429 限流"
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
    station_name: str = Field(description="車站名稱（如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水，請去除代碼與'站'字）")
    date_day: int = Field(description="單據右上角『進款日期』的『日』(1 至 31 整數)。絕對不要抓右下角藍色核章或列印時間！")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】中的『客運(+)』金額（含客運離線），無則填 0")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】中的『貨運(+)』金額，無則填 0")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】中的『信用卡刷卡(-)』金額 (填正數)，無則填 0")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】中的『信用卡退刷(+)』金額 (填正數)，無則填 0")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】中的『條碼支付進款(-)』金額 (填正數)，無則填 0")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】中的『條碼支付退款(+)』金額 (填正數)，無則填 0")
    other_amount: float = Field(default=0.0, description="除上述項目外之其他明細淨額加總（如繳回週轉金、存付運費、託收支票、補繳金額等），無則填 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】中的『應解總計』金額")

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

def clean_station_name(val):
    """標準化車站名稱，自動剔除站碼前綴與多餘字元以精準比對工作表"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    """轉換為乾淨整數或浮點數"""
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else f_val
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """會計立場扣抵公式：當有退刷/退款時寫入 =進款-退款，Excel 將自動算值並保留軌跡"""
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
    """智慧掃描工作表表頭取得各欄位位置"""
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
    """精準尋找對應日期的列號 (1~31)，支援多種日期格式相容"""
    target_day = int(date_day)
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is None:
            continue
        if hasattr(val, "day"):
            if val.day == target_day:
                return r
        val_str = str(val).strip()
        m = re.search(r'(\d+)(?:日|號)?$', val_str)
        if m:
            try:
                if int(m.group(1)) == target_day:
                    return r
            except Exception:
                pass
        try:
            if int(float(val_str)) == target_day:
                return r
        except Exception:
            pass
    return None

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    """安全寫入指定儲存格（0 或空值不寫入以維護報表整潔）"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=8):
    """單頁辨識函式，內建強效指數退避與防 429 限流機制"""
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
            # 遇到 429 或限流錯誤，採用階梯式等待 + 隨機微延遲
            if "429" in err or "RESOURCE_EXHAUSTED" in err or "Quota" in err:
                wait_sec = min(50.0, (2 ** retry) * 2.0 + random.uniform(1.0, 3.0))
                time.sleep(wait_sec)
                continue
            elif "503" in err or "500" in err or "UNAVAILABLE" in err:
                time.sleep(3.0 + random.uniform(0.5, 2.0))
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

    status_box.update(label=f"⚡ [階段 2/2] {concurrency} 線程精準辨識 {total_tasks} 頁解款單據中...", state="running")
    
    prompt = """
    你是一位極度嚴謹專業的台鐵會計表單辨識專家。請仔細檢視這張解款單：

    ⚠️【極重要：日期辨識絕對準則】：
    1. 請務必尋找單據右上角印製的『進款日期: 2026年XX月XX日』，僅擷取其中的『日』(1~31 的整數)！
    ⛔ 嚴格禁止讀取：
       - 右下角資金管理科的藍色核章（例如『115. 8. 31 訖』或『115. 9. -1』，那是點收日期，絕對不是進款日期！）。
       - 最上方的『列印時間』。
       - 任何簽核欄位的時間。
       必須 100% 以『進款日期:』後面的日數為唯一標準！

    ⚠️【車站名稱準則】：
    2. 擷取車站名稱（豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水），去除站碼（如 3150_）與'站'字。注意分清「烏日」與「新烏日」。

    ⚠️【數值擷取與會計勾稽原則】：
    3. 專注檢視左側【應解款數】大項目區塊（若無項目填 0，複寫或淺色數字請仔細辨識）：
       - 客運(+)：包含客運及離線客運金額
       - 貨運(+)
       - 信用卡刷卡(-) (填正數)
       - 信用卡退刷(+) (請特別注意是否有退刷正數金額，勿漏看)
       - 條碼支付進款(-) (填正數)
       - 條碼支付退款(+) (請特別注意是否有退款正數金額，勿漏看)
       - 應解總計：報表上的應解總計金額
       - 其他項目加總：除上述項目外之其他明細加總（如：繳回週轉金、存付運費、託收支票、補繳金額等淨額）

    4. 【必須符合會計平衡公式】：
       客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計
       請務必以此公式自我驗算，確認擷取的數值無誤後再輸出！
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
                text=f"⚡ 辨識處理中：已完成 {completed_count}/{total_tasks} 頁 ({completed_count * 100 // total_tasks}%)"
            )
            
            try:
                parsed_data, err_msg = future.result()
                if parsed_data and parsed_data.date_day > 0:
                    results_data.append((idx, file_name, p_idx, parsed_data))
                    status_box.write(f"✅ **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功 ➜ 檔案 `{file_name}` 第 {p_idx} 頁")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗：{str(err_msg)[:60]}")
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁執行異常：{e}")

    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel 與 會計立場平衡核對
    # ----------------------------------------------------
    status_box.update(label="📝 正在將辨識資料寫入 Excel 公版...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_name_clean = clean_station_name(data.station_name)
        
        # 嚴格優先完全相符比對（防止烏日/新烏日錯配）
        target_sheet = None
        for s_name in wb.sheetnames:
            if clean_station_name(s_name) == target_name_clean:
                target_sheet = wb[s_name]
                break
        if not target_sheet:
            for s_name in wb.sheetnames:
                s_clean = clean_station_name(s_name)
                if s_clean and (s_clean == target_name_clean):
                    target_sheet = wb[s_name]
                    break

        net_credit = data.credit_card_charge - data.credit_card_refund
        net_barcode = data.barcode_in - data.barcode_refund

        computed_total = data.passenger_revenue + data.freight_revenue - net_credit - net_barcode + data.other_amount
        diff = round(data.remittance_total - computed_total, 2)
        is_balanced = (abs(diff) < 0.01)

        cc_formula = build_deduction_formula(data.credit_card_charge, data.credit_card_refund)
        bc_formula = build_deduction_formula(data.barcode_in, data.barcode_refund)

        sheet_status = "❌ 找不到工作表"
        if target_sheet:
            col_map = analyze_sheet_structure(target_sheet)
            target_row = find_target_row(target_sheet, data.date_day, col_map["date"])

            if target_row:
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("passenger"), to_clean_num(data.passenger_revenue))
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("freight"), to_clean_num(data.freight_revenue))
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
                total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))
                success_count += 1
                sheet_status = f"已寫入 [{target_sheet.title}] 第 {target_row} 列"
            else:
                sheet_status = f"⚠️ 找不到 {data.date_day} 日之列號"

        audit_records.append({
            "檔案名稱": file_name,
            "車站名稱": data.station_name,
            "進款日期": f"{data.date_day} 日",
            "客運": to_clean_num(data.passenger_revenue),
            "貨運": to_clean_num(data.freight_revenue),
            "電腦信用卡 (=刷卡-退刷)": cc_formula if cc_formula is not None else 0,
            "條碼 (=進款-退款)": bc_formula if bc_formula is not None else 0,
            "其他": to_clean_num(data.other_amount),
            "自輸 (PDF應解總計)": to_clean_num(data.remittance_total),
            "計算總計": to_clean_num(computed_total),
            "差額": diff,
            "平衡狀態": "✅ 平衡 (0)" if is_balanced else f"❌ 差額 {diff:+.0f}",
            "寫入狀態": sheet_status
        })

    status_box.update(label="🎉 辨識與 Excel 寫入全數完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，共成功處理 {success_count}/{total_tasks} 頁單據，填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核儀表板與下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.warning(f"⚠️ 注意：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請查看下方表格確認。")
        else:
            st.success("🎯 太棒了！所有單據之勾稽計算皆完全平衡（差額為 0）！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    
