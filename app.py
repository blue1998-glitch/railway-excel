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
st.caption("🚀 完整保留原始公式與格式 ｜ ⚡ 高速穩定防 429 機制 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

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

    # 鎖定官方推薦穩定、高辨識率的正式模型，徹底防止 404 錯誤
    curated_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro"
    ]

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=curated_models,
        index=0,
        help="建議預設 gemini-2.5-flash，辨識最清晰且速度最快"
    )

    concurrency = st.slider(
        "⚡ 並行辨識線程數",
        min_value=1,
        max_value=4,
        value=2,
        help="建議設為 2。設為 2~3 既可大幅加速，又能杜絕 429 速率超限錯誤"
    )

    if active_api_key:
        st.success("✅ API 金鑰設定正常")
    else:
        st.warning("⚠️ 請在 secrets 設定或在此輸入 API Key")

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
    station_name: str = Field(description="車站名稱（如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）")
    date_day: int = Field(description="報表『進款日期』中的日/號數 (1 至 31 整數，切勿抓成列印時間)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】的『客運(+)』金額")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】的『貨運(+)』金額")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】的『信用卡刷卡(-)』金額 (填正數)")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】的『信用卡退刷(+)』金額 (無填 0)")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】的『條碼支付進款(-)』金額 (填正數)")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】的『條碼支付退款(+)』金額 (無填 0)")
    other_amount: float = Field(default=0.0, description="左側【應解款數】除上述外的其他項目總和（如存付運費、託收支票、補繳金額、繳回週轉金、零用金等淨額），無填 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】的『應解總計』金額")

def extract_json_str(text: str) -> str:
    """去除 Markdown 標籤以解析 JSON"""
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
    """標準化車站名稱"""
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    """轉為乾淨整數或浮點數"""
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else round(f_val, 2)
    except Exception:
        return 0

def build_deduction_formula(charge_val, refund_val):
    """有退刷/退款時寫入 =進款-退款，純刷卡寫數值"""
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
    """精準鎖定表頭所在列，判斷欄位位置"""
    col_map = {}
    header_row = None
    for r in range(1, 7):
        row_text = "".join([str(sheet.cell(row=r, column=c).value or "") for c in range(1, 15)])
        if "客運" in row_text or ("信用卡" in row_text and "條碼" in row_text):
            header_row = r
            break
            
    scan_rows = [header_row] if header_row else range(1, 6)
    for r in scan_rows:
        for c in range(1, min(sheet.max_column + 1, 15)):
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
            elif ("自輸" in val or "字輸" in val or "應解" in val) and "remittance" not in col_map:
                col_map["remittance"] = c
            elif ("日" in val or "號" in val) and "date" not in col_map and c <= 2:
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
    """支援整數、字串與 Excel datetime 格式的精準列定位"""
    target_day = int(date_day)
    # 優先從指定日期欄掃描
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            if hasattr(val, 'day') and val.day == target_day:
                return r
            try:
                if int(float(val)) == target_day:
                    return r
            except Exception:
                pass
            val_str = str(val).strip()
            m = re.search(r'(?:^|[^\d])0*(\d{1,2})\s*(?:日|號|$)', val_str)
            if m and int(m.group(1)) == target_day:
                return r

    # 備援從第 1 欄掃描
    if date_col != 1:
        for r in range(1, 45):
            val = sheet.cell(row=r, column=1).value
            if val is not None:
                if hasattr(val, 'day') and val.day == target_day:
                    return r
                try:
                    if int(float(val)) == target_day:
                        return r
                except Exception:
                    pass

    # 若未標記日期，預設 Day 1 為 Row 4
    return target_day + 3

def match_station_sheet(wb, raw_name):
    """嚴格防止「新烏日」與「烏日」誤配的車站比對機制"""
    clean_target = clean_station_name(raw_name)
    # 1. 完全相符優先
    for s_name in wb.sheetnames:
        if clean_station_name(s_name) == clean_target:
            return wb[s_name]

    # 2. 烏日系列嚴格隔離
    if "新烏日" in clean_target:
        for s_name in wb.sheetnames:
            if "新烏日" in clean_station_name(s_name):
                return wb[s_name]
    elif "烏日" in clean_target:
        for s_name in wb.sheetnames:
            s_clean = clean_station_name(s_name)
            if "烏日" in s_clean and "新烏日" not in s_clean:
                return wb[s_name]

    # 3. 依字串長度遞減模糊比對（避免短字誤入長字）
    for s_name in sorted(wb.sheetnames, key=len, reverse=True):
        s_clean = clean_station_name(s_name)
        if s_clean and (s_clean in clean_target or clean_target in s_clean):
            return wb[s_name]
    return None

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    """安全寫入指定儲存格（0 與空值不覆蓋）"""
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=5):
    """單頁辨識：具備 404 自動模型降級與 429 抖動指數退避機制"""
    curr_model = model_name
    for retry in range(max_retries):
        try:
            res = client.models.generate_content(
                model=curr_model,
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
            # 自動處理 404 錯誤：自動切換至穩定的 flash 模型
            if "404" in err or "NOT_FOUND" in err:
                if curr_model != "gemini-2.0-flash":
                    curr_model = "gemini-2.0-flash"
                    time.sleep(1.0)
                    continue
                elif curr_model != "gemini-1.5-flash":
                    curr_model = "gemini-1.5-flash"
                    time.sleep(1.0)
                    continue

            # 自動處理 429 / RESOURCE_EXHAUSTED / 伺服器 503：退避時間足以渡過 60 秒視窗
            if "429" in err or "RESOURCE_EXHAUSTED" in err or "503" in err:
                wait_sec = (4.0 * (retry + 1)) + random.uniform(1.0, 3.0)
                time.sleep(wait_sec)
                continue

            if retry == max_retries - 1:
                return None, err
            time.sleep(2.0)
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
# 4. 核心並行辨識與填寫
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
        st.error(f"❌ Excel 讀取失敗：{e}")
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

    status_box.update(label=f"⚡ [階段 2/2] {concurrency} 線程精確辨識 {total_tasks} 頁單據中...", state="running")

    prompt = """
    你是一位極度嚴謹的台鐵會計表單辨識專家。請仔細辨識這張解款單：
    【關鍵欄位擷取原則】：
    1. 車站名稱（station_name）：精準辨識車站名（例如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等），去除站碼與'站'字。特別注意「烏日」與「新烏日」務必精準區分！
    2. 日期（date_day）：單據上方同時有「列印時間」與「進款日期」，請務必擷取【進款日期】中的『日/號』（1至31整數），絕對不可誤抓成列印時間！
    3. 左側【應解款數】大項目區塊金額（請去除千分位逗號，無則填 0）：
       - 客運(+)：passenger_revenue
       - 貨運(+)：freight_revenue
       - 信用卡刷卡(-)：credit_card_charge（填正數）
       - 信用卡退刷(+)：credit_card_refund（仔細看是否有退刷，有則填正數，無填 0）
       - 條碼支付進款(-)：barcode_in（填正數）
       - 條碼支付退款(+)：barcode_refund（有則填正數，無填 0）
       - 應解總計：remittance_total（左側應解總計金額）
       - 其他項目加總：other_amount（除上述項目外，左側【應解款數】內若有存付運費(+)、託收支票、補繳金額(+)、繳回週轉金(+)、繳回零用金(+)等任何明細，請加總淨額計入；無則填 0。例如若有繳回週轉金 200,000，則為 200000）
    4. 勾稽平衡：客運 + 貨運 - (信用卡刷卡 - 信用卡退刷) - (條碼進款 - 條碼退款) + 其他 ＝ 應解總計。請以此原則驗算所有金額！
    """

    results_data = []
    progress_bar = st.progress(0)
    completed_count = 0

    # 執行並行辨識，加入提交微幅間隔避免併發脈衝衝擊 API
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        future_map = {}
        for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1):
            future = executor.submit(call_gemini_page, client, selected_model, page_bytes, prompt)
            future_map[future] = (idx, file_name, p_idx, total_p)
            time.sleep(0.15)  # 關鍵：微幅交錯啟動，杜絕同時併發造成的 429

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
                    status_box.write(f"✅ **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功 - `{file_name}` 第 {p_idx} 頁")
                else:
                    status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁辨識失敗 ({str(err_msg)[:60]})")
            except Exception as e:
                status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁執行異常：{e}")

    results_data.sort(key=lambda x: x[0])

    # ----------------------------------------------------
    # 5. 回填 Excel 與 會計檢核（主線程安全寫入）
    # ----------------------------------------------------
    status_box.update(label="📝 正在將辨識資料寫入 Excel 公版...", state="running")
    total_written = 0
    success_count = 0
    audit_records = []

    for idx, file_name, p_idx, data in results_data:
        target_sheet = match_station_sheet(wb, data.station_name)

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
            "電腦信用卡 (=刷卡-退刷)": cc_formula if cc_formula is not None else 0,
            "條碼 (=進款-退款)": bc_formula if bc_formula is not None else 0,
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
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("credit"), cc_formula)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("barcode"), bc_formula)
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("other"), to_clean_num(data.other_amount))
        total_written += write_cell_if_valid(target_sheet, target_row, col_map.get("remittance"), to_clean_num(data.remittance_total))

        success_count += 1

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
            st.error(f"⚠️ 警告：共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請依下方表格核對單據金額！")
        else:
            st.success("🎯 太棒了！所有辨識成功的單據「自輸 - 總計」皆等於 0，會計平衡完全正確！")

        st.dataframe(df_audit, use_container_width=True, hide_index=True)

    st.download_button(
        label="📥 點擊下載已自動填寫完成的 Excel 報表 (.xlsx)",
        data=out_stream,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    
