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

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰讀取
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🚀 完整保留原始公式與樣式 ｜ 🛡️ 智能限流防 429 封鎖 ｜ ⚖️ 會計扣抵公式 (=刷卡-退刷) ｜ 🔍 自動平衡檢核")

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

    # 鎖定正式、穩定、相容度最高的多模態模型
    model_options = [
        "gemini-1.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-pro"
    ]

    selected_model = st.selectbox(
        "AI 辨識核心模型",
        options=model_options,
        index=0,
        help="預設 gemini-1.5-flash 擁有最高配額與最穩定的視覺辨識力"
    )

    # 金鑰額度健康檢查按鈕
    if st.button("🔍 檢測目前 API Key 額度狀態", use_container_width=True):
        if not active_api_key:
            st.warning("請先輸入 API Key！")
        else:
            with st.spinner("連線 Google 伺服器驗證中..."):
                try:
                    chk_client = genai.Client(api_key=active_api_key)
                    chk_res = chk_client.models.generate_content(
                        model=selected_model,
                        contents="測試連線，請回覆 OK"
                    )
                    st.success("✅ 金鑰額度完全正常，可正常處理！")
                except Exception as e:
                    chk_err = str(e)
                    if "429" in chk_err or "RESOURCE_EXHAUSTED" in chk_err:
                        st.error("❌ 此 API Key 的免費配額已耗盡（429）。請更換一組新 Google 帳號申請的 API Key！")
                    elif "404" in chk_err:
                        st.error("❌ 此金鑰不支援此模型名稱，請切換其他模型！")
                    else:
                        st.error(f"連線異常：{chk_err[:100]}")

    st.markdown("---")
    st.markdown("""
    💡 **會計計算原則**：
    * **電腦信用卡**：`信用卡刷卡(-)` 減 `信用卡退刷(+)`
    * **條碼**：`條碼支付進款(-)` 減 `條碼支付退款(+)`
    * **平衡檢核**：客運 + 貨運 - 電腦信用卡 - 條碼 + 其他 ＝ 自輸(應解總計)
    """)

# ----------------------------------------------------
# 2. 資料結構與工具函式
# ----------------------------------------------------
class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱（如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）")
    date_day: int = Field(description="報表上方『進款日期』中的日/號數 (1 至 31 整數，切勿抓成列印時間)")
    passenger_revenue: float = Field(default=0.0, description="左側【應解款數】的『客運(+)』金額")
    freight_revenue: float = Field(default=0.0, description="左側【應解款數】的『貨運(+)』金額")
    credit_card_charge: float = Field(default=0.0, description="左側【應解款數】的『信用卡刷卡(-)』金額 (填正數)")
    credit_card_refund: float = Field(default=0.0, description="左側【應解款數】的『信用卡退刷(+)』金額 (無填 0)")
    barcode_in: float = Field(default=0.0, description="左側【應解款數】的『條碼支付進款(-)』金額 (填正數)")
    barcode_refund: float = Field(default=0.0, description="左側【應解款數】的『條碼支付退款(+)』金額 (無填 0)")
    other_amount: float = Field(default=0.0, description="左側【應解款數】除上述外的其他項目總額（如存付運費、補繳、週轉金等），無填 0")
    remittance_total: float = Field(default=0.0, description="左側【應解款數】的『應解總計』金額")

def extract_json_str(text: str) -> str:
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
    if not val:
        return ""
    s = str(val).replace("臺", "台").replace("站", "").strip()
    s = re.sub(r'^\d+[_ ]*', '', s)
    return s.replace(" ", "").replace("　", "").strip()

def to_clean_num(val):
    try:
        f_val = float(val)
        return int(f_val) if f_val.is_integer() else round(f_val, 2)
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
    col_map = {}
    for r in range(1, 6):
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
            elif ("自輸" in val or "字輸" in val) and "remittance" not in col_map:
                col_map["remittance"] = c
            elif ("日" in val or "號" in val) and "date" not in col_map and c <= 2:
                col_map["date"] = c

    defaults = {"date": 1, "passenger": 2, "freight": 3, "credit": 4, "barcode": 5, "other": 6, "remittance": 7}
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
    return col_map

def find_target_row(sheet, date_day, date_col=1):
    target_day = int(date_day)
    for r in range(1, 45):
        val = sheet.cell(row=r, column=date_col).value
        if val is not None:
            if hasattr(val, 'day') and val.day == target_day:
                return r
            try:
                val_str = str(val).replace("日", "").replace("號", "").strip()
                if int(float(val_str)) == target_day:
                    return r
            except Exception:
                pass
    return target_day + 1

def match_station_sheet(wb, raw_name):
    clean_target = clean_station_name(raw_name)
    # 1. 完全符合優先
    for s_name in wb.sheetnames:
        if clean_station_name(s_name) == clean_target:
            return wb[s_name]

    # 2. 烏日與新烏日隔離
    if "新烏日" in clean_target:
        for s_name in wb.sheetnames:
            if "新烏日" in clean_station_name(s_name):
                return wb[s_name]
    elif "烏日" in clean_target:
        for s_name in wb.sheetnames:
            s_clean = clean_station_name(s_name)
            if "烏日" in s_clean and "新烏日" not in s_clean:
                return wb[s_name]

    # 3. 依長度模糊比對
    for s_name in sorted(wb.sheetnames, key=len, reverse=True):
        s_clean = clean_station_name(s_name)
        if s_clean and (s_clean in clean_target or clean_target in s_clean):
            return wb[s_name]
    return None

def write_cell_if_valid(sheet, row_idx, col_idx, val):
    if val is not None and val != 0 and val != "0" and val != "":
        if col_idx:
            sheet.cell(row=row_idx, column=col_idx, value=val)
            return 1
    return 0

def call_gemini_page(client, model_name, page_bytes, prompt, max_retries=4):
    """單頁辨識：嚴格退避機制"""
    last_err = ""
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
            last_err = str(e)

            if "404" in last_err or "NOT_FOUND" in last_err:
                return None, "模型名稱不支援 (404)，請在左側選單切換模型"

            # 429 速率限制退避：等待足夠秒數恢復配額
            if "429" in last_err or "RESOURCE_EXHAUSTED" in last_err:
                wait_sec = 10.0 + (retry * 8.0)
                time.sleep(wait_sec)
                continue

            time.sleep(2.0)

    return None, f"額度不足或超限 (429): 請更換新金鑰後再試"

# ----------------------------------------------------
# 3. 介面上傳區塊
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳公版 Excel 檔案 (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF 解款單 (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 核心辨識與填寫（穩健依序處理，徹底不卡 429）
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

    status_box.update(label=f"⚡ [階段 2/2] 使用 [{selected_model}] 依序平穩辨識 {total_tasks} 頁單據中...", state="running")

    prompt = """
    你是一位專業精確的台鐵會計表單辨識專家。請仔細檢視這張解款單：
    1. 車站名稱（station_name）：辨識車站名（如：豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水等，去除站碼與'站'字）。請特別精準區分「烏日」與「新烏日」！
    2. 日期（date_day）：單據上方有「列印時間」與「進款日期」，務必擷取【進款日期】中的『日/號』（1至31整數），切勿誤抓成列印時間！
    3. 左側【應解款數】金額（無則填 0）：
       - 客運(+)：passenger_revenue
       - 貨運(+)：freight_revenue
       - 信用卡刷卡(-)：credit_card_charge（填正數）
       - 信用卡退刷(+)：credit_card_refund（有則填正數，無填 0）
       - 條碼支付進款(-)：barcode_in（填正數）
       - 條碼支付退款(+)：barcode_refund（有則填正數，無填 0）
       - 應解總計：remittance_total
       - 其他項目加總：other_amount（除上述外，左側存付運費、補繳、週轉金等其他項目淨額加總，無填 0）
    4. 勾稽驗算：客運 + 貨運 - (信用卡刷卡 - 信用卡退刷) - (條碼進款 - 條碼退款) + 其他 ＝ 應解總計。請務必驗算確認！
    """

    results_data = []
    progress_bar = st.progress(0)
    failed_consecutive = 0

    # 採用受控平滑步調依序發送，徹底避開 15 RPM 上限
    for idx, (file_name, p_idx, total_p, page_bytes) in enumerate(all_pages, 1):
        parsed_data, err_msg = call_gemini_page(client, selected_model, page_bytes, prompt)

        if parsed_data and parsed_data.date_day > 0:
            failed_consecutive = 0
            results_data.append((idx, file_name, p_idx, parsed_data))
            status_box.write(f"✅ **{parsed_data.station_name}**（{parsed_data.date_day} 日）辨識成功 - `{file_name}` 第 {p_idx} 頁")
        else:
            failed_consecutive += 1
            status_box.write(f"⚠️ `{file_name}` 第 {p_idx} 頁：{err_msg}")
            # 若連續 2 頁都因額度上限失敗，代表金鑰額度已被 Google 鎖死，立即停止避免徒勞耗時
            if failed_consecutive >= 2 and "429" in str(err_msg):
                status_box.update(label="❌ 金鑰額度已完全耗盡，請更換 API Key", state="error")
                st.error("🚨 偵測到您的 Gemini API Key 當日免費額度已完全耗盡（或遭短期封鎖）。請點擊側邊欄檢測按鈕確認，並換一組新帳號的 API Key 再試！")
                st.stop()

        progress_bar.progress(
            idx / total_tasks,
            text=f"⚡ 辨識進度：已處理 {idx}/{total_tasks} 頁 ({idx * 100 // total_tasks}%)"
        )
        
        # 安全流速限制：間隔約 4 秒，確保穩定控制在 15 RPM 內
        time.sleep(4.0)

    # ----------------------------------------------------
    # 5. 回填 Excel
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

    status_box.update(label="🎉 辨識與 Excel 寫入完成！", state="complete")
    progress_bar.empty()

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    elapsed = time.time() - start_time

    st.balloons()
    st.success(f"✨ 處理完成！耗時 {elapsed:.1f} 秒，成功處理 {success_count}/{total_tasks} 頁單據，填寫了 {total_written} 個儲存格！")

    # ----------------------------------------------------
    # 6. 會計平衡檢核儀表板與下載
    # ----------------------------------------------------
    st.subheader("⚖️ 單據會計平衡勾稽核對表")
    if audit_records:
        df_audit = pd.DataFrame(audit_records)
        unbalanced_count = sum(1 for r in audit_records if "❌" in r["平衡狀態"])
        if unbalanced_count > 0:
            st.warning(f"共有 **{unbalanced_count}** 筆單據「自輸 - 總計」不等於 0，請依下方表格核對單據金額！")
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
    
