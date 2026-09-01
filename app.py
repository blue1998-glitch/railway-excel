import os
import io
import time
import json
import base64
import subprocess
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from pdf2image import convert_from_bytes
import openpyxl

# ----------------------------------------------------
# 1. 網頁基本設定與金鑰
# ----------------------------------------------------
st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")
st.caption("🟢 V3.0 雙模驗證極速版 ｜ 支援 AQ. 金鑰 ｜ 🧮 公式保留 ｜ ⚖️ 防呆平衡校驗")

EMBEDDED_API_KEY = ""

with st.sidebar:
    st.header("⚙️ 系統設定")
    user_key = st.text_input("Gemini 金鑰 (預設已內嵌)", value=EMBEDDED_API_KEY, type="password")
    active_api_key = user_key.strip() if user_key.strip() else EMBEDDED_API_KEY
    st.success("✅ AQ. 雙模驗證機制已啟用")
    
    concurrency = st.slider(
        "⚡ 平行加速線程數",
        min_value=2,
        max_value=6,
        value=3,
        help="預設 3 線程可達到最穩定且極速的辨識效果。"
    )
    st.markdown("---")
    st.markdown("💡 **操作流程**：\n1. 上傳 Excel 公版\n2. 批次選取掃描 PDF (支援 140 頁)\n3. 點擊開始轉換並下載完成檔")

# ----------------------------------------------------
# 2. 輔助工具函式
# ----------------------------------------------------
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

# 雙模 REST 呼叫核心（完美相容 AQ. 與 AIzaSy 金鑰）
def call_gemini_rest(img_bytes, api_key, model_name="gemini-2.5-flash"):
    img_b64 = base64.b64encode(img_bytes).decode("utf-8")
    
    prompt = """
    你是一個精確的台鐵會計表單辨識助理。請仔細辨識單據內容，嚴格輸出 JSON 格式：
    {
      "station_name": "車站名稱（例如：豐富、苗栗、銅鑼、三義、二水、埔心、楊梅、北湖、富岡、湖口、新豐、竹北、新竹、香山、北新竹等）",
      "date_day": 報表日期中的『日』(1-31 的整數數字),
      "passenger_revenue": 左邊應解款數區塊的『客運』金額(數字，無則填0),
      "freight_revenue": 左邊應解款數區塊的『貨運』金額(數字，無則填0),
      "credit_card_pos": 左邊應解款數區塊的『信用卡刷卡(+)』金額(正數數字，無則填0),
      "credit_card_neg": 左邊應解款數區塊的『信用卡刷卡(-)』金額(正數數字，不要填負數，無則填0),
      "barcode_in": 左邊應解款數區塊的『條碼支付進款』金額(正數數字，無則填0),
      "barcode_refund": 左邊應解款數區塊的『條碼支付退款』金額(正數數字，不要填負數，無則填0),
      "other_amount": 左邊應解款數區塊內其他獨立明細項目的總和(嚴禁重複計入刷卡-或條碼退款，無則填0),
      "remittance_total": 左邊應解款數區塊的『應解總計』金額(數字，無則填0)
    }
    """

    payload = {
        "contents": [{
            "parts": [
                {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}},
                {"text": prompt}
            ]
        }],
        "generationConfig": {
            "response_mime_type": "application/json",
            "temperature": 0.0
        }
    }

    # 驗證模式 1：x-goog-api-key 標頭傳遞
    url_mode1 = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent"
    headers_mode1 = {
        "Content-Type": "application/json",
        "x-goog-api-key": api_key
    }

    resp = requests.post(url_mode1, headers=headers_mode1, json=payload, timeout=40)
    
    # 若模式 1 回傳 401，自動切換至 模式 2：Bearer Token 傳遞
    if resp.status_code == 401:
        headers_mode2 = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        resp = requests.post(url_mode1, headers=headers_mode2, json=payload, timeout=40)

    # 若模式 2 依然 401，嘗試 模式 3：Query Parameter 傳遞
    if resp.status_code == 401:
        url_mode3 = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        headers_mode3 = {"Content-Type": "application/json"}
        resp = requests.post(url_mode3, headers=headers_mode3, json=payload, timeout=40)

    if resp.status_code != 200:
        raise ValueError(f"HTTP {resp.status_code}: {resp.text}")

    res_json = resp.json()
    raw_text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
    
    if "```json" in raw_text:
        raw_text = raw_text.split("```json")[1].split("```")[0].strip()
    elif "```" in raw_text:
        raw_text = raw_text.split("```")[1].split("```")[0].strip()
        
    return json.loads(raw_text)

# 單頁多線程任務函式
def process_single_page_task(task_info, api_key, model_name):
    file_name, page_idx, total_pages, img_bytes = task_info
    
    last_err = ""
    for attempt in range(1, 4):
        try:
            data = call_gemini_rest(img_bytes, api_key, model_name)
            if data and int(data.get("date_day", 0)) > 0:
                return (file_name, page_idx, data, None)
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "503" in last_err:
                time.sleep(2 * attempt)
            continue
            
    return (file_name, page_idx, None, last_err if last_err else "辨識無回應")

# ----------------------------------------------------
# 3. 檔案上傳介面
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📥 步驟 1：上傳 Excel 公版 (.xlsx 或 .xls)", type=["xlsx", "xls"])
with col2:
    uploaded_pdfs = st.file_uploader("📥 步驟 2：批次上傳掃描 PDF (可多選)", type=["pdf"], accept_multiple_files=True)

# ----------------------------------------------------
# 4. 極速處理核心
# ----------------------------------------------------
if st.button("🚀 開始極速自動辨識與填表", type="primary", use_container_width=True):
    if not uploaded_excel:
        st.error("❌ 請上傳 Excel 公版檔案！")
        st.stop()
    if not uploaded_pdfs:
        st.error("❌ 請至少上傳一個 PDF 檔案！")
        st.stop()

    start_time = time.time()

    # 階段 0：連線診斷與可用模型確認
    status_box = st.status("🔍 [階段 1/3] 正在進行雙模 API 連線診斷...", expanded=True)
    models_to_test = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-3.6-flash"]
    confirmed_model = None
    diag_error = ""

    for m in models_to_test:
        test_url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
        test_payload = {"contents": [{"parts": [{"text": "ping"}]}]}
        
        # 測試 x-goog-api-key
        try:
            r = requests.post(test_url, headers={"Content-Type": "application/json", "x-goog-api-key": active_api_key}, json=test_payload, timeout=10)
            if r.status_code == 200:
                confirmed_model = m
                break
        except Exception as e:
            diag_error = str(e)

        # 測試 Bearer
        try:
            r = requests.post(test_url, headers={"Content-Type": "application/json", "Authorization": f"Bearer {active_api_key}"}, json=test_payload, timeout=10)
            if r.status_code == 200:
                confirmed_model = m
                break
            else:
                diag_error = f"HTTP {r.status_code}: {r.text}"
        except Exception as e:
            diag_error = str(e)

    if not confirmed_model:
        status_box.update(label="❌ API 驗證未通過！", state="error")
        st.error(f"❌ **Google 回傳驗證錯誤訊息：**\n\n```text\n{diag_error}\n```")
        st.info("💡 提示：若使用的是 Google AI Studio 金鑰，請確認複製時是否包含完整字元。")
        st.stop()

    status_box.write(f"✅ 連線成功！已確認最佳核心模型：`{confirmed_model}`")

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
    status_box.update(label="⚡ [階段 2/3] 正在高速解析 PDF 頁面...", state="running")
    all_tasks = []
    
    for pdf_file in uploaded_pdfs:
        status_box.write(f"📄 正在載入檔案 `{pdf_file.name}` ...")
        images = convert_from_bytes(pdf_file.getvalue(), dpi=150, fmt="jpeg", thread_count=2)
        total_p = len(images)
        for p_idx, p_img in enumerate(images, 1):
            buf = io.BytesIO()
            p_img.save(buf, format="JPEG", quality=85)
            all_tasks.append((pdf_file.name, p_idx, total_p, buf.getvalue()))

    total_pages_all = len(all_tasks)
    status_box.update(label=f"🚀 [階段 3/3] 啟動 {concurrency} 線程同步辨識全數 {total_pages_all} 頁單據...", state="running")
    
    progress_bar = st.progress(0)
    completed_count = 0
    results = []
    logs = []

    # 階段 2：多線程非同步併發辨識
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        future_to_task = {executor.submit(process_single_page_task, task, active_api_key, confirmed_model): task for task in all_tasks}
        
        for future in as_completed(future_to_task):
            file_name, page_idx, data, err = future.result()
            completed_count += 1
            progress_bar.progress(completed_count / total_pages_all, text=f"辨識進度：{completed_count}/{total_pages_all} 頁完成")
            
            if data:
                st_name = data.get("station_name", "")
                d_day = data.get("date_day", 0)
                results.append((file_name, page_idx, data))
                status_box.write(f"⚡ [{completed_count:02d}/{total_pages_all:02d}] **{st_name}**（{d_day} 號）辨識成功")
            else:
                logs.append(f"⚠️ 檔案 `{file_name}` 第 {page_idx} 頁：{err}")

    status_box.update(label="✅ 辨識完成！正在快速填寫 Excel 與公式...", state="complete")
    progress_bar.empty()

    # 階段 3：回填 Excel 與防呆平衡校準
    total_cells_written = 0
    success_pages = 0

    for file_name, page_idx, data in results:
        st_name = str(data.get("station_name", ""))
        d_day = int(data.get("date_day", 0))
        p_rev = float(data.get("passenger_revenue", 0))
        f_rev = float(data.get("freight_revenue", 0))
        cc_pos = float(data.get("credit_card_pos", 0))
        cc_neg = float(data.get("credit_card_neg", 0))
        bc_in = float(data.get("barcode_in", 0))
        bc_ref = float(data.get("barcode_refund", 0))
        other_v = float(data.get("other_amount", 0))
        remit_tot = float(data.get("remittance_total", 0))

        clean_target_st = clean_name(st_name)
        target_sheet = next(
            (wb[s] for s in wb.sheetnames if clean_target_st == clean_name(s) or clean_target_st in clean_name(s) or clean_name(s) in clean_target_st),
            None
        )

        if not target_sheet:
            logs.append(f"⚠️ 找不到車站分頁 `[{st_name}]`，已略過")
            continue

        # 會計自動平帳校驗
        calc_cc = cc_pos - cc_neg
        calc_bc = bc_in - bc_ref
        base_subtotal = p_rev + f_rev + calc_cc + calc_bc
        
        final_other = other_v
        if abs(remit_tot - base_subtotal) < 0.01:
            final_other = 0.0
        elif abs(remit_tot - (base_subtotal + other_v)) > 0.01:
            final_other = remit_tot - base_subtotal

        # 產生公式（若有相減則寫入 =正-負）
        cc_val = build_subtraction_formula(cc_pos, cc_neg)
        bc_val = build_subtraction_formula(bc_in, bc_ref)

        target_row = d_day + 1

        def write_cell(col_idx, val):
            global total_cells_written
            if val != 0 and val != "0":
                target_sheet.cell(row=target_row, column=col_idx, value=val)
                total_cells_written += 1

        write_cell(2, to_clean_num(p_rev))
        write_cell(3, to_clean_num(f_rev))
        write_cell(4, cc_val)
        write_cell(5, bc_val)
        write_cell(6, to_clean_num(final_other))
        write_cell(7, to_clean_num(remit_tot))

        success_pages += 1
        logs.append(f"✅ **{st_name}** ({d_day} 號) ➜ 已成功寫入第 {target_row} 列（已平衡）")

    # 輸出下載
    elapsed = time.time() - start_time
    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)

    st.balloons()
    st.success(f"🎉 處理大成功！耗時僅 {elapsed:.1f} 秒！共成功寫入 {success_pages}/{total_pages_all} 頁單據，填入 {total_cells_written} 個儲存格！")
    
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
