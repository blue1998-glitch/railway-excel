import io
import re
import time
import random
import openpyxl
import pandas as pd
import streamlit as st
import fitz
from pydantic import BaseModel, Field
from google import genai
from google.genai import types, errors

st.set_page_config(page_title="台鐵解款單自動化填表系統", page_icon="🚆", layout="wide")
st.title("🚆 台鐵掃描解款單 ➜ Excel 智慧自動填表系統")

raw_key = st.secrets.get("GEMINI_API_KEY", "")
with st.sidebar:
    user_key = st.text_input("Gemini API Key", value=str(raw_key).replace('"', "").replace("'", "").strip(), type="password")
    active_api_key = user_key.strip()
    speed = st.select_slider("辨識速度 / 穩定度", options=["最穩定", "快速"], value="最穩定")
    min_interval = 4.2 if speed == "最穩定" else 2.5
    st.caption("建議使用「最穩定」：逐頁送出，避免並行造成 429。")

class StationReport(BaseModel):
    station_name: str = Field(description="車站名稱，只填站名，不要站碼、不要「站」字")
    date_day: int = Field(description="進款日期的日/號，1到31")
    passenger_revenue: float = Field(default=0, description="應解款數：客運(+)")
    freight_revenue: float = Field(default=0, description="應解款數：貨運(+)")
    credit_card_charge: float = Field(default=0, description="應解款數：信用卡刷卡(-)，填正數")
    credit_card_refund: float = Field(default=0, description="應解款數：信用卡退刷(+)，填正數")
    barcode_in: float = Field(default=0, description="應解款數：條碼支付進款(-)，填正數")
    barcode_refund: float = Field(default=0, description="應解款數：條碼支付退款(+)，填正數")
    other_amount: float = Field(default=0, description="應解款數：其他明細淨額")
    remittance_total: float = Field(default=0, description="應解款數：應解總計")

STATIONS = "豐富、苗栗、銅鑼、三義、泰安、后里、豐原、栗林、潭子、頭家厝、松竹、太原、精武、台中、五權、大慶、新烏日、烏日、成功、彰化、花壇、大村、員林、社頭、田中、二水"

def clean_station_name(v):
    s = str(v or "").replace("臺", "台").replace("站", "").strip()
    return re.sub(r"^\d+[_ ]*", "", s).replace(" ", "").replace("　", "")

def num(v):
    try:
        x = float(v)
        return int(x) if x.is_integer() else x
    except Exception:
        return 0

def formula(a, b):
    a, b = num(a), num(b)
    if a == 0 and b == 0:
        return 0
    return f"={a}-{b}" if a and b else (a if a else f"=-{b}")

def find_header(sheet, words):
    for r in range(1, 8):
        for c in range(1, sheet.max_column + 1):
            v = str(sheet.cell(r, c).value or "").replace(" ", "").replace("\n", "")
            if all(w in v for w in words):
                return c
    return None

def sheet_columns(sheet):
    # 優先找「明確標題」，找不到才用公版預設欄位。
    return {
        "date": find_header(sheet, ["日"]) or find_header(sheet, ["日期"]) or 1,
        "passenger": find_header(sheet, ["客運"]) or 2,
        "freight": find_header(sheet, ["貨運"]) or 3,
        "credit": find_header(sheet, ["信用卡"]) or 4,
        "barcode": find_header(sheet, ["條碼"]) or 5,
        "other": find_header(sheet, ["其他"]) or 6,
        "remittance": find_header(sheet, ["自輸"]) or find_header(sheet, ["應解總計"]) or 7,
    }

def find_row(sheet, day, col):
    for r in range(1, min(sheet.max_row, 60) + 1):
        v = sheet.cell(r, col).value
        try:
            if int(float(str(v).replace("日", "").replace("號", "").strip())) == int(day):
                return r
        except Exception:
            pass
    return None

def write(sheet, row, col, value):
    if col:
        sheet.cell(row, col).value = value
        return 1
    return 0

class FatalAPIError(Exception):
    pass

PROMPT = f"""
你是台鐵解款單 OCR 專家。請只辨識這「一張圖片」，不可猜測，不可把其他區域數字當成應解款數。
車站只可能是：{STATIONS}。
請放大並逐格檢查左側「應解款數」區塊，尤其注意淺色複寫、0/3/8/1/7。
必須辨識：
1. 車站名稱
2. 日期的「日/號」
3. 客運(+)
4. 貨運(+)
5. 信用卡刷卡(-)
6. 信用卡退刷(+)
7. 條碼支付進款(-)
8. 條碼支付退款(+)
9. 其他明細的淨額
10. 應解總計

沒有金額的項目填 0。所有金額填數字，不要千分位逗號。
交叉驗算：客運 + 貨運 - 信用卡刷卡 + 信用卡退刷 - 條碼進款 + 條碼退款 + 其他 = 應解總計。
若影像中數字不清楚，先仔細重新查看原圖，不要用公式硬猜；只有看清楚後才能填值。
"""

def call_page(client, model, image_bytes, max_retries=5):
    last = ""
    for retry in range(max_retries):
        try:
            res = client.models.generate_content(
                model=model,
                contents=[types.Part.from_bytes(data=image_bytes, mime_type="image/png"), PROMPT],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=StationReport,
                    temperature=0,
                ),
            )
            if res and res.text:
                return StationReport.model_validate_json(res.text), None
            last = "模型沒有回傳內容"
        except errors.ClientError as e:
            code = getattr(e, "code", None)
            last = f"{code} {getattr(e, 'message', str(e))}"
            if code in (401, 403, 404):
                raise FatalAPIError(last)
            if code != 429:
                break
            # 429 時先等更久；每次重試時間增加，避免連續撞限制。
            time.sleep(min(60, 8 * (retry + 1) + random.uniform(0.5, 2)))
            continue
        except errors.ServerError as e:
            last = str(e)
            time.sleep(min(30, 4 * (retry + 1)))
            continue
        except Exception as e:
            last = str(e)
            time.sleep(2)
    return None, f"辨識失敗：{last}"

def render_page(page, dpi=180):
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    return pix.tobytes("png")

# 上傳
c1, c2 = st.columns(2)
with c1:
    uploaded_excel = st.file_uploader("📥 1. 上傳公版 Excel", type=["xlsx"])
with c2:
    uploaded_pdfs = st.file_uploader("📥 2. 上傳掃描 PDF（可多選）", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 開始辨識並填入 Excel", type="primary", use_container_width=True):
    if not active_api_key or not uploaded_excel or not uploaded_pdfs:
        st.error("❌ 請確認 Gemini API Key、Excel、PDF 都已提供。")
        st.stop()

    start = time.time()
    client = genai.Client(api_key=active_api_key)

    try:
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_excel.getvalue()), data_only=False)
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗：{e}")
        st.stop()

    # 動態取得可用模型；優先 Flash，避免寫死不存在的模型名稱。
    models = []
    try:
        for m in client.models.list():
            name = getattr(m, "name", "").replace("models/", "")
            actions = getattr(m, "supported_actions", None)
            if "gemini" in name.lower() and (not actions or "generateContent" in actions) and "embedding" not in name.lower():
                models.append(name)
    except Exception:
        pass
    if not models:
        models = ["gemini-2.5-flash"]
    models.sort(key=lambda x: (0 if "flash" in x.lower() and "lite" not in x.lower() else 1 if "flash" in x.lower() else 2, x))
    model = st.sidebar.selectbox("AI 模型", models)

    pages = []
    status = st.status("📄 正在準備 PDF 頁面...", expanded=True)
    for pdf in uploaded_pdfs:
        try:
            doc = fitz.open(stream=pdf.getvalue(), filetype="pdf")
            for i, page in enumerate(doc):
                pages.append((pdf.name, i + 1, len(doc), render_page(page)))
            doc.close()
        except Exception as e:
            status.write(f"⚠️ {pdf.name}：{e}")

    if not pages:
        st.error("❌ 找不到 PDF 頁面。")
        st.stop()

    status.update(label=f"🔎 逐頁辨識中：共 {len(pages)} 頁", state="running")
    progress = st.progress(0)
    results = []
    failures = []

    # 刻意不用多執行緒：Gemini 的 RPM/TPM 限制下，多線程會同時撞 API，造成整批 429。
    for n, (file_name, page_no, total, image) in enumerate(pages, 1):
        if n > 1:
            time.sleep(min_interval)
        data, err = call_page(client, model, image)
        if data and 1 <= data.date_day <= 31 and clean_station_name(data.station_name):
            results.append((n, file_name, page_no, total, data))
            status.write(f"✅ {file_name} 第 {page_no}/{total} 頁：{data.station_name}、{data.date_day} 日")
        else:
            failures.append((file_name, page_no, err or "資料不完整"))
            status.write(f"❌ {file_name} 第 {page_no}/{total} 頁：{err or '資料不完整'}")
        progress.progress(n / len(pages), text=f"已完成 {n}/{len(pages)} 頁")

    status.update(label="📝 正在精準寫入 Excel...", state="running")
    audit = []
    written = 0

    for _, file_name, page_no, _, d in results:
        target = None
        clean = clean_station_name(d.station_name)
        for name in wb.sheetnames:
            if clean_station_name(name) == clean:
                target = wb[name]
                break
        if target is None:
            failures.append((file_name, page_no, f"找不到車站分頁：{d.station_name}"))
            continue

        cols = sheet_columns(target)
        row = find_row(target, d.date_day, cols["date"])
        if row is None:
            failures.append((file_name, page_no, f"找不到日期列：{d.date_day}"))
            continue

        cc = formula(d.credit_card_charge, d.credit_card_refund)
        bc = formula(d.barcode_in, d.barcode_refund)
        calc = d.passenger_revenue + d.freight_revenue - d.credit_card_charge + d.credit_card_refund - d.barcode_in + d.barcode_refund + d.other_amount
        diff = round(d.remittance_total - calc, 2)

        # 0 也寫入：避免「辨識成功但儲存格仍空白」。
        vals = {
            "passenger": num(d.passenger_revenue),
            "freight": num(d.freight_revenue),
            "credit": cc,
            "barcode": bc,
            "other": num(d.other_amount),
            "remittance": num(d.remittance_total),
        }
        for key, value in vals.items():
            written += write(target, row, cols[key], value)

        audit.append({
            "檔案": file_name, "頁": page_no, "車站": d.station_name, "日期": d.date_day,
            "客運": num(d.passenger_revenue), "貨運": num(d.freight_revenue),
            "信用卡": cc, "條碼": bc, "其他": num(d.other_amount),
            "PDF應解總計": num(d.remittance_total), "計算總計": num(calc), "差額": diff,
            "狀態": "✅ 平衡" if abs(diff) < 0.01 else "⚠️ 請核對"
        })

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    status.update(label="🎉 完成", state="complete")
    progress.empty()

    ok = len(audit)
    st.success(f"完成：成功填入 {ok}/{len(pages)} 頁、{written} 個儲存格；耗時 {time.time() - start:.1f} 秒。")

    if audit:
        st.subheader("⚖️ 辨識與平衡檢核")
        st.dataframe(pd.DataFrame(audit), use_container_width=True, hide_index=True)
    if failures:
        st.warning(f"⚠️ 有 {len(failures)} 頁需要人工確認：")
        st.dataframe(pd.DataFrame(failures, columns=["檔案", "頁", "原因"]), use_container_width=True, hide_index=True)

    st.download_button(
        "📥 下載完成的 Excel",
        data=out,
        file_name="台鐵解款單_彙總完成表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
