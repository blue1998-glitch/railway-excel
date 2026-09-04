# ============================================================
# 台鐵站務解款單 PDF 辨識與公版 Excel 填報程式
# ============================================================
# 若在 Google Colab 執行，請先執行：
# !pip install -q -U google-genai openpyxl

import json
import os
import re
import openpyxl

# ------------------------------------------------------------
# 參數設定區（請依實際檔名與金鑰修改）
# ------------------------------------------------------------
API_KEY = "YOUR_GEMINI_API_KEY"  # 請填入 Google AI Studio API Key
PDF_PATH = "20260903084029_2.pdf"  # 掃描之 PDF 檔案路徑
TEMPLATE_PATH = "豐富至二水解款單(公版)2.xlsx"  # 公版 Excel（請確認為 .xlsx 格式）
OUTPUT_PATH = "台鐵解款單_彙總完成表_產出.xlsx"  # 填報完成產出路徑

STATIONS = [
    "豐富",
    "苗栗",
    "銅鑼",
    "三義",
    "泰安",
    "后里",
    "豐原",
    "栗林",
    "潭子",
    "頭家厝",
    "松竹",
    "太原",
    "精武",
    "臺中",
    "五權",
    "大慶",
    "烏日",
    "新烏日",
    "成功",
    "彰化",
    "花壇",
    "大村",
    "員林",
    "社頭",
    "田中",
    "二水",
]


def extract_data_from_pdf(pdf_path, api_key):
  """呼叫 Gemini 辨識 PDF 站務解款單，提取各站純數值資料"""
  prompt = """
    你是一位台鐵會計審核員，請逐頁辨識此 PDF 中的「站務解款單」，擷取所有車站數據。
    請嚴格輸出純 JSON 陣列（Array of Objects），禁止夾帶任何其他文字或 markdown，格式如下：
    [
      {
        "station_name": "豐富",
        "station_code": "3150",
        "date": "2026-08-28",
        "passenger": 26491,
        "freight": 0,
        "stored_freight": 0,
        "card_charge": 2233,
        "card_refund": 0,
        "barcode_charge": 287,
        "barcode_refund": 0,
        "cheque": 0,
        "supplementary": 0,
        "revolving_fund": 0,
        "expected_total": 23971,
        "actual_total": 23971,
        "cash_subtotal": 23971,
        "voucher": 0,
        "bills_count": {"2000":0,"1000":9,"500":7,"200":0,"100":91,"50":29,"20":0,"10":92,"5":0,"1":1},
        "bills_amount": {"2000":0,"1000":9000,"500":3500,"200":0,"100":9100,"50":1450,"20":0,"10":920,"5":0,"1":1}
      }
    ]
    規則：
    1. 刷卡(-) 與 退刷(+) 請分別填寫正整數原始金額（不帶負號）。
    2. 條碼進款(-) 與 條碼退款(+) 請分別填寫正整數原始金額。
    3. 無資料或為0請填 0，不可為 null。
    """
  raw_text = ""
  try:
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    f = client.files.upload(file=pdf_path)
    res = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[f, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        ),
    )
    raw_text = res.text
  except Exception:
    import google.generativeai as genai

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        "gemini-2.5-flash",
        generation_config={"response_mime_type": "application/json"},
    )
    f = genai.upload_file(pdf_path)
    res = model.generate_content([f, prompt])
    raw_text = res.text

  match = re.search(r"\[.*\]", raw_text, re.DOTALL)
  return json.loads(match.group(0)) if match else json.loads(raw_text)


def locate_sheet(wb, date_str):
  """依據進款日期比對工作表（如 8.28、28日、28），比對不到則取啟用工作表"""
  if date_str:
    parts = date_str.replace("/", "-").split("-")
    if len(parts) >= 3:
      m, d = str(int(parts[1])), str(int(parts[2]))
      for cand in [f"{m}.{d}", f"{int(m):02d}.{int(d):02d}", f"{d}日", d]:
        if cand in wb.sheetnames:
          return wb[cand]
  return wb.active


def locate_structure(ws):
  """自動定位車站列號與各欄位索引"""
  station_rows = {}
  col_map = {}

  # 掃描前 45 列定位車站
  for r in range(1, min(ws.max_row + 1, 45)):
    for c in range(1, min(ws.max_column + 1, 6)):
      val = str(ws.cell(row=r, column=c).value or "").strip()
      for st in STATIONS:
        if st in val or st.replace("臺", "台") in val.replace("臺", "台"):
          if st not in station_rows:
            station_rows[st] = r

  # 掃描前 5 列定位欄位
  keywords = {
      "客運": "客運",
      "貨運": "貨運",
      "存付": "存付",
      "信用卡": "信用卡",
      "條碼": "條碼",
      "支票": "支票",
      "補繳": "補繳",
      "週轉": "週轉金",
      "應解": "應解總計",
      "實解": "實解總計",
      "現金": "現金小計",
      "憑證": "憑證",
  }
  denoms = [
      ("2000", "貳仟"),
      ("1000", "壹仟"),
      ("500", "伍佰"),
      ("200", "貳佰"),
      ("100", "壹佰"),
      ("50", "伍拾"),
      ("20", "貳拾"),
      ("10", "拾元"),
      ("5", "伍元"),
      ("1", "壹元"),
  ]

  for r in range(1, min(ws.max_row + 1, 6)):
    for c in range(1, ws.max_column + 1):
      txt = (
          str(ws.cell(row=r, column=c).value or "").replace(" ", "").replace("\n", "")
      )
      if not txt:
        continue
      for k, name in keywords.items():
        if k in txt and name not in col_map:
          col_map[name] = c
      for d_num, d_chn in denoms:
        if (d_num in txt or d_chn in txt) and d_num not in col_map:
          col_map[d_num] = c
          col_map[f"{d_num}_is_amt"] = "額" in txt

  return station_rows, col_map


def safe_set(ws, r, c, val, force_formula=False):
  """保護儲存格寫入：若非強制公式且儲存格原已有公式（=開頭），則保留原公版公式"""
  if not r or not c:
    return
  cell = ws.cell(row=r, column=c)
  if not force_formula and str(cell.value or "").strip().startswith("="):
    return
  cell.value = val


def fill_excel(data_list, template_path, output_path):
  """開啟公版活頁簿並回填數據與公式"""
  wb = openpyxl.load_workbook(template_path)
  first_date = data_list[0].get("date", "") if data_list else ""
  ws = locate_sheet(wb, first_date)
  station_rows, col_map = locate_structure(ws)

  print(f"正在填入工作表：{ws.title}，已匹配 {len(station_rows)} 個車站列...")

  for item in data_list:
    st_name = item.get("station_name", "")
    r = None
    for k, row_num in station_rows.items():
      if (
          k in st_name
          or st_name in k
          or k.replace("臺", "台") in st_name.replace("臺", "台")
      ):
        r = row_num
        break
    if not r:
      continue

    # 1. 電腦信用卡：刷卡(-) 減 退刷(+)，建立如 =62526-7915 之公式紀錄
    c_in = item.get("card_charge", 0) or 0
    c_out = item.get("card_refund", 0) or 0
    card_val = f"={c_in}-{c_out}" if (c_in or c_out) else 0

    # 2. 條碼：條碼進款(-) 減 條碼退款(+)，建立如 =3486-0 之公式紀錄
    b_in = item.get("barcode_charge", 0) or 0
    b_out = item.get("barcode_refund", 0) or 0
    barcode_val = f"={b_in}-{b_out}" if (b_in or b_out) else 0

    # 寫入各欄位（保留原有公版樣式）
    safe_set(ws, r, col_map.get("客運"), item.get("passenger", 0))
    safe_set(ws, r, col_map.get("貨運"), item.get("freight", 0))
    safe_set(ws, r, col_map.get("存付"), item.get("stored_freight", 0))
    safe_set(ws, r, col_map.get("信用卡"), card_val, force_formula=True)
    safe_set(ws, r, col_map.get("條碼"), barcode_val, force_formula=True)
    safe_set(ws, r, col_map.get("支票"), item.get("cheque", 0))
    safe_set(ws, r, col_map.get("補繳"), item.get("supplementary", 0))
    safe_set(ws, r, col_map.get("週轉金"), item.get("revolving_fund", 0))
    safe_set(ws, r, col_map.get("應解總計"), item.get("expected_total", 0))
    safe_set(ws, r, col_map.get("實解總計"), item.get("actual_total", 0))
    safe_set(ws, r, col_map.get("現金小計"), item.get("cash_subtotal", 0))
    safe_set(ws, r, col_map.get("憑證"), item.get("voucher", 0))

    # 券幣明細填入
    counts = item.get("bills_count", {})
    amounts = item.get("bills_amount", {})
    for denom in [
        "2000",
        "1000",
        "500",
        "200",
        "100",
        "50",
        "20",
        "10",
        "5",
        "1",
    ]:
      c_idx = col_map.get(denom)
      if c_idx:
        val = (
            amounts.get(denom, 0)
            if col_map.get(f"{denom}_is_amt")
            else counts.get(denom, 0)
        )
        safe_set(ws, r, c_idx, val)

  wb.save(output_path)
  print(f"處理完成！已成功輸出檔案至：{output_path}")


# ------------------------------------------------------------
# 主程式執行入口
# ------------------------------------------------------------
if __name__ == "__main__":
  if not os.path.exists(PDF_PATH):
    print(f"錯誤：找不到 PDF 檔案 {PDF_PATH}")
  elif not os.path.exists(TEMPLATE_PATH):
    print(f"錯誤：找不到公版檔案 {TEMPLATE_PATH}")
  else:
    print("正在呼叫 Gemini 進行解款單辨識...")
    extracted_data = extract_data_from_pdf(PDF_PATH, API_KEY)
    print(f"辨識成功，共取得 {len(extracted_data)} 個車站數據。開始填報公版...")
    fill_excel(extracted_data, TEMPLATE_PATH, OUTPUT_PATH)
