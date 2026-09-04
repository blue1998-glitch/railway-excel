# 1. 安裝必要套件（第一次執行請取消註解）
# !pip install -q openpyxl google-generativeai

import json
import os
import re
import subprocess
import google.generativeai as genai
import openpyxl

# ==================== 參數設定區 ====================
API_KEY = "YOUR_GEMINI_API_KEY"  # 請填入您的 Gemini API 金鑰
TEMPLATE_FILE = "豐富至二水解款單(公版)2.xlsx"  # 公版檔案路徑（若為 .xls 程式會自動轉檔）
OUTPUT_FILE = "台鐵解款單_彙總完成表_最新.xlsx"  # 完成匯總的輸出檔名

# 待辨識的 PDF 檔案清單（若在 Colab 可直接填寫檔名）
PDF_FILES = [
    "20260903084029_2.pdf",  # 8/28
    "20260903084105_2.pdf",  # 8/29
    "20260903084158_2.pdf",  # 8/30
]
# ====================================================

genai.configure(api_key=API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

# 提示詞：只抓取原始數字，禁止 AI 心算，杜絕算錯
PROMPT = """
你是台鐵站務會計專業助理。這是一份掃描的「站務解款單」PDF 文件，每一頁代表一個車站。
請依序擷取每一頁的資料，輸出為純 JSON 陣列格式。
【重要】：請直接提取表上的原始數字，千萬不要自行加減！
每個車站物件欄位格式如下（若無該項目或為 0 請填 0）：
[
  {
    "站名": "豐富",
    "進款日期": "2026-08-28",
    "客運": 26491,
    "貨運": 0,
    "信用卡刷卡": 2233,
    "信用卡退刷": 0,
    "條碼支付進款": 287,
    "條碼支付退款": 0,
    "現金小計": 23971,
    "面額_2000": 0,
    "面額_1000": 9,
    "面額_500": 7,
    "面額_200": 0,
    "面額_100": 91,
    "面額_50": 29,
    "面額_20": 0,
    "面額_10": 92,
    "面額_5": 0,
    "面額_1": 1
  }
]
請只回傳 JSON 陣列（以 [ 開頭，以 ] 結尾），不要有任何額外說明。
"""


def convert_xls_to_xlsx(filename):
  """若為舊版 .xls，自動轉換為 .xlsx 以保留全部樣式與公式"""
  if filename.endswith(".xls") and not filename.endswith(".xlsx"):
    new_name = filename + "x"
    if not os.path.exists(new_name):
      subprocess.run(
          ["libreoffice", "--headless", "--convert-to", "xlsx", filename],
          check=False,
      )
    return new_name if os.path.exists(new_name) else filename
  return filename


def parse_pdf_data(pdf_path):
  """透過 Gemini API 辨識整份 PDF"""
  print(f"正在透過 AI 辨識：{pdf_path} ...")
  uploaded_file = genai.upload_file(pdf_path, mime_type="application/pdf")
  response = model.generate_content([uploaded_file, PROMPT])

  # 清理 Markdown 標記並轉為 Python 字典
  text = response.text.strip()
  text = re.sub(r"^```json\s*", "", text)
  text = re.sub(r"\s*```$", "", text)
  return json.loads(text)


def find_mappings(ws):
  """動態找出工作表中『車站列號』與『欄位欄號』，避免硬編碼跑位"""
  stations = [
      "新烏日",
      "烏日",
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
      "台中",
      "五權",
      "大慶",
      "成功",
      "彰化",
      "花壇",
      "大村",
      "員林",
      "社頭",
      "田中",
      "二水",
  ]
  st_rows = {}
  for r in range(1, ws.max_row + 1):
    for c in range(1, min(10, ws.max_column + 1)):
      val = str(ws.cell(row=r, column=c).value or "").strip()
      for st in stations:
        std_name = "臺中" if st in ["臺中", "台中"] else st
        if st in val and len(val) >= 2 and std_name not in st_rows:
          st_rows[std_name] = r

  col_map = {}
  for r in range(1, min(7, ws.max_row + 1)):
    for c in range(1, ws.max_column + 1):
      val = (
          str(ws.cell(row=r, column=c).value or "").replace("\n", "").strip()
      )
      if "客運" in val and "客運" not in col_map:
        col_map["客運"] = c
      elif "貨運" in val and "貨運" not in col_map:
        col_map["貨運"] = c
      elif (
          "信用卡" in val or "電腦信用卡" in val
      ) and "電腦信用卡" not in col_map:
        col_map["電腦信用卡"] = c
      elif ("條碼" in val or "行動支付" in val) and "條碼" not in col_map:
        col_map["條碼"] = c
      elif (
          "現金" in val or "實解" in val or "小計" in val
      ) and "現金小計" not in col_map:
        col_map["現金小計"] = c
      # 面額欄位比對
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
        if (
            denom in val
            and f"面額_{denom}" not in col_map
            and not any(x in val for x in ["總計", "應解"])
        ):
          col_map[f"面額_{denom}"] = c

  return st_rows, col_map


def main():
  real_template = convert_xls_to_xlsx(TEMPLATE_FILE)
  if not os.path.exists(real_template):
    print(f"找不到公版檔案：{real_template}，請確認檔名及路徑！")
    return

  # 載入公版（保留所有樣式與原公式）
  wb = openpyxl.load_workbook(real_template)

  for pdf_file in PDF_FILES:
    if not os.path.exists(pdf_file):
      print(f"跳過不存在的 PDF：{pdf_file}")
      continue

    data = parse_pdf_data(pdf_file)
    if not data:
      continue

    # 取得此 PDF 的日期（例如 2026-08-28 -> day = 28）
    date_str = data[0].get("進款日期", "")
    day_match = re.search(r"(\d{4})[-年/](\d{1,2})[-月/](\d{1,2})", date_str)
    day = int(day_match.group(3)) if day_match else None

    # 尋找或複製對應日期的工作表
    target_ws = None
    if day:
      for name in wb.sheetnames:
        if f"{day:02d}" in name or f"{day}日" in name or name == str(day):
          target_ws = wb[name]
          break

    if target_ws is None:
      # 若公版只有一張表，則以公版為範本複製出當天的工作表
      base_ws = wb.sheetnames[0]
      target_ws = wb.copy_worksheet(wb[base_ws])
      target_ws.title = (
          f"{day_match.group(2)}月{day_match.group(3)}日"
          if day_match
          else f"彙總_{os.path.basename(pdf_file)}"
      )

    st_rows, col_map = find_mappings(target_ws)

    # 逐站寫入資料
    for row_data in data:
      st_name = row_data.get("站名", "").replace("站", "").strip()
      if st_name == "台中":
        st_name = "臺中"

      row_idx = st_rows.get(st_name)
      if not row_idx:
        continue

      # 1. 客運、貨運
      if "客運" in col_map:
        target_ws.cell(
            row=row_idx, column=col_map["客運"]
        ).value = row_data.get("客運", 0)
      if "貨運" in col_map:
        target_ws.cell(
            row=row_idx, column=col_map["貨運"]
        ).value = row_data.get("貨運", 0)

      # 2. 電腦信用卡公式：信用卡刷卡(-) 減 信用卡退刷(+)
      c_in = int(row_data.get("信用卡刷卡", 0) or 0)
      c_out = int(row_data.get("信用卡退刷", 0) or 0)
      if "電腦信用卡" in col_map:
        cell = target_ws.cell(row=row_idx, column=col_map["電腦信用卡"])
        if c_in > 0 or c_out > 0:
          cell.value = f"={c_in}-{c_out}"  # 保留如 =62526-7915 的公式軌跡
        else:
          cell.value = 0

      # 3. 條碼支付公式：條碼支付進款(-) 減 條碼支付退款(+)
      b_in = int(row_data.get("條碼支付進款", 0) or 0)
      b_out = int(row_data.get("條碼支付退款", 0) or 0)
      if "條碼" in col_map:
        cell = target_ws.cell(row=row_idx, column=col_map["條碼"])
        if b_in > 0 or b_out > 0:
          cell.value = f"={b_in}-{b_out}"  # 保留如 =2486-1676 的公式軌跡
        else:
          cell.value = 0

      # 4. 現金小計 / 實解總計
      if "現金小計" in col_map:
        target_ws.cell(
            row=row_idx, column=col_map["現金小計"]
        ).value = row_data.get("現金小計", 0)

      # 5. 各面額張數（若公版有相應欄位則寫入）
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
        denom_key = f"面額_{denom}"
        if denom_key in col_map:
          target_ws.cell(
              row=row_idx, column=col_map[denom_key]
          ).value = row_data.get(denom_key, 0)

    print(f"已成功填入工作表：{target_ws.title}")

  wb.save(OUTPUT_FILE)
  print(f"\n全部處理完成！已成功儲存至：{OUTPUT_FILE}")


if __name__ == "__main__":
  main()
