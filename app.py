import streamlit as st, yfinance as yf, pandas as pd, numpy as np, json, os, re, requests
from datetime import datetime, timezone, timedelta
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_gsheets import GSheetsConnection
from google import genai

# 基本頁面配置與時區工具
st.set_page_config(layout="wide", initial_sidebar_state="collapsed", page_title="台股動能 RS 與大盤寬度監控")
DATA_FILE, TW_TZ = "portfolio.json", timezone(timedelta(hours=8))
get_tw_now = lambda: datetime.now(TW_TZ)
get_tw_now_str = lambda fmt="%Y-%m-%d %H:%M:%S": get_tw_now().strftime(fmt)

for k, v in [("last_portfolio_refresh", get_tw_now_str()), ("search_input_val", ""), ("chat_history", [])]:
    st.session_state.setdefault(k, v)

def _load_names():
    try:
        if os.path.exists("stock_names.json"):
            with open("stock_names.json", "r", encoding="utf-8") as f: return json.load(f)
    except Exception: pass
    return {}

OFFICIAL_STOCK_NAMES = _load_names()
clean_sym = lambda v: str(v or "").strip()[:-2] if str(v or "").strip().endswith(".0") else str(v or "").strip()

def clean_stock_name(name, symbol=None):
    sym = clean_sym(symbol).upper()
    if sym in OFFICIAL_STOCK_NAMES: return OFFICIAL_STOCK_NAMES[sym]
    if not name: return sym
    raw = str(name).strip()
    for std_n in OFFICIAL_STOCK_NAMES.values():
        if (raw == std_n or raw.startswith(std_n)) and len(raw) <= len(std_n) + 12: return std_n
    for suf in ["股份有限公司台灣分公司", "股份有限公司", "有限股份公司", "有限公司", "(股)公司", "（股）公司"]:
        raw = raw.replace(suf, "")
    return raw.strip() or sym

def _clean_date_series(df):
    if df is None or df.empty: return pd.DataFrame()
    d = df.reset_index() if "Date" not in df.columns else df.copy()
    for col in ["Date", "Datetime", "index", "date"]:
        if col in d.columns:
            try: d["Date"] = pd.to_datetime(pd.to_datetime(d[col], utc=True).dt.tz_convert("Asia/Taipei").dt.strftime("%Y-%m-%d"))
            except Exception: d["Date"] = pd.to_datetime(pd.to_datetime(d[col]).dt.strftime("%Y-%m-%d"))
            if col != "Date": d.drop(columns=[col], inplace=True)
            break
    return d

@st.cache_data(ttl=1800)
def get_benchmark_returns():
    bm_data = {}
    for mkt_key, sym in [("TW", "^TWII"), ("TWO", "^TWOII")]:
        try:
            df = yf.Ticker(sym).history(period="1y")
            if df.empty or len(df) < 20: df = yf.Ticker("0050.TW" if mkt_key == "TW" else "^TWII").history(period="1y")
        except Exception:
            try: df = yf.Ticker("0050.TW").history(period="1y")
            except Exception: df = pd.DataFrame()
        if not df.empty:
            df_c = _clean_date_series(df)
            c = df_c["Close"].values
            calc_r = lambda d: round(((c[-1] - c[-d-1]) / c[-d-1]) * 100, 2) if len(c) > d else 0.0
            bm_data[mkt_key] = {"df": df_c[["Date", "Close"]].rename(columns={"Close": "benchmark_close"}), "r_5d": calc_r(5), "r_20d": calc_r(20), "r_60d": calc_r(60)}
    bm_data.setdefault("TW", {"df": pd.DataFrame(), "r_5d": 0.0, "r_20d": 0.0, "r_60d": 0.0})
    bm_data.setdefault("TWO", bm_data["TW"])
    return bm_data

def calculate_rs_ratio_series(target_df, benchmark_df, rs_w=60, mom_w=20):
    try:
        if target_df is None or target_df.empty or benchmark_df is None or benchmark_df.empty: return pd.DataFrame()
        t_df, b_df = _clean_date_series(target_df), _clean_date_series(benchmark_df)
        if "Close" not in t_df.columns: return pd.DataFrame()
        b_df = b_df.rename(columns={"Close": "benchmark_close"}) if "benchmark_close" not in b_df.columns and "Close" in b_df.columns else b_df
        if "benchmark_close" not in b_df.columns: return pd.DataFrame()

        merged = pd.merge(t_df[["Date", "Close"]].rename(columns={"Close": "target_close"}), b_df[["Date", "benchmark_close"]], on="Date", how="inner").sort_values("Date").reset_index(drop=True)
        if len(merged) < 10: merged = merged.ffill().bfill()
        merged = merged[(merged["benchmark_close"] > 0) & (merged["target_close"] > 0)].copy()
        if merged.empty: return pd.DataFrame()

        merged["rs_raw"] = (merged["target_close"] / merged["benchmark_close"]) * 100.0
        merged["rs_ma60"] = merged["rs_raw"].rolling(rs_w, min_periods=min(len(merged), max(5, rs_w // 4))).mean().bfill()
        merged["rs_ratio"] = np.where(merged["rs_ma60"] > 0, 100.0 * (merged["rs_raw"] / merged["rs_ma60"]), 100.0)
        rs_ratio_ma20 = merged["rs_ratio"].rolling(mom_w, min_periods=min(len(merged), max(3, mom_w // 4))).mean().bfill()
        merged["rs_momentum"] = np.where(rs_ratio_ma20 > 0, 100.0 * (merged["rs_ratio"] / rs_ratio_ma20), 100.0)
        return merged
    except Exception: return pd.DataFrame()

def get_trend_master_status(row):
    rs, badge = float(row.get("rs_rating", 50) or 50), str(row.get("pattern_badge", "") or "")
    r_5d, rs_ratio = float(row.get("r_5d", 0.0) or 0.0), float(row.get("rs_ratio", 100.0) or 100.0)
    p = "🔥[強勢] " if rs_ratio >= 100.0 else "❄️[弱勢] "
    if rs >= 95: sub = "👑 頂級領袖・突破新高 (主力首選)" if "新高" in badge or r_5d >= 10.0 else ("🎯 頂級VCP・即將噴出 (極限強勢)" if "VCP" in badge else "🚀 極致飆股・主升奔馳 (最強5%)")
    elif rs >= 90: sub = "🎯 VCP蓄勢・突破在即 (黃金買點)" if "VCP" in badge else ("⭐ 領袖新高・順風追擊 (多頭先鋒)" if "新高" in badge else "🚀 狂暴主升・沿線抱牢 (第一梯隊)")
    elif rs >= 80: sub = "🎯 VCP收縮・縮量待發 (觀察進場)" if "VCP" in badge else ("⭐ 區間突破・趨勢確立 (順勢加碼)" if "新高" in badge else ("⚠️ 短線強彈・觀察季線 (謹慎試單)" if "反彈" in badge else "⚡ 強大多頭・順勢推升 (右側安全)"))
    elif rs >= 75: sub = "⚠️ 左側反彈・上方有壓 (短打勿追)" if "反彈" in badge else ("🎯 底部收斂・轉強蓄勢 (第二梯隊)" if "VCP" in badge else "🔥 突破初升・動能成型 (第三梯隊)")
    elif rs >= 50: sub = "📦 區間整理・等待表態 (動能平平)"
    else: sub = "⛔ 弱勢落後・左側不碰 (避開死水)"
    return f"{p}{sub}"

def _get_gsheet_conn():
    try: return st.connection("gsheets", type=GSheetsConnection)
    except Exception: return None

def load_data():
    conn = _get_gsheet_conn()
    if conn:
        try:
            df = conn.read(ttl=0)
            if df is not None and not df.empty:
                records = []
                for _, r in df.iterrows():
                    d = r.dropna().to_dict()
                    sym = clean_sym(d.get("symbol", ""))
                    if not sym: continue
                    hist = d.get("history", "[]")
                    records.append({
                        "symbol": sym, "name": clean_stock_name(d.get("name"), sym),
                        "market": str(d.get("market", "TW")).strip().upper(),
                        "entry_date": str(d.get("entry_date", get_tw_now_str("%Y-%m-%d"))).strip(),
                        "avg_cost": float(d.get("avg_cost", 0.0) or 0.0),
                        "shares": int(float(d.get("shares", 0) or 0)),
                        "record_high": float(d.get("record_high", d.get("avg_cost", 0.0)) or 0.0),
                        "realized_pnl": float(d.get("realized_pnl", 0) or 0.0),
                        "status_override": str(d.get("status_override", "")),
                        "history": json.loads(hist) if isinstance(hist, str) else (hist if isinstance(hist, list) else [])
                    })
                if records: return records
        except Exception as e:
            st.sidebar.warning(f"Google Sheets 載入異常（切換至備用檔）: {e}")

    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for d in data:
                    d["symbol"] = clean_sym(d.get("symbol", ""))
                    d["name"] = clean_stock_name(d.get("name"), d.get("symbol"))
                return data
        except Exception: return []
    return []

def save_data(data):
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception: pass
    conn = _get_gsheet_conn()
    if conn:
        try:
            cols = ["symbol", "name", "market", "entry_date", "avg_cost", "shares", "record_high", "realized_pnl", "status_override", "history"]
            if not data:
                conn.update(data=pd.DataFrame(columns=cols))
                return
            rows = [{
                "symbol": clean_sym(it.get("symbol", "")), "name": it.get("name", ""),
                "market": it.get("market", "TW"), "entry_date": str(it.get("entry_date", "")),
                "avg_cost": float(it.get("avg_cost", 0.0)), "shares": int(it.get("shares", 0)),
                "record_high": float(it.get("record_high", 0.0)), "realized_pnl": float(it.get("realized_pnl", 0)),
                "status_override": str(it.get("status_override", "")),
                "history": json.dumps(it.get("history", []), ensure_ascii=False)
            } for it in data]
            conn.update(data=pd.DataFrame(rows))
        except Exception as e: st.error(f"Google Sheets 寫入失敗: {e}")

def make_log_entry(action, price, share_delta, remaining_shares, pnl_text, note):
    return {"時間": get_tw_now_str("%Y-%m-%d %H:%M"), "動作": action, "成交價": price, "異動股數": share_delta, "剩餘股數": remaining_shares, "單筆實現損益": pnl_text, "備註": note}

@st.cache_data(ttl=60)
def load_market_data():
    raw_list, status_msg = [], "無可用資料"
    if os.path.exists("market_rankings.json"):
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime("market_rankings.json"), tz=TW_TZ).strftime("%Y-%m-%d %H:%M:%S")
            with open("market_rankings.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and data: raw_list, status_msg = data, f"本機檔案載入成功 (產出時間: {mtime})"
        except Exception: pass
    if not raw_list:
        try:
            res = requests.get("https://raw.githubusercontent.com/blue1998-glitch/-/main/market_rankings.json", timeout=8)
            if res.status_code == 200:
                data = res.json()
                if isinstance(data, list) and data: raw_list, status_msg = data, f"線上同步成功 (同步時間: {get_tw_now_str()})"
        except Exception as e: return [], f"連線異常: {str(e)}"

    bm_dict = get_benchmark_returns()
    for item in raw_list:
        item["symbol"], item["name"] = clean_sym(item.get("symbol", "")), clean_stock_name(item.get("name"), item.get("symbol"))
        mkt_key = "TWO" if "上櫃" in str(item.get("market", "")) or "TWO" in str(item.get("market", "")).upper() else "TW"
        bm_info = bm_dict.get(mkt_key, bm_dict["TW"])
        bm_r60, bm_r20 = bm_info.get("r_60d", 0.0), bm_info.get("r_20d", 0.0)
        s_r60, s_r20 = float(item.get("r_60d", 0.0) or 0.0), float(item.get("r_20d", 0.0) or 0.0)
        if "rs_ratio" not in item or item["rs_ratio"] in (100.0, None):
            item["rs_ratio"] = round(100.0 * (1.0 + s_r60 / 100.0) / max(0.01, (1.0 + bm_r60 / 100.0)), 2)
        if "rs_momentum" not in item or item["rs_momentum"] in (100.0, None):
            item["rs_momentum"] = round(100.0 * (1.0 + s_r20 / 100.0) / max(0.01, (1.0 + bm_r20 / 100.0)), 2)
    return raw_list, status_msg

def fetch_stock_and_momentum(symbol, market, entry_date_str=None):
    sym_clean = clean_sym(symbol)
    is_otc = "TWO" in str(market).upper() or "上櫃" in str(market)
    ticker, alt_ticker, bm_key = f"{sym_clean}.TWO" if is_otc else f"{sym_clean}.TW", f"{sym_clean}.TW" if is_otc else f"{sym_clean}.TWO", "TWO" if is_otc else "TW"
    try:
        df_all = yf.Ticker(ticker).history(period="1y")
        if df_all.empty:
            df_all = yf.Ticker(alt_ticker).history(period="1y")
            if df_all.empty: return None, None, None, 0.0, 0.0, 0.0, 100.0, 100.0
        cur = round(float(df_all["Close"].iloc[-1]), 2)
        try:
            df_e = df_all.loc[df_all.index.astype(str) >= str(entry_date_str)] if entry_date_str else pd.DataFrame()
            max_h = round(float(df_e["High"].max()), 2) if not df_e.empty else cur
        except Exception: max_h = cur
        ma20 = round(float(df_all["Close"].tail(20).mean()), 2) if len(df_all) >= 20 else cur
        c = df_all["Close"]
        r5 = round(((c.iloc[-1] - c.iloc[-6]) / c.iloc[-6]) * 100, 2) if len(c) >= 6 else 0.0
        r1m = round(((c.iloc[-1] - c.iloc[-21]) / c.iloc[-21]) * 100, 2) if len(c) >= 21 else r5
        r1q = round(((c.iloc[-1] - c.iloc[-61]) / c.iloc[-61]) * 100, 2) if len(c) >= 61 else r1m
        bm_info = get_benchmark_returns().get(bm_key, {})
        rs_calc = calculate_rs_ratio_series(df_all, bm_info.get("df", pd.DataFrame()), 60, 20)
        if not rs_calc.empty and "rs_ratio" in rs_calc.columns:
            vr, vm = rs_calc["rs_ratio"].dropna(), rs_calc["rs_momentum"].dropna()
            rs_r = round(float(vr.iloc[-1]), 2) if not vr.empty else 100.0
            rs_m = round(float(vm.iloc[-1]), 2) if not vm.empty else 100.0
        else:
            rs_r = round(100.0 * (1.0 + r1q / 100.0) / max(0.01, (1.0 + bm_info.get("r_60d", 0.0) / 100.0)), 2)
            rs_m = round(100.0 * (1.0 + r1m / 100.0) / max(0.01, (1.0 + bm_info.get("r_20d", 0.0) / 100.0)), 2)
        return cur, max_h, ma20, r5, r1m, r1q, rs_r, rs_m
    except Exception: return None, None, None, 0.0, 0.0, 0.0, 100.0, 100.0

def calc_pnl(shares, avg_cost, current_price, discount):
    fee = 0.001425 * discount
    t_cost, t_sell = (shares * avg_cost) * (1 + fee), (shares * current_price) * (1 - fee - 0.003)
    pnl = round(t_sell - t_cost)
    return pnl, round((pnl / t_cost) * 100, 2) if t_cost > 0 else 0.0, round(avg_cost * (1 + fee * 2 + 0.003), 2)

@st.cache_data(ttl=3600)
def compute_market_breadth_data(market_list, mkt_filter="TW"):
    filtered = [f"{clean_sym(it.get('symbol','')).upper()}.{'TWO' if '上櫃' in str(it.get('market','')) or 'TWO' in str(it.get('market','')).upper() else 'TW'}" for it in market_list if mkt_filter in ("ALL", "TWO" if "上櫃" in str(it.get("market","")) or "TWO" in str(it.get("market","")).upper() else "TW") and clean_sym(it.get("symbol",""))]
    if not filtered: return None
    try:
        bm_hist = yf.Ticker("^TWII" if mkt_filter == "TW" else "^TWOII").history(period="1y")
        if bm_hist.empty: bm_hist = yf.Ticker("0050.TW").history(period="1y")
        bm_clean = _clean_date_series(bm_hist).set_index("Date")
        data = yf.download(filtered, period="1y", interval="1d", group_by="column", auto_adjust=True, progress=False)
    except Exception: return None
    if data.empty: return None

    closes, highs, lows = (data[k].to_frame() if isinstance(data[k], pd.Series) else data[k] for k in ["Close", "High", "Low"])
    closes, highs, lows = closes.dropna(how="all").ffill(), highs.dropna(how="all").ffill(), lows.dropna(how="all").ffill()
    volumes = (data["Volume"].to_frame() if isinstance(data["Volume"], pd.Series) else data["Volume"]).dropna(how="all").fillna(0)
    if len(closes) < 30: return None

    base_dates = closes.index.strftime("%Y-%m-%d").tolist()
    bm_closes = bm_clean["Close"].reindex(pd.to_datetime(base_dates)).ffill().bfill().values if not bm_clean.empty and "Close" in bm_clean.columns else np.linspace(20000, 23000, len(base_dates))
    bm_vols = bm_clean["Volume"].reindex(pd.to_datetime(base_dates)).ffill().fillna(0).values if not bm_clean.empty and "Volume" in bm_clean.columns else np.zeros(len(base_dates))

    # 1. 均線廣度 (MAB20, MAB60, MAB240)
    ma20, ma60, ma120, ma240 = [closes.rolling(w, min_periods=min(5, w//4)).mean() for w in [20, 60, 120, 240]]
    total_valid = closes.notna().sum(axis=1).replace(0, np.nan)
    calc_ratio = lambda c: (c.sum(axis=1) / total_valid * 100).round(2)
    above_20, above_60, above_240 = calc_ratio(closes > ma20), calc_ratio(closes > ma60), calc_ratio(closes > ma240)

    # 2. 60 日淨創新高指數 (NNH60)
    h60, l60 = highs.rolling(60, min_periods=20).max(), lows.rolling(60, min_periods=20).min()
    nh_60 = (highs >= (h60 - 1e-4)).sum(axis=1)
    nl_60 = (lows <= (l60 + 1e-4)).sum(axis=1)
    nnh_60 = (nh_60 - nl_60).values

    # 3. 20 日突破延續度 (BFTR20)
    prev_h20 = highs.shift(1).rolling(20, min_periods=10).max()
    vol_ma20 = volumes.shift(1).rolling(20, min_periods=5).mean()
    is_breakout = (closes > prev_h20) & ((volumes >= 1.3 * vol_ma20) | (vol_ma20 == 0))
    b_count = is_breakout.shift(2).sum(axis=1)
    s_count = ((closes > prev_h20.shift(2)) & is_breakout.shift(2)).sum(axis=1)
    b_20 = b_count.rolling(20, min_periods=1).sum()
    s_20 = s_count.rolling(20, min_periods=1).sum()
    bftr_20 = np.where(b_20 > 0, (s_20 / b_20) * 100.0, 50.0).round(1)

    # 4. 20 日出貨日計數 (DDC20)
    bm_s_close = pd.Series(bm_closes)
    bm_s_vol = pd.Series(bm_vols)
    is_dist = (bm_s_close.pct_change() <= -0.002) & (bm_s_vol > bm_s_vol.shift(1))
    ddc_20 = is_dist.rolling(20, min_periods=1).sum().fillna(0).astype(int).values

    # 5. 60 日相對強度領頭羊動能 (Top 10% RS Momentum)
    ret_60 = closes.pct_change(60)
    ranks_60 = ret_60.rank(axis=1, pct=True)
    daily_ret = closes.pct_change()
    leader_daily = daily_ret.where(ranks_60 >= 0.90).mean(axis=1).fillna(0)
    leader_cum = (1 + leader_daily).cumprod()
    leader_index = (leader_cum / leader_cum.iloc[0] * 100.0).round(2)
    leader_ma20 = leader_index.rolling(20, min_periods=5).mean().round(2)
    leader_dist_ma20 = (((leader_index - leader_ma20) / leader_ma20) * 100.0).round(2)

    bm_ma60 = pd.Series(bm_closes).rolling(60, min_periods=5).mean()

    return pd.DataFrame({
        "Date": base_dates, "close": bm_closes,
        "above_20ma": above_20.values, "above_60ma": above_60.values, "above_240ma": above_240.values,
        "nh_60": nh_60.values, "nl_60": nl_60.values, "net_high_low_60": nnh_60,
        "nh_60_ratio": (nh_60 / total_valid * 100).round(2).values,
        "nl_60_ratio": (nl_60 / total_valid * 100).round(2).values,
        "bftr_20": bftr_20, "ddc_20": ddc_20,
        "leader_index": leader_index.values, "leader_ma20": leader_ma20.values, "leader_dist_ma20": leader_dist_ma20.values,
        "dist_60ma_pct": (((pd.Series(bm_closes) - bm_ma60) / bm_ma60) * 100.0).round(2).values,
        "short_bull_ratio": calc_ratio((closes > ma20) & (ma20 > ma60)).values,
        "long_bull_ratio": calc_ratio((closes > ma20) & (ma20 > ma60) & (ma60 > ma120) & (ma120 > ma240)).values
    }).set_index("Date")

def render_metric_grid(pairs):
    for (l1, v1, d1), (l2, v2, d2) in pairs:
        c1, c2 = st.columns(2)
        c1.metric(l1, v1, d1)
        c2.metric(l2, v2, d2)

# ==========================================
# 順勢交易生命週期狀態機與 Watchlist 模組
# ==========================================
WATCHLIST_COLS = ["symbol", "name", "theme", "created_date", "stage", "prev_stage", "substate", "base_count", "pivot_price", "strategy_tranches", "transition_date", "is_active"]

STRATEGY_TEMPLATES = {
    "醞釀期 (Incubation / VCP)": {
        "tranches": 2, "ratios": "各 50%", 
        "rule": "第 1 批：極度窒息量（VDU）收縮試單；第 2 批：放量突破頸線樞紐加碼",
        "stop_loss_type": "整理區最新收縮波段低點 (Swing Low)"
    },
    "初升段 (Stage 2A Breakout)": {
        "tranches": 2, "ratios": "各 50%", 
        "rule": "第 1 批：長紅實體 >=3% 帶量突破進場；第 2 批：回測守穩頸線/5MA 補滿",
        "stop_loss_type": "突破長紅低點或起漲 20MA"
    },
    "主升段 (Stage 2B Trend)": {
        "tranches": 1, "ratios": "100% (一次到位)", 
        "rule": "嚴格多頭排列，縮量回測 10MA/20MA 守穩轉強單筆進場",
        "stop_loss_type": "20MA 下方 1~2 檔或前波段低點"
    },
    "末升段 (Stage 3 Climax)": {
        "tranches": 0, "ratios": "0% (禁止追買)", 
        "rule": "嚴禁開新倉；持有者啟動 5MA 緊縮移動停利（跌破先調節半數）",
        "stop_loss_type": "跌破 5MA/10MA 強制停利"
    },
    "出貨期 (Distribution)": {
        "tranches": 0, "ratios": "0% (全面防守)", 
        "rule": "嚴禁買進；全面清空持股現金為王",
        "stop_loss_type": "無條件停損出場"
    },
    "打底期 (Stage 1 Basing / Idle)": {
        "tranches": 0, "ratios": "0% (觀望等待)", 
        "rule": "均線無方向性，等待收斂或右側轉強訊號",
        "stop_loss_type": "無"
    }
}

def load_watchlist():
    conn = _get_gsheet_conn()
    if conn:
        try:
            df = conn.read(worksheet="watchlist", ttl=0)
            if df is not None and not df.empty:
                for c in WATCHLIST_COLS:
                    if c not in df.columns: df[c] = ""
                active_mask = df["is_active"].astype(str).str.lower() == "true"
                return df[active_mask].to_dict("records")
        except Exception: pass
    return []

def save_to_watchlist(record):
    conn = _get_gsheet_conn()
    if conn:
        try:
            try: df = conn.read(worksheet="watchlist", ttl=0)
            except Exception: df = pd.DataFrame(columns=WATCHLIST_COLS)
            if df is None or df.empty: df = pd.DataFrame(columns=WATCHLIST_COLS)
            for c in WATCHLIST_COLS:
                if c not in df.columns: df[c] = ""
            sym = str(record.get("symbol", "")).strip()
            df["symbol"] = df["symbol"].astype(str).str.strip()
            if sym in df["symbol"].values:
                idx = df[df["symbol"] == sym].index[0]
                old_stg = str(df.at[idx, "stage"])
                if old_stg and old_stg != record.get("stage"):
                    record["prev_stage"] = old_stg
                    record["transition_date"] = get_tw_now_str("%Y-%m-%d")
                for k, v in record.items(): df.at[idx, k] = v
            else:
                record.setdefault("prev_stage", "")
                record.setdefault("transition_date", get_tw_now_str("%Y-%m-%d"))
                df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
            conn.update(worksheet="watchlist", data=df)
            return True
        except Exception as e:
            st.error(f"Watchlist 寫入失敗: {e}")
    return False

def remove_from_watchlist(sym):
    conn = _get_gsheet_conn()
    if conn:
        try:
            df = conn.read(worksheet="watchlist", ttl=0)
            if df is not None and not df.empty:
                df["symbol"] = df["symbol"].astype(str).str.strip()
                df = df[df["symbol"] != str(sym).strip()]
                conn.update(worksheet="watchlist", data=df)
                return True
        except Exception as e:
            st.error(f"剔除失敗: {e}")
    return False

def calculate_technical_features(df_daily: pd.DataFrame, df_weekly: pd.DataFrame) -> dict:
    c, v, h, l = df_daily['Close'], df_daily['Volume'], df_daily['High'], df_daily['Low']
    ma = {f"ma{d}": c.rolling(d, min_periods=min(len(c), max(3, d // 4))).mean() for d in [5, 20, 60, 130, 260]}
    w_close = df_weekly['Close'] if not df_weekly.empty and 'Close' in df_weekly.columns else c
    w_ma = {f"wma{w}": w_close.rolling(w, min_periods=min(len(w_close), max(2, w // 4))).mean() for w in [4, 12, 26, 52]}

    slope_20 = (ma['ma20'].iloc[-1] - ma['ma20'].iloc[-6]) / ma['ma20'].iloc[-6] if len(ma['ma20']) >= 6 and ma['ma20'].iloc[-6] > 0 else 0.0
    slope_60 = (ma['ma60'].iloc[-1] - ma['ma60'].iloc[-11]) / ma['ma60'].iloc[-11] if len(ma['ma60']) >= 11 and ma['ma60'].iloc[-11] > 0 else 0.0
    slope_130 = (ma['ma130'].iloc[-1] - ma['ma130'].iloc[-11]) / ma['ma130'].iloc[-11] if len(ma['ma130']) >= 11 and ma['ma130'].iloc[-11] > 0 else 0.0
    slope_260 = (ma['ma260'].iloc[-1] - ma['ma260'].iloc[-21]) / ma['ma260'].iloc[-21] if len(ma['ma260']) >= 21 and ma['ma260'].iloc[-21] > 0 else 0.0

    tr = pd.concat([h - l, (h - c.shift(1)).abs(), (l - c.shift(1)).abs()], axis=1).max(axis=1)
    atr10 = tr.rolling(10, min_periods=3).mean().iloc[-1] if len(tr) >= 3 else 1.0
    atr60 = tr.rolling(60, min_periods=10).mean().iloc[-1] if len(tr) >= 10 else atr10
    atr_contraction = atr10 / atr60 if atr60 > 0 else 1.0

    vol_ma20 = v.rolling(20, min_periods=5).mean().iloc[-1] if len(v) >= 5 else 1.0
    vol_ma60 = v.rolling(60, min_periods=10).mean().iloc[-1] if len(v) >= 10 else vol_ma20
    rel_vol_20 = v.iloc[-1] / vol_ma20 if vol_ma20 > 0 else 1.0
    rel_vol_60 = v.iloc[-1] / vol_ma60 if vol_ma60 > 0 else 1.0
    min_vol_5d = v.tail(5).min() if len(v) >= 5 else v.iloc[-1]
    vdu_ratio = min_vol_5d / vol_ma60 if vol_ma60 > 0 else 1.0

    cur_p = float(c.iloc[-1])
    high_60 = float(h.tail(60).max()) if len(h) >= 60 else float(h.max())
    high_260 = float(h.tail(260).max()) if len(h) >= 260 else float(h.max())
    low_260 = float(l.tail(260).min()) if len(l) >= 260 else float(l.min())

    dist_high_60 = (high_60 - cur_p) / high_60 if high_60 > 0 else 0.0
    dist_high_260 = (high_260 - cur_p) / high_260 if high_260 > 0 else 0.0
    rise_from_low_260 = (cur_p - low_260) / low_260 if low_260 > 0 else 0.0

    bias_20 = (cur_p - ma['ma20'].iloc[-1]) / ma['ma20'].iloc[-1] if ma['ma20'].iloc[-1] > 0 else 0.0
    bias_60 = (cur_p - ma['ma60'].iloc[-1]) / ma['ma60'].iloc[-1] if ma['ma60'].iloc[-1] > 0 else 0.0
    bias_260 = (cur_p - ma['ma260'].iloc[-1]) / ma['ma260'].iloc[-1] if ma['ma260'].iloc[-1] > 0 else 0.0

    pct_chg = c.pct_change()
    dist_days = int(((pct_chg <= -0.002) & (v > v.shift(1))).tail(20).sum())
    open_p = float(df_daily['Open'].iloc[-1]) if 'Open' in df_daily.columns else cur_p

    return {
        "c": c, "v": v, "h": h, "l": l, "ma": ma, "w_ma": w_ma,
        "slope_20": slope_20, "slope_60": slope_60, "slope_130": slope_130, "slope_260": slope_260,
        "atr10": atr10, "atr60": atr60, "atr_contraction": atr_contraction,
        "vol_ma20": vol_ma20, "vol_ma60": vol_ma60, "rel_vol_20": rel_vol_20, "rel_vol_60": rel_vol_60,
        "min_vol_5d": min_vol_5d, "vdu_ratio": vdu_ratio,
        "cur_p": cur_p, "high_60": high_60, "high_260": high_260, "low_260": low_260,
        "dist_high_60": dist_high_60, "dist_high_260": dist_high_260, "rise_from_low_260": rise_from_low_260,
        "bias_20": bias_20, "bias_60": bias_60, "bias_260": bias_260, "dist_days": dist_days,
        "open_p": open_p
    }

def evaluate_dow_structure(df_d: pd.DataFrame) -> str:
    if df_d is None or len(df_d) < 20: return "收斂/震盪"
    recent_h = df_d['High'].tail(20).values
    recent_l = df_d['Low'].tail(20).values
    if recent_h[-1] >= np.max(recent_h) * 0.98 and recent_l[-1] > np.min(recent_l):
        return "HH + HL"
    elif recent_l[-1] <= np.min(recent_l) * 1.02 and recent_h[-1] < np.max(recent_h):
        return "LH + LL"
    return "收斂/震盪"

def evaluate_state_machine(feat: dict, dow_structure: str, base_count: int = 1) -> dict:
    cur_p, ma, w_ma, c_series = feat['cur_p'], feat['ma'], feat['w_ma'], feat['c']

    # Step 2: 中斷態檢驗 (Correction Interrupt Check)
    is_markdown = (
        (len(c_series) >= 2 and c_series.iloc[-1] < ma['ma130'].iloc[-1] and c_series.iloc[-2] < ma['ma130'].iloc[-2] and
         c_series.iloc[-1] < ma['ma260'].iloc[-1] and c_series.iloc[-2] < ma['ma260'].iloc[-2]) or
        (ma['ma60'].iloc[-1] < ma['ma130'].iloc[-1] and feat['slope_60'] < 0) or
        (dow_structure == "LH + LL")
    )
    if is_markdown:
        return {"stage": "打底期 (Stage 1 Basing / Idle)", "substate": "結構走空 (Markdown)", "base_count": 0, "can_trade": False}

    is_pullback = (
        (cur_p <= ma['ma20'].iloc[-1] * 1.02 and cur_p >= ma['ma60'].iloc[-1]) and
        (ma['ma20'].iloc[-1] > ma['ma60'].iloc[-1] > ma['ma130'].iloc[-1]) and
        (feat['rel_vol_20'] < 0.7) and
        (feat['dist_days'] < 4)
    )
    substate = "良性回檔修正 (Pullback)" if is_pullback else "無"

    # Step 3: 五階段推進主鏈 (P1 -> P5)
    # [P1 出貨期 (Distribution)]
    candle_body_pct = (cur_p - feat['open_p']) / feat['open_p'] if feat['open_p'] > 0 else 0
    p1_cond = (
        (feat['dist_days'] >= 4) or
        (feat['rel_vol_20'] >= 1.8 and candle_body_pct < 0.005) or
        (cur_p < ma['ma20'].iloc[-1] and ma['ma5'].iloc[-1] < ma['ma20'].iloc[-1])
    )
    if p1_cond:
        return {"stage": "出貨期 (Distribution)", "substate": substate, "base_count": base_count}

    # [P2 末升段 (Stage 3 Climax)]
    ret_10d = (cur_p - c_series.iloc[-11]) / c_series.iloc[-11] if len(c_series) >= 11 else 0
    p2_cond = (
        (cur_p / ma['ma60'].iloc[-1] > 1.25 or cur_p / ma['ma260'].iloc[-1] > 1.60) or
        (ret_10d > 0.30 and cur_p / ma['ma20'].iloc[-1] > 1.15)
    )
    if p2_cond:
        return {"stage": "末升段 (Stage 3 Climax)", "substate": substate, "base_count": base_count}

    # [P3 主升段 (Stage 2B Trend)]
    p3_cond = (
        (cur_p > ma['ma20'].iloc[-1] > ma['ma60'].iloc[-1] > ma['ma130'].iloc[-1] > ma['ma260'].iloc[-1]) and
        (feat['slope_20'] > 0 and feat['slope_60'] > 0 and feat['slope_130'] > 0) and
        (w_ma['wma4'].iloc[-1] > w_ma['wma12'].iloc[-1] > w_ma['wma26'].iloc[-1] > w_ma['wma52'].iloc[-1]) and
        (feat['rise_from_low_260'] >= 0.30 and feat['dist_high_260'] <= 0.15) and
        (dow_structure == "HH + HL")
    )
    if p3_cond:
        return {"stage": "主升段 (Stage 2B Trend)", "substate": substate, "base_count": base_count}

    # [P4 初升段 (Stage 2A Breakout)]
    bull_k_body = (cur_p - feat['open_p']) / feat['open_p'] if feat['open_p'] > 0 else 0
    p4_cond = (
        (cur_p >= feat['high_60'] * 0.99 and bull_k_body >= 0.03) and
        (feat['rel_vol_20'] >= 1.5 or feat['rel_vol_60'] >= 2.0) and
        (cur_p > ma['ma5'].iloc[-1] > ma['ma20'].iloc[-1] > ma['ma60'].iloc[-1]) and
        (feat['slope_20'] > 0 and feat['slope_60'] > 0)
    )
    if p4_cond:
        return {"stage": "初升段 (Stage 2A Breakout)", "substate": substate, "base_count": max(1, base_count)}

    # [P5 醞釀期 (Incubation / VCP)]
    ma_conv = abs(ma['ma20'].iloc[-1] - ma['ma60'].iloc[-1]) / ma['ma60'].iloc[-1] if ma['ma60'].iloc[-1] > 0 else 1.0
    p5_cond = (
        (cur_p > ma['ma60'].iloc[-1] and cur_p > ma['ma130'].iloc[-1]) and
        (feat['slope_130'] >= 0 and feat['slope_260'] >= -0.002) and
        (ma_conv < 0.03) and
        (feat['vdu_ratio'] <= 0.5) and
        (feat['atr_contraction'] < 0.65 and feat['dist_high_60'] < 0.05)
    )
    if p5_cond:
        return {"stage": "醞釀期 (Incubation / VCP)", "substate": substate, "base_count": base_count}

    # Step 4: 基底兜底態 (Fallback: Basing)
    return {"stage": "打底期 (Stage 1 Basing / Idle)", "substate": substate, "base_count": base_count}

def map_theme_to_stocks(theme_prompt: str, client: genai.Client, model_name="gemini-2.5-flash") -> list:
    prompt = f"""
    你是一名精通台股上市櫃產業供應鏈的資深研究員。
    請分析使用者給定的題材或問題：「{theme_prompt}」
    列出最直接受惠、具代表性的台灣上市/上櫃個股（最多 8~10 檔）。

    請嚴格返回 JSON Array 格式，不要有額外的 Markdown 說明：
    [
      {{
        "symbol": "3450",
        "name": "聯鈞",
        "market": "TW",
        "relevance": "核心受惠 (高純度)",
        "business_role": "矽光子雷射雷射模組封裝，受惠資料中心升級"
      }}
    ]
    """
    try:
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"題材映射失敗: {e}")
        return []

def run_auto_scan_pipeline(theme: str, target_stages: list = None) -> str:
    """自動掃描指定台股題材或產業鏈，以量化狀態機篩選符合特定生命週期階段的個股，並自動寫入 Watchlist 進行追蹤。

    Args:
        theme: 欲掃描的產業題材關鍵字，例如「CPO」、「矽光子」、「低軌衛星」。
        target_stages: 欲篩選保留的階段名稱清單，例如 ["初升段 (Stage 2A Breakout)", "醞釀期 (Incubation / VCP)"]。若未提供則預設為初升段與醞釀期。
    """
    if not target_stages:
        target_stages = ["初升段 (Stage 2A Breakout)", "醞釀期 (Incubation / VCP)"]

    api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
    target_model = st.secrets.get("GEMINI_MODEL", "gemini-2.5-flash")
    if not api_key:
        return "⚠️ 未設定 GEMINI_API_KEY，無法執行題材標的映射。"

    client = genai.Client(api_key=api_key)
    stocks = map_theme_to_stocks(theme, client, target_model)
    if not stocks:
        return f"⚠️ 未能找到與題材「{theme}」相關的台股標的。"

    matched_records = []
    tw_today = get_tw_now_str("%Y-%m-%d")

    for stk in stocks:
        sym = clean_sym(stk.get("symbol", ""))
        if not sym: continue
        name = clean_stock_name(stk.get("name", sym), sym)
        mkt = str(stk.get("market", "TW")).upper()
        ticker_str = f"{sym}.TWO" if "TWO" in mkt or "上櫃" in mkt else f"{sym}.TW"

        try:
            df_d = yf.Ticker(ticker_str).history(period="14mo")
            if df_d.empty:
                alt_str = f"{sym}.TW" if "TWO" in ticker_str else f"{sym}.TWO"
                df_d = yf.Ticker(alt_str).history(period="14mo")

            if df_d.empty or len(df_d) < 60:
                continue

            df_w = df_d.resample('W-FRI').agg({
                'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'
            }).dropna()

            feat = calculate_technical_features(df_d, df_w)
            dow = evaluate_dow_structure(df_d)
            eval_res = evaluate_state_machine(feat, dow, base_count=1)
            stage_name = eval_res["stage"]
            substate = eval_res["substate"]

            is_matched = False
            for tg in target_stages:
                tg_str = str(tg).strip()
                if tg_str in stage_name:
                    is_matched = True
                    break
                if "初升" in tg_str and "初升" in stage_name:
                    is_matched = True
                    break
                if ("醞釀" in tg_str or "VCP" in tg_str.upper()) and ("醞釀" in stage_name or "VCP" in stage_name.upper()):
                    is_matched = True
                    break
                if "主升" in tg_str and "主升" in stage_name:
                    is_matched = True
                    break
                if "末升" in tg_str and "末升" in stage_name:
                    is_matched = True
                    break
                if "出貨" in tg_str and "出貨" in stage_name:
                    is_matched = True
                    break
                if "打底" in tg_str and "打底" in stage_name:
                    is_matched = True
                    break

            if is_matched:
                strat = STRATEGY_TEMPLATES.get(stage_name, STRATEGY_TEMPLATES.get("打底期 (Stage 1 Basing / Idle)"))
                rec = {
                    "symbol": sym,
                    "name": name,
                    "theme": theme,
                    "created_date": tw_today,
                    "stage": stage_name,
                    "prev_stage": "",
                    "substate": substate,
                    "base_count": eval_res.get("base_count", 1),
                    "pivot_price": round(feat["high_60"], 2),
                    "strategy_tranches": json.dumps(strat, ensure_ascii=False),
                    "transition_date": tw_today,
                    "is_active": True
                }
                save_to_watchlist(rec)
                matched_records.append({
                    "name": name,
                    "symbol": sym,
                    "stage": stage_name,
                    "substate": substate,
                    "cur_p": feat["cur_p"],
                    "pivot": round(feat["high_60"], 2)
                })
        except Exception:
            continue

    if matched_records:
        lines = [f"- **{r['name']} ({r['symbol']})** ｜ 現價 `${r['cur_p']:.2f}` ｜ 階段：`{r['stage']}` ｜ 子狀態：`{r['substate']}` ｜ 突破樞紐：`${r['pivot']:.2f}`" for r in matched_records]
        return f"🎯 **【{theme}】題材掃描完成！共篩選出 {len(matched_records)} 檔符合條件標的，已自動寫入 Google Sheets Watchlist：**\n\n" + "\n".join(lines) + "\n\n💡 *可切換至「🎯 題材診斷與追蹤」分頁檢視最新 Watchlist 清單。*"
    else:
        return f"🔍 **【{theme}】題材掃描完成**，相關供應鏈標的目前皆未處於指定的階段條件。"

def render_transition_status_cards():
    conn = _get_gsheet_conn()
    if not conn: return
    try: df = conn.read(worksheet="watchlist", ttl=0)
    except Exception: return
    if df is None or df.empty or "stage" not in df.columns or "prev_stage" not in df.columns: return

    is_active_col = df["is_active"].astype(str).str.lower() == "true"
    stage_changed = (df["stage"].astype(str) != df["prev_stage"].astype(str)) & df["prev_stage"].notna() & (df["prev_stage"].astype(str) != "")
    alerts = df[is_active_col & stage_changed].copy()
    if alerts.empty: return

    critical_items, climax_items, breakout_items = [], [], []
    for _, row in alerts.iterrows():
        info = {
            "symbol": str(row.get("symbol", "")), "name": str(row.get("name", "")),
            "prev": str(row.get("prev_stage", "未知")), "cur": str(row.get("stage", "未知")),
            "date": str(row.get("transition_date", "")), "substate": str(row.get("substate", "無"))
        }
        c_stg, sub = info["cur"], info["substate"]
        if "出貨" in c_stg or "走空" in sub or "Markdown" in sub: critical_items.append(info)
        elif "末升" in c_stg: climax_items.append(info)
        elif "初升" in c_stg or "主升" in c_stg: breakout_items.append(info)

    if critical_items:
        card_c = "\n".join([f"- **{it['name']} ({it['symbol']})**：由 `{it['prev']}` ➔ **{it['cur']}**（子狀態: `{it['substate']}`，躍遷日: {it['date']}）  \n  ↳ **策略防禦：嚴禁買進，持股啟動強制停損/出清！**" for it in critical_items])
        st.error(f"🚨 **【危險防守・結構破壞/出貨告警】**\n\n{card_c}")
    if climax_items:
        card_c = "\n".join([f"- **{it['name']} ({it['symbol']})**：由 `{it['prev']}` ➔ **{it['cur']}**（躍遷日: {it['date']}）  \n  ↳ **策略執行：禁止追高，啟動 5MA 緊縮移動停利！**" for it in climax_items])
        st.warning(f"⚠️ **【高檔過熱・末升段噴出告警】**\n\n{card_c}")
    if breakout_items:
        card_c = "\n".join([f"- **{it['name']} ({it['symbol']})**：由 `{it['prev']}` ➔ **{it['cur']}**（躍遷日: {it['date']}）  \n  ↳ **策略執行：完成底部整理，觸發第 1 批 50% 資金進場！**" for it in breakout_items])
        st.success(f"🎯 **【動能發動・初升段突破確認】**\n\n{card_c}")

# ==========================================
# 介面渲染
# ==========================================
market_rankings, db_status = load_market_data()
st.title("🚀 台股儀表板")

# 渲染狀態躍遷即時字卡
render_transition_status_cards()

with st.expander("🛡️ 說明", expanded=False):
    st.markdown("**RS_ratio 雙軸指標**：以 60 日季線為強弱中軸（≥100 為 🔥[強勢]，<100 為 ❄️[弱勢]）；以 20 日 SMA 為短線動能加速度。")
    r1, r2 = st.columns(2)
    r1.markdown("**1. 🔴 初始停損**：跌破預設趴數無條件停損。\n\n**2. 🛡️ 保本停損**：獲利達標鎖定零虧損。\n\n**3. 🟣 高點回檔**：自高點拉回觸發分批停利。")
    r2.markdown("**4. 🟠 月線過熱**：20MA 正乖離過大建議調節。\n\n**5. ⏳ 時間停損**：持股過久動能停滯建議換股。")

if market_rankings: st.info(f"🟢 **全市場 RS 資料庫已就緒** ｜ 收錄 **{len(market_rankings)}** 檔台股 ｜ 狀態：{db_status}")
else: st.warning("🟡 正在等待全市場 RS 排名資料載入...")

tab_portfolio, tab_leaderboard, tab_theme, tab_market_breadth, tab_ai = st.tabs(["📈 獲利監控系統", "🏆 個股查詢", "🎯 題材診斷與追蹤", "📊 大盤", "🤖 Gemini 智能助理"])
portfolio_live_summary = []

with tab_portfolio:
    with st.expander("⚙️ 參數設定", expanded=False):
        c1, c2 = st.columns(2)
        stop_loss_pct = c1.number_input("🔴 初始停損趴數 (%)", 1.0, 50.0, 7.0, 0.5, format="%.1f")
        breakeven_trigger_pct = c1.number_input("🛡️ 保本停損啟動門檻 (%)", 1.0, 50.0, 8.0, 0.5, format="%.1f")
        pullback_target_pct = c1.number_input("🟣 高點回檔停利趴數 (%)", 1.0, 50.0, 10.0, 0.5, format="%.1f")
        bias_threshold = c2.number_input("🟠 月線正乖離過熱閥值 (%)", 5.0, 100.0, 30.0, 1.0, format="%.0f")
        time_stop_days = c2.number_input("⏳ 時間停損天數（天）", 1, 100, 10, 1)
        discount_display = c2.number_input("💰 券商手續費折數", 0.01, 1.0, 0.60, 0.05, format="%.2f")

    portfolio = load_data()

    with st.expander("➕ 新增持股", expanded=False):
        with st.form("add_stock_form"):
            fc1, fc2 = st.columns(2)
            sym = fc1.text_input("股票代號", placeholder="例如: 3441 或 2330")
            name = fc1.text_input("股票名稱", placeholder="例如: 聯一光")
            mkt = fc1.selectbox("市場別", ["TWO (上櫃)", "TW (上市)"])
            entry_d = fc2.date_input("進場日期", value=get_tw_now().date())
            price = fc2.number_input("買進價格", min_value=0.1, step=0.1, value=100.0)
            shs = fc2.number_input("買進股數", min_value=1, step=1000, value=1000)
                
            if st.form_submit_button("確認建立持倉", use_container_width=True) and sym:
                sym_clean = clean_sym(sym)
                clean_n = clean_stock_name(name.strip(), sym_clean) if name else clean_stock_name(sym_clean, sym_clean)
                portfolio.append({
                    "symbol": sym_clean, "name": clean_n, "market": "TWO" if "TWO" in mkt else "TW",
                    "entry_date": str(entry_d), "avg_cost": float(price), "shares": int(shs),
                    "record_high": float(price), "realized_pnl": 0.0, "status_override": "",
                    "history": [make_log_entry("🌱 初始建倉", price, f"+{int(shs)}", int(shs), "0 元", f"起始成本 ${price}")]
                })
                save_data(portfolio)
                st.success(f"已新增 {clean_n} ({sym_clean})")
                st.rerun()

    if not portfolio:
        st.info("目前尚無持倉，請點擊上方「➕ 新增持股」建立第一檔股票。")
    else:
        if st.button("🔄 刷新資料", use_container_width=True):
            st.cache_data.clear()
            st.session_state.last_portfolio_refresh = get_tw_now_str()
            st.rerun()
        st.caption(f"🕒 最新市價更新時間：{st.session_state.last_portfolio_refresh}")

        for idx, item in enumerate(portfolio):
            sym, name, mkt, entry_d = clean_sym(item["symbol"]), clean_stock_name(item.get("name", ""), item.get("symbol")), item["market"], item["entry_date"]
            avg_cost, shares, stored_high = item["avg_cost"], item["shares"], item.get("record_high", item["avg_cost"])
            realized_pnl, history_logs = item.get("realized_pnl", 0.0), item.get("history", [])

            info = next((it for it in market_rankings if clean_sym(it.get("symbol", "")).upper() == sym.upper()), None)
            rs_score = info.get("rs_rating", 50) if info else 50
            cur_price, max_high, ma20, r_5d, r_1m, r_1q, rs_ratio_val, rs_mom_val = fetch_stock_and_momentum(sym, mkt, entry_d)
            if cur_price is None: cur_price, max_high, ma20 = avg_cost, stored_high, avg_cost

            actual_high = max(stored_high, avg_cost, max_high or stored_high)
            if actual_high != stored_high:
                portfolio[idx]["record_high"] = actual_high
                save_data(portfolio)

            net_pnl, roi, breakeven_p = calc_pnl(shares, avg_cost, cur_price, discount_display)
            pullback_pct = round(((actual_high - cur_price) / actual_high) * 100, 1) if actual_high > 0 else 0
            bias_20 = round(((cur_price - ma20) / ma20) * 100, 1) if ma20 > 0 else 0
            try: days_held = (get_tw_now().date() - datetime.strptime(entry_d, "%Y-%m-%d").date()).days
            except Exception: days_held = 0

            status_item = (info.copy() if info else {"rs_rating": rs_score, "pattern_badge": "", "r_5d": r_5d})
            status_item["rs_ratio"] = rs_ratio_val
            status_badge = get_trend_master_status(status_item)

            is_breakeven_active = ((actual_high - avg_cost) / avg_cost) * 100 >= breakeven_trigger_pct
            init_stop = round(avg_cost * (1 - stop_loss_pct / 100), 2)
            effective_stop = max(init_stop, breakeven_p) if is_breakeven_active else init_stop
            pullback_p = round(actual_high * (1 - pullback_target_pct / 100), 2)

            status_text, status_color = "⚪ 持股續抱中", "gray"
            if item.get("status_override") == "持股續抱中":
                status_text, status_color = "⚪ 持股續抱中 (已執行減碼調節)", "green"
            elif cur_price <= effective_stop:
                status_text, status_color = f"🛡️ 觸發保本出場線（{effective_stop} 元）！強制保護本金零虧損出場" if is_breakeven_active else f"🔴 觸發 -{stop_loss_pct}% 停損線（{effective_stop} 元）！全數出場", "red"
            elif cur_price <= pullback_p and cur_price > avg_cost:
                status_text, status_color = f"🟣 觸發高點回檔 {pullback_target_pct}%（跌破 {pullback_p} 元）！建議減碼", "purple"
            elif bias_20 >= bias_threshold:
                status_text, status_color = f"🟠 月線正乖離達 {bias_20}%（過熱）！建議減碼", "orange"
            elif days_held >= time_stop_days and abs(roi) <= 2.0:
                status_text, status_color = f"⏳ 觸發時間停損（持股已 {days_held} 天，動能停滯）！建議換股", "orange"

            portfolio_live_summary.append({
                "股票": f"{name} ({sym})", "持股數": shares, "成本價": avg_cost, "現價": cur_price,
                "未實現損益": net_pnl, "報酬率%": roi, "RS評分": rs_score, "RS_ratio": rs_ratio_val,
                "狀態": status_text, "持有天數": days_held
            })

            with st.container():
                st.divider()
                st.subheader(f"{name} ({sym}.{mkt}) ｜ 📦 {shares:,} 股 ｜ {status_badge}")
                render_metric_grid([
                    (("RS Rating 評分", f"{rs_score} 分", None), ("RS動能比率(20MA)", f"{rs_mom_val}", "🔥 短期動能增強" if rs_mom_val >= 100 else "❄️ 動能未達臨界點")),
                    (("RS_ratio 比率 (60MA)", f"{rs_ratio_val}", "🔥 超越大盤" if rs_ratio_val >= 100 else "❄️ 落後大盤"), ("近 5 日動能", f"{r_5d:+}%", None)),
                    (("近 20 日動能", f"{r_1m:+}%", None), ("近 60 日動能", f"{r_1q:+}%", None)),
                    (("高點回檔", f"${actual_high}", f"-{pullback_pct}%"), ("最新市價", f"${cur_price}", None)),
                    (("剩餘股數 / 均價", f"{shares:,} 股", f"均價: ${avg_cost}"), ("未實現損益", f"{net_pnl:+,} 元", f"{roi:+}%")),
                    (("累積已實現損益", f"{realized_pnl:+,} 元", None), ("🛡️ 保本停損線" if is_breakeven_active else f"🔴 初始停損 (-{stop_loss_pct}%)", f"${effective_stop}", None))
                ])

                st.markdown(f"**風控狀態：** :{status_color}[{status_text}]")

                with st.expander(f"⚙️ 操作 {name}（加碼 / 減碼 / 結清）"):
                    st.write("##### 🔼 順勢加碼")
                    add_p = st.number_input("加碼價格", min_value=0.1, step=0.1, value=cur_price, key=f"add_p_{idx}")
                    add_s = st.number_input("加碼股數", min_value=1, step=100, value=1000, key=f"add_s_{idx}")
                    new_tot = shares + int(add_s)
                    sim_avg = round(((shares * avg_cost) + (int(add_s) * add_p)) / new_tot, 2)
                    buf = round(((cur_price - sim_avg) / cur_price) * 100, 1)
                    st.caption(f"試算新均價：**${sim_avg}** ｜ 安全緩衝：**{buf:+}%**")
                    if st.button("確認加碼", key=f"btn_add_{idx}", use_container_width=True):
                        portfolio[idx].setdefault("history", []).append(make_log_entry("🔼 順勢加碼", add_p, f"+{int(add_s)}", new_tot, "-", f"新均價 ${sim_avg} (緩衝 {buf:+}%)"))
                        portfolio[idx]["shares"], portfolio[idx]["avg_cost"], portfolio[idx]["status_override"] = new_tot, sim_avg, ""
                        save_data(portfolio)
                        st.rerun()

                    st.divider()
                    st.write("##### 🔽 分批減碼")
                    red_p = st.number_input("減碼價格", min_value=0.1, step=0.1, value=cur_price, key=f"red_p_{idx}")
                    red_s = st.number_input("減碼股數", min_value=1, max_value=shares, step=100, value=min(1000, shares), key=f"red_s_{idx}")
                    
                    default_reason_idx = 0
                    if "高點回檔" in status_text: default_reason_idx = 1
                    elif "正乖離" in status_text or "過熱" in status_text: default_reason_idx = 2
                    elif "時間停損" in status_text: default_reason_idx = 3
                    elif "停損" in status_text or "保本" in status_text: default_reason_idx = 4
                    
                    risk_reasons = ["🎯 自行主動調節", "🟣 高點回檔停利", "🟠 月線正乖離過熱", "⏳ 時間停損換股", "🔴 保本/停損觸發", "📦 其他策略調節"]
                    selected_reason = st.selectbox("減碼風控原因", risk_reasons, index=default_reason_idx, key=f"risk_reason_{idx}")

                    sim_red_pnl, sim_red_roi, _ = calc_pnl(int(red_s), avg_cost, red_p, discount_display)
                    st.caption(f"試算本次損益：**{sim_red_pnl:+,} 元** ({sim_red_roi:+}%)")
                    if st.button("確認減碼", key=f"btn_red_{idx}", use_container_width=True):
                        new_shares = shares - int(red_s)
                        note_text = f"【{selected_reason}】報酬率 {sim_red_roi:+}%"
                        portfolio[idx].setdefault("history", []).append(make_log_entry("🔽 分批減碼", red_p, f"-{int(red_s)}", new_shares, f"{sim_red_pnl:+,} 元", note_text))
                        if new_shares > 0:
                            portfolio[idx]["shares"] = new_shares
                            portfolio[idx]["realized_pnl"] = item.get("realized_pnl", 0.0) + sim_red_pnl
                            portfolio[idx]["status_override"] = "持股續抱中"
                        else:
                            portfolio.pop(idx)
                        save_data(portfolio)
                        st.rerun()

                    st.divider()
                    if st.button("🗑️ 結清出場", key=f"del_{idx}", use_container_width=True):
                        portfolio.pop(idx)
                        save_data(portfolio)
                        st.rerun()

                if history_logs:
                    with st.expander(f"📜 {name} 交易歷程", expanded=False):
                        st.dataframe(pd.DataFrame(history_logs), use_container_width=True, hide_index=True)

with tab_leaderboard:
    st.subheader("🔍 個股查詢")
    typed_search = st.text_input("輸入股票代號或名稱查詢（支援單檔或多檔，多檔請用空白、逗號或換行分隔）", placeholder="例如：2330 聯一光 3441 2454", key="typed_search_field")
    search_query = typed_search.strip() or st.session_state.search_input_val

    if search_query:
        if not typed_search.strip() and st.session_state.search_input_val:
            st.info(f"📌 目前正檢視排行榜點選之標的：**{search_query}** （如需搜尋其他標的，請直接在上方輸入框輸入）")
        matched_dict = {}
        for tok in [t.strip() for t in re.split(r"[\s,;，、\n]+", search_query) if t.strip()]:
            q_token = clean_sym(tok).upper()
            found = False
            for item in market_rankings:
                s_i, n_i = clean_sym(item.get("symbol", "")).upper(), str(item.get("name", "")).upper()
                if q_token in (s_i, n_i) or q_token in s_i or q_token in n_i:
                    matched_dict[s_i] = item
                    found = True
            if not found and (q_token.isdigit() or len(q_token) >= 2):
                std_n = clean_stock_name(q_token, q_token)
                matched_dict[q_token] = {"symbol": q_token, "name": std_n if std_n != q_token else q_token, "market": "TW", "rs_rating": 50, "score": 0.0}

        matched = list(matched_dict.values())
        if matched:
            st.write(f"找到 **{len(matched)}** 筆符合標的：")
            compare_rows, detailed_data = [], []
            for m in matched:
                score, m_type = m.get("rs_rating", 50), m.get("market", "上市/上櫃")
                sym, name = clean_sym(m.get("symbol")), clean_stock_name(m.get("name", m.get("symbol")), m.get("symbol"))
                
                cur_p, _, _, q_r5, q_r20, q_r60, query_rs_ratio, query_rs_mom = fetch_stock_and_momentum(sym, m_type, get_tw_now_str("%Y-%m-%d"))
                if "rs_ratio" in m: query_rs_ratio = m["rs_ratio"]
                if "rs_momentum" in m: query_rs_mom = m["rs_momentum"]
                if "r_5d" in m: q_r5 = float(m.get("r_5d", 0.0) or 0.0)
                if "r_20d" in m: q_r20 = float(m.get("r_20d", 0.0) or 0.0)
                if "r_60d" in m: q_r60 = float(m.get("r_60d", 0.0) or 0.0)

                badge_style = get_trend_master_status(m)
                detailed_data.append({"name": name, "sym": sym, "m_type": m_type, "score": score, "query_rs_mom": query_rs_mom, "query_rs_ratio": query_rs_ratio, "q_r5": q_r5, "q_r20": q_r20, "q_r60": q_r60, "badge_style": badge_style})
                compare_rows.append({"股票代號": sym, "股票名稱": name, "市場別": m_type, "目前市價": f"${cur_p:.2f}" if cur_p is not None else "-", "RS 評分": score, "RS 動能 (20MA)": query_rs_mom, "RS_ratio (60MA)": query_rs_ratio, "5日漲跌幅 (%)": f"{q_r5:+0.2f}%", "20日漲跌幅 (%)": f"{q_r20:+0.2f}%", "60日漲跌幅 (%)": f"{q_r60:+0.2f}%", "動能狀態": badge_style})

            st.markdown("#### 📊 查詢標的數值比較表")
            st.dataframe(pd.DataFrame(compare_rows), use_container_width=True, hide_index=True)
            st.divider()
            st.markdown("#### 📌 查詢標的詳細指標")
            for d in detailed_data:
                st.columns(1)[0].metric("標的與市場", f"{d['name']} ({d['sym']})", f"{d['m_type']} ｜ {d['badge_style']}")
                render_metric_grid([
                    (("RS Rating 評分", f"{d['score']} 分", None), ("RS動能比率(20MA)", f"{d['query_rs_mom']}", "🔥 短期動能增強" if d['query_rs_mom'] >= 100 else "❄️ 短期動能減弱")),
                    (("RS_ratio (60MA)", f"{d['query_rs_ratio']}", "🔥 大盤領先者" if d['query_rs_ratio'] >= 100 else "❄️ 大盤落後者"), ("近 5 日動能", f"{d['q_r5']:+}%", None)),
                    (("近 20 日動能", f"{d['q_r20']:+}%", None), ("近 60 日動能", f"{d['q_r60']:+}%", None))
                ])
                st.divider()
        else: st.error(f"查無符合「{search_query}」的標的。")

    st.subheader("🏆 強勢股排行榜")
    df_raw = pd.DataFrame(market_rankings)
    if not df_raw.empty:
        for c, default_v in [("name", df_raw.get("symbol")), ("market", "上市"), ("rs_ratio", 100.0), ("rs_momentum", 100.0)]:
            if c not in df_raw.columns: df_raw[c] = default_v
        f1, f2 = st.columns(2)
        min_rs = f1.number_input("最低 RS 門檻篩選", 1, 99, 85, 1)
        market_filter = f2.multiselect("市場別篩選", ["上市", "上櫃"], default=["上市", "上櫃"])
        filtered_df = df_raw[(df_raw["rs_rating"] >= min_rs) & (df_raw["market"].isin(market_filter))].copy()
        filtered_df["name"] = filtered_df.apply(lambda r: clean_stock_name(r.get("name"), r.get("symbol")), axis=1)
        filtered_df = filtered_df.sort_values(by="rs_rating", ascending=False).reset_index(drop=True)
        filtered_df["順勢操作狀態"] = filtered_df.apply(get_trend_master_status, axis=1)
        display_df = filtered_df[["rs_rating", "symbol", "name", "market", "score", "rs_ratio", "rs_momentum", "順勢操作狀態"]].rename(columns={"rs_rating": "RS Rating (PR)", "symbol": "股票代號", "name": "中文名稱", "market": "上市櫃", "score": "綜合動能得分", "rs_ratio": "RS_ratio (60MA)", "rs_momentum": "RS動能比率(20MA)"})
        st.caption(f"共計 **{len(display_df)}** 檔標的符合條件（RS ≥ {min_rs}） ｜ 💡 **提示：勾選表格左側任意個股，上方會自動帶出詳細分析**")
        
        col_cfg = {
            "RS Rating (PR)": st.column_config.NumberColumn(width=110),
            "股票代號": st.column_config.TextColumn(width=90),
            "中文名稱": st.column_config.TextColumn(width=120),
            "上市櫃": st.column_config.TextColumn(width=80),
            "綜合動能得分": st.column_config.NumberColumn(width=110, format="%.2f"),
            "RS_ratio (60MA)": st.column_config.NumberColumn(width=130, format="%.2f"),
            "RS動能比率(20MA)": st.column_config.NumberColumn(width=140, format="%.2f"),
            "順勢操作狀態": st.column_config.TextColumn(width=340)
        }
        
        event = st.dataframe(display_df, column_config=col_cfg, use_container_width=True, hide_index=True, height=450, on_select="rerun", selection_mode="single-row", key="rank_df_table")
        if event and hasattr(event, "selection") and event.selection.get("rows"):
            sel_idx = event.selection["rows"][0]
            chosen_sym = str(display_df.iloc[sel_idx]["股票代號"])
            if st.session_state.search_input_val != chosen_sym:
                st.session_state.search_input_val = chosen_sym
                st.rerun()
    else: st.info("尚無排名資料。")

# ==========================================
# 分頁：🎯 題材診斷與追蹤 (Watchlist 管理)
# ==========================================
with tab_theme:
    st.subheader("🎯 題材映射與客觀生命週期判定")
    
    col_in, col_btn = st.columns([4, 1])
    user_theme = col_in.text_input("輸入題材、產業關鍵字或個股需求", placeholder="例如：矽光子 CPO、低軌衛星、機器人軸承", key="theme_input_query")
    
    if col_btn.button("🔍 探索題材標的", use_container_width=True) and user_theme:
        api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
        target_model = st.secrets.get("GEMINI_MODEL", "gemini-2.5-flash")
        if not api_key:
            st.error("⚠️ 請先在 secrets 中設定 `GEMINI_API_KEY`。")
        else:
            client = genai.Client(api_key=api_key)
            with st.spinner("AI 正在解析產業鏈與對應台股標的..."):
                candidates = map_theme_to_stocks(user_theme, client, target_model)
                st.session_state["theme_candidates"] = candidates
                st.session_state["active_theme_name"] = user_theme

    if st.session_state.get("theme_candidates"):
        st.markdown(f"##### 📋 題材「**{st.session_state.get('active_theme_name')}**」映射標的（請勾選欲分析個股）：")
        candidates_df = pd.DataFrame(st.session_state["theme_candidates"])
        if "選取" not in candidates_df.columns:
            candidates_df.insert(0, "選取", True)
        
        edited_df = st.data_editor(
            candidates_df,
            column_config={
                "選取": st.column_config.CheckboxColumn("分析", default=True),
                "symbol": st.column_config.TextColumn("代號", width=80),
                "name": st.column_config.TextColumn("名稱", width=100),
                "market": st.column_config.TextColumn("市場", width=70),
                "relevance": st.column_config.TextColumn("題材純度", width=140),
                "business_role": st.column_config.TextColumn("受惠主因 / 產品定位", width=320),
            },
            disabled=["symbol", "name", "market", "relevance", "business_role"],
            hide_index=True,
            use_container_width=True,
            key="candidate_editor"
        )
        
        selected_stocks = edited_df[edited_df["選取"] == True].to_dict("records")
        
        if st.button("🚀 開始量化運算與階段診斷", type="primary", use_container_width=True):
            if not selected_stocks:
                st.warning("請至少勾選一檔股票進行分析。")
            else:
                progress_bar = st.progress(0)
                results = []
                for i, stk in enumerate(selected_stocks):
                    sym = clean_sym(stk.get("symbol", ""))
                    mkt = str(stk.get("market", "TW")).upper()
                    ticker_str = f"{sym}.TWO" if "TWO" in mkt or "上櫃" in mkt else f"{sym}.TW"
                    
                    df_d = yf.Ticker(ticker_str).history(period="14mo")
                    if df_d.empty:
                        alt_str = f"{sym}.TW" if "TWO" in ticker_str else f"{sym}.TWO"
                        df_d = yf.Ticker(alt_str).history(period="14mo")
                    
                    if not df_d.empty and len(df_d) >= 60:
                        df_w = df_d.resample('W-FRI').agg({
                            'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'
                        }).dropna()
                        
                        feat = calculate_technical_features(df_d, df_w)
                        dow = evaluate_dow_structure(df_d)
                        eval_res = evaluate_state_machine(feat, dow, base_count=1)
                        stage_name = eval_res["stage"]
                        strat = STRATEGY_TEMPLATES.get(stage_name, STRATEGY_TEMPLATES["打底期 (Stage 1 Basing / Idle)"])
                        
                        results.append({"stock": stk, "feat": feat, "eval": eval_res, "strategy": strat})
                    progress_bar.progress((i + 1) / len(selected_stocks))
                
                st.session_state["theme_results"] = results

        if st.session_state.get("theme_results"):
            st.divider()
            st.subheader("📊 量化階段診斷與分批執行建議")
            for r in st.session_state["theme_results"]:
                stk, feat, ev, strat = r["stock"], r["feat"], r["eval"], r["strategy"]
                cur_p = feat["cur_p"]
                sym_clean = clean_sym(stk["symbol"])
                stock_title = f"📌 {stk['name']} ({sym_clean}) ｜ 生命週期：【{ev['stage']}】 ｜ {ev['substate']}"
                
                with st.expander(stock_title, expanded=True):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("當前市價", f"${cur_p:.2f}")
                    c2.metric("20MA 乖離率", f"{feat['bias_20']*100:+.1f}%")
                    c3.metric("ATR 收縮比 (10/60)", f"{feat['atr_contraction']:.2f}", "🔥 波動收縮" if feat['atr_contraction'] < 0.65 else "正常")
                    c4.metric("相對均量比 (20MV)", f"{feat['rel_vol_20']:.2f} 倍")
                    
                    st.markdown(f"""
                    **【標準量化輸出報告】**
                    - **當前判定狀態**：`{ev['stage']}`
                    - **狀態附加屬性**：Base 計數：`Base {ev['base_count']}` ｜ 中斷態標籤：`{ev['substate']}`
                    - **關鍵量化證據**：
                      - 均線坐標：現價=${cur_p:.2f} ｜ 20MA=${feat['ma']['ma20'].iloc[-1]:.2f} ｜ 60MA=${feat['ma']['ma60'].iloc[-1]:.2f} ｜ 260MA=${feat['ma']['ma260'].iloc[-1]:.2f}
                      - 滾動位階：距 260日高點 `{feat['dist_high_260']*100:.1f}%` ｜ 距 260日低點漲幅 `{feat['rise_from_low_260']*100:.1f}%`
                      - 量能收縮：近 5 日最低量與 60MV 比值 `{feat['vdu_ratio']:.2f}`（窒息量臨界 <= 0.5）
                    - **對應策略配置**：
                      - **建議分批**：`{strat['tranches']} 批`（資金比例：`{strat['ratios']}`）
                      - **執行原則**：{strat['rule']}
                      - **初始停損防守位**：{strat['stop_loss_type']}
                    """)
                    
                    btn_c1, _ = st.columns([3, 7])
                    if btn_c1.button(f"📥 追蹤 {stk['name']} ({sym_clean})", key=f"track_btn_{sym_clean}"):
                        rec = {
                            "symbol": sym_clean,
                            "name": stk["name"],
                            "theme": st.session_state.get("active_theme_name", "自訂"),
                            "created_date": get_tw_now_str("%Y-%m-%d"),
                            "stage": ev["stage"],
                            "prev_stage": "",
                            "substate": ev["substate"],
                            "base_count": ev["base_count"],
                            "pivot_price": round(feat["high_60"], 2),
                            "strategy_tranches": json.dumps(strat, ensure_ascii=False),
                            "transition_date": get_tw_now_str("%Y-%m-%d"),
                            "is_active": True
                        }
                        if save_to_watchlist(rec):
                            st.success(f"已成功將 {stk['name']} 加入 Google Sheets Watchlist 持續追蹤！")
                            st.rerun()

    st.divider()
    st.subheader("👁️ 持久化題材追蹤名單 (Watchlist)")
    watchlist = load_watchlist()
    if not watchlist:
        st.info("目前無追蹤中的觀察標的。請由上方探索題材並點擊「📥 追蹤」加入。")
    else:
        df_watch = pd.DataFrame(watchlist)
        for idx, row in df_watch.iterrows():
            sym_w, name_w = str(row.get("symbol", "")), str(row.get("name", ""))
            wc1, wc2, wc3, wc4, wc5 = st.columns([2, 2, 2, 3, 1])
            wc1.write(f"**{name_w} ({sym_w})**")
            wc2.write(f"題材：`{row.get('theme', '-')}`")
            wc3.write(f"階段：`{row.get('stage', '-')}`")
            wc4.write(f"追蹤起始：{row.get('created_date', '-')}")
            if wc5.button("🗑️ 剔除", key=f"del_wl_{sym_w}_{idx}"):
                if remove_from_watchlist(sym_w):
                    st.success(f"已剔除 {name_w} ({sym_w})")
                    st.rerun()

latest_breadth_dict = {}

with tab_market_breadth:
    st.subheader("📊 大盤指標")

    with st.expander("📖 說明：順勢操作模式切換指標與判斷準則", expanded=False):
        st.markdown("""
        本分頁依據台股月週期（20日）與季週期（60日）量化市場動能環境，作為切換**「主升段波段進攻」**與**「震盪弱勢防守」**的客觀依據：

        #### 🎯 核心指標定義與公式
        1. **月均線廣度 ($MAB_{20}$)**：`全市場收盤 > 20MA 的個股比例 (%)`。代表短線多頭土壤是否肥沃。
        2. **季均線廣度 ($MAB_{60}$)**：`全市場收盤 > 60MA 的個股比例 (%)`。代表中期多頭趨勢的穩定度。
        3. **60日領頭羊動能 (Top 10% RS)**：全市場過去 60 日相對大盤漲幅前 10% 強勢股的等權動能指數與其 20MA。站穩 20MA 代表領頭羊持續攻堅；若跌破 20MA 則意味著最強的部隊已被主力提款，波段行情面臨終結。
        4. **60日淨創新高指數 ($NNH_{60}$)**：`創 60 日新高家數 - 創 60 日新低家數`。衡量全市場領先股與破底股的多空差額。
        5. **20日突破延續度 ($BFTR_{20}$)**：`近 20 日放量突破型態中，突破後第 2 日仍守穩頸線的比例 (%)`。量化隔日沖與假突破（Squat）的殺傷力。
        6. **20日出貨日計數 ($DDC_{20}$)**：`近 20 日指數單日跌幅 ≥ 0.2% 且成交量大於前一日的次數`。評估機構大戶是否在逢高派發。

        ---

        #### 🚦 三階段模式切換決策矩陣

        | 監控指標 | 🟢 綠燈：主升段模式 (進攻) | 🟡 黃燈：震盪整理模式 (短打) | 🔴 紅燈：弱勢出貨模式 (防守) |
        | :--- | :--- | :--- | :--- |
        | **均線廣度 ($MAB_{20}$)** | **$> 60\%$** 且向上發散 | **$40\% \sim 60\%$** 區間震盪 | **$< 40\%$** 且持續下滑 |
        | **季均廣度 ($MAB_{60}$)** | **$> 50\%$** | **$35\% \sim 50\%$** | **$< 35\%$** |
        | **領頭羊動能 (Top 10% RS)**| **站穩 20MA** 且持續創高 | 於 20MA 附近來回洗盤 | **帶量摜破 20MA** (領先族群補跌) |
        | **新高差額 ($NNH_{60}$)** | 穩定為正（**$> +30$ 家**） | 接近 0 軸（**$-20 \sim +20$ 家**） | 明顯翻負（**$< -30$ 家**） |
        | **突破延續度 ($BFTR_{20}$)**| **$\\ge 60\%$**（突破推進順暢）| **$40\% \sim 60\%$**（頻繁橫盤洗盤）| **$< 40\%$**（突破即長黑誘多） |
        | **出貨日計數 ($DDC_{20}$)**| **$\\le 2$ 次**（無密集拋售） | **$3 \sim 4$ 次**（主力派發警戒） | **$\\ge 5$ 次**（機構集中出貨） |
        | **建議總持股曝險** | **80% ~ 100%** | **30% ~ 40%** | **0% ~ 10%（保留現金）** |
        | **停利戰術設定** | 追求波段：未實現達 **20%+** 或月線正乖離 **30%** 移動停利 | 短線兌現：達 **6%~9%** 先出半數，持股 **3天不動時間停損** | 停止追買，嚴格執行損益兩平保本或停損砍倉 |
        """)

    b_col1, b_col2 = st.columns(2)
    mkt_view = b_col1.selectbox("市場選擇", ["上市 (TWSE)", "上櫃 (TPEX)"], index=0)
    show_days = {"近 20 個交易日": 20, "近 60 個交易日": 60, "近 120 個交易日": 120}[b_col2.selectbox("時間跨度", ["近 20 個交易日", "近 60 個交易日", "近 120 個交易日"], index=1)]

    with st.spinner("正在計算大盤指標..."):
        breadth_df = compute_market_breadth_data(market_rankings, "TW" if "上市" in mkt_view else "TWO")

    if breadth_df is None or breadth_df.empty:
        st.warning("⚠️ 暫時無法取得大盤寬度資料。")
    else:
        plot_df = breadth_df.tail(show_days)
        latest, prev = plot_df.iloc[-1], plot_df.iloc[-2] if len(plot_df) >= 2 else plot_df.iloc[-1]
        latest_breadth_dict = latest.to_dict()

        is_green = (latest["above_20ma"] >= 60 and latest["bftr_20"] >= 60 and latest["ddc_20"] <= 2 and latest["leader_dist_ma20"] >= 0)
        is_red = (latest["above_20ma"] < 40 or latest["ddc_20"] >= 5 or latest["bftr_20"] < 40 or latest["leader_dist_ma20"] < -3.0)

        if is_green:
            st.success("🟢 **【綠燈：主升段模式】全力進攻** ｜ 獲利目標：未實現 20%+ / 月線乖離 30% ｜ 領頭羊強勁守穩20MA ｜ 建議總持股：80% ~ 100%")
        elif is_red:
            st.error("🔴 **【紅燈：弱勢出貨模式】強制防守** ｜ 全面停止追突破，嚴守保本與停損 ｜ 領頭羊轉弱或出貨密集 ｜ 建議總持股：0% ~ 10%（現金為王）")
        else:
            st.warning("🟡 **【黃燈：震盪整理模式】短打防守** ｜ 獲利目標：6% ~ 9% 先出半數 ｜ 3 天不動時間停損 ｜ 建議總持股：30% ~ 40%")

        st.markdown("##### 📌 當日即時總覽")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("月均線廣度 (MAB20)", f"{latest['above_20ma']:.1f}%", f"{latest['above_20ma'] - prev['above_20ma']:+.1f}%")
        k2.metric("季均線廣度 (MAB60)", f"{latest['above_60ma']:.1f}%", f"{latest['above_60ma'] - prev['above_60ma']:+.1f}%")
        k3.metric("領頭羊動能 (Top10%)", f"{latest['leader_index']:.1f}", f"乖離20MA: {latest['leader_dist_ma20']:+.1f}%")
        k4.metric("60日淨創新高 (NNH60)", f"{int(latest['net_high_low_60']):+d} 家", f"新高 {int(latest['nh_60'])} / 新低 {int(latest['nl_60'])}")
        k5.metric("20日突破延續度 (BFTR)", f"{latest['bftr_20']:.1f}%", "🔥 突破順暢" if latest['bftr_20']>=60 else ("❄️ 假突破多" if latest['bftr_20']<40 else "⚡ 橫盤洗盤"))
        k6.metric("20日出貨日計數 (DDC)", f"{int(latest['ddc_20'])} 次", "🔴 高風險" if latest['ddc_20']>=5 else ("🟡 警戒中" if latest['ddc_20']>=3 else "🟢 安全"))

        st.divider()
        cfg = {"scrollZoom": False, "displayModeBar": False}
        layout = dict(
            hovermode="x unified",
            margin=dict(l=45, r=30, t=55, b=35),
            dragmode=False,
            legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="left", x=0, font=dict(size=11))
        )

        # 1. 均線覆蓋率
        st.markdown("#### 1. 均線廣度覆蓋率 (%) ｜ MAB20 & MAB60")
        fig1 = go.Figure([go.Scatter(x=plot_df.index, y=plot_df[c], mode="lines", name=n, line=dict(color=col, width=2)) for c, n, col in [("above_20ma", "站上 20MA (月線廣度 MAB20)", "#FF5722"), ("above_60ma", "站上 60MA (季線廣度 MAB60)", "#2196F3"), ("above_240ma", "站上 240MA (年線)", "#4CAF50")]])
        fig1.add_hline(y=60, line_dash="dot", line_color="green", annotation_text="60% 強勢擴張", annotation_position="top left", annotation_font=dict(size=10, color="green"))
        fig1.add_hline(y=50, line_dash="dash", line_color="gray", annotation_text="50% 多空分水嶺", annotation_position="bottom right", annotation_font=dict(size=10, color="gray"))
        fig1.add_hline(y=40, line_dash="dot", line_color="red", annotation_text="40% 防守警戒", annotation_position="bottom left", annotation_font=dict(size=10, color="red"))
        fig1.update_layout(layout, yaxis=dict(title="比例 (%)", range=[0, 100], fixedrange=True), xaxis=dict(fixedrange=True))
        st.plotly_chart(fig1, use_container_width=True, config=cfg)

        # 2. 60日創新高/新低指標與淨差
        st.markdown("#### 2. 60日淨創新高指標 (NNH60)")
        fig2 = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.12, subplot_titles=("60日創新高 / 創新低比例 (%)", "60日淨創新高家數差 (NNH60)"))
        fig2.add_trace(go.Scatter(x=plot_df.index, y=plot_df["nh_60_ratio"], mode="lines", name="60日新高比例 (%)", line=dict(color="#E91E63", width=2)), row=1, col=1)
        fig2.add_trace(go.Scatter(x=plot_df.index, y=plot_df["nl_60_ratio"], mode="lines", name="60日新低比例 (%)", line=dict(color="#00BCD4", width=2)), row=1, col=1)
        fig2.add_trace(go.Bar(x=plot_df.index, y=plot_df["net_high_low_60"], name="60日新高新低差 (家數)", marker_color=["#4CAF50" if v >= 0 else "#F44336" for v in plot_df["net_high_low_60"]]), row=2, col=1)
        fig2.add_hline(y=30, line_dash="dot", line_color="green", annotation_text="+30 家強勢門檻", annotation_position="top left", annotation_font=dict(size=10, color="green"), row=2, col=1)
        fig2.add_hline(y=0, line_dash="dash", line_color="gray", row=2, col=1)
        fig2.add_hline(y=-30, line_dash="dot", line_color="red", annotation_text="-30 家弱勢門檻", annotation_position="bottom left", annotation_font=dict(size=10, color="red"), row=2, col=1)
        fig2.update_layout(layout, height=520)
        fig2.update_xaxes(fixedrange=True); fig2.update_yaxes(fixedrange=True)
        st.plotly_chart(fig2, use_container_width=True, config=cfg)

        # 3. 突破延續度與出貨日計數
        st.markdown("#### 3. 動能品質與出貨監控 ｜ BFTR20 ＆ DDC20")
        fig3 = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.12, subplot_titles=(f"20 日突破延續度 (BFTR20) ｜ 最新: {latest['bftr_20']:.1f}%", f"20 日出貨日計數器 (DDC20) ｜ 最新: {int(latest['ddc_20'])} 次"))
        fig3.add_trace(go.Scatter(x=plot_df.index, y=plot_df["bftr_20"], mode="lines+markers", name="突破延續度 (%)", line=dict(color="#673AB7", width=2)), row=1, col=1)
        fig3.add_hline(y=60, line_dash="dot", line_color="green", annotation_text="60% 真突破順暢", annotation_position="top left", annotation_font=dict(size=10, color="green"), row=1, col=1)
        fig3.add_hline(y=40, line_dash="dot", line_color="red", annotation_text="40% 假突破高危", annotation_position="bottom left", annotation_font=dict(size=10, color="red"), row=1, col=1)
        fig3.add_trace(go.Bar(x=plot_df.index, y=plot_df["ddc_20"], name="20日出貨天數", marker_color=["#D32F2F" if v >= 5 else ("#FF9800" if v >= 3 else "#388E3C") for v in plot_df["ddc_20"]]), row=2, col=1)
        fig3.add_hline(y=5, line_dash="dash", line_color="red", annotation_text="5 次強制防守", annotation_position="top left", annotation_font=dict(size=10, color="red"), row=2, col=1)
        fig3.add_hline(y=3, line_dash="dot", line_color="orange", annotation_text="3 次減速警戒", annotation_position="bottom left", annotation_font=dict(size=10, color="orange"), row=2, col=1)
        fig3.update_layout(layout, height=540)
        fig3.update_xaxes(fixedrange=True); fig3.update_yaxes(fixedrange=True)
        st.plotly_chart(fig3, use_container_width=True, config=cfg)

        # 4. 60日領頭羊動能 (Top 10% RS Momentum)
        st.markdown(f"#### 4. 60日相對強度領頭羊動能 (Top 10% RS) ｜ 最新指數: **{latest['leader_index']:.1f}** (乖離 20MA: **{latest['leader_dist_ma20']:+.1f}%**)")
        fig4 = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.12, subplot_titles=("領頭羊動能指數 (Top 10% RS) 與 20MA 軌跡", "領頭羊指數與 20MA 乖離距離 (%)"))
        fig4.add_trace(go.Scatter(x=plot_df.index, y=plot_df["leader_index"], mode="lines", name="領頭羊動能指數", line=dict(color="#FF9800", width=2.2)), row=1, col=1)
        fig4.add_trace(go.Scatter(x=plot_df.index, y=plot_df["leader_ma20"], mode="lines", name="領頭羊 20MA", line=dict(color="#1976D2", width=1.8, dash="dash")), row=1, col=1)
        fig4.add_trace(go.Bar(x=plot_df.index, y=plot_df["leader_dist_ma20"], name="領頭羊乖離 20MA (%)", marker_color=["#4CAF50" if v >= 0 else "#F44336" for v in plot_df["leader_dist_ma20"]]), row=2, col=1)
        fig4.add_hline(y=0, line_dash="solid", line_color="black", row=2, col=1)
        fig4.add_hline(y=5, line_dash="dot", line_color="green", annotation_text="+5% 動能噴出", annotation_position="top left", annotation_font=dict(size=10, color="green"), row=2, col=1)
        fig4.add_hline(y=-3, line_dash="dot", line_color="red", annotation_text="-3% 領頭羊轉弱", annotation_position="bottom left", annotation_font=dict(size=10, color="red"), row=2, col=1)
        fig4.update_layout(layout, height=540)
        fig4.update_xaxes(fixedrange=True); fig4.update_yaxes(fixedrange=True)
        st.plotly_chart(fig4, use_container_width=True, config=cfg)

        # 5. 均線多頭排列
        st.markdown("#### 5. 均線多頭排列比例 (%)")
        fig5 = go.Figure([
            go.Scatter(x=plot_df.index, y=plot_df["short_bull_ratio"], mode="lines", name="短均多頭排列", line=dict(color="#9C27B0", width=2)),
            go.Scatter(x=plot_df.index, y=plot_df["long_bull_ratio"], mode="lines", name="長均多頭排列", line=dict(color="#3F51B5", width=2))
        ])
        fig5.add_hline(y=50, line_dash="dash", line_color="gray", annotation_text="50% 多空分水嶺", annotation_position="top left", annotation_font=dict(size=10, color="gray"))
        fig5.update_layout(layout, yaxis=dict(title="多頭排列比例 (%)", range=[0, 100], fixedrange=True), xaxis=dict(fixedrange=True))
        st.plotly_chart(fig5, use_container_width=True, config=cfg)

        # 6. 季線距離
        st.markdown(f"#### 6. 大盤現價與 60日 MA 距離 (%) ｜ 最新：**{latest['dist_60ma_pct']:+.2f}%**")
        fig6 = go.Figure([
            go.Bar(x=plot_df.index, y=plot_df["dist_60ma_pct"], name="季線乖離距離 (%)", marker_color=["#F44336" if v >= 0 else "#4CAF50" for v in plot_df["dist_60ma_pct"]]),
            go.Scatter(x=plot_df.index, y=plot_df["dist_60ma_pct"], mode="lines+markers", name="趨勢軌跡", line=dict(color="#1976D2", width=1.5), marker=dict(size=4))
        ])
        fig6.add_hline(y=0, line_dash="solid", line_color="black")
        fig6.add_hline(y=10, line_dash="dash", line_color="#E91E63", annotation_text="+10% 正向過熱區", annotation_position="top left", annotation_font=dict(size=10, color="#E91E63"))
        fig6.add_hline(y=-10, line_dash="dash", line_color="#00BCD4", annotation_text="-10% 負向超跌區", annotation_position="bottom left", annotation_font=dict(size=10, color="#00BCD4"))
        fig6.update_layout(layout, yaxis=dict(title="距離 (%)", fixedrange=True), xaxis=dict(fixedrange=True))
        st.plotly_chart(fig6, use_container_width=True, config=cfg)

with tab_ai:
    st.subheader("🤖 Gemini 順勢交易智能助理")
    st.caption("具備全域數據感知與工具調用能力：能直接分析持股風控狀態、全市場排行榜，並支援自動掃描題材寫入 Watchlist。")

    top_leaders = sorted(market_rankings, key=lambda x: x.get("rs_rating", 0), reverse=True)[:15] if market_rankings else []
    leaders_summary = [{"代號": x.get("symbol"), "名稱": x.get("name"), "RS評分": x.get("rs_rating"), "RS_ratio": x.get("rs_ratio"), "RS_mom": x.get("rs_momentum")} for x in top_leaders]

    system_context = f"""
    你是一位專精於台股動能順勢交易（Trend Following / CANSLIM / 領導股操作）的頂級操盤手顧問。
    你目前整合在用戶的即時監控儀表板中，以下是即時的系統記憶體數據：

    【用戶目前持股與損益風控狀態】
    {json.dumps(portfolio_live_summary, ensure_ascii=False, indent=2) if portfolio_live_summary else "目前無持倉"}

    【最新大盤寬度與市場健康度指標】
    {json.dumps(latest_breadth_dict, ensure_ascii=False, indent=2) if latest_breadth_dict else "大盤寬度未載入"}

    【全市場 RS 評分前 15 大強勢領導股】
    {json.dumps(leaders_summary, ensure_ascii=False, indent=2) if leaders_summary else "排行榜未載入"}

    【核心任務與工具調用準則】
    1. 當使用者要求「掃描題材」、「尋找/篩選概念股」或「加入 Watchlist」時，請務必直接調用 `run_auto_scan_pipeline` 工具執行自動化掃描。
    2. 請根據上述數據與嚴謹的右側順勢交易邏輯（關注 RS 強度、回檔風控、停損守紀律、汰弱留強），為用戶提供具體、冷靜且有條理的分析建議。
    """

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_prompt := st.chat_input("請輸入你的問題（例如：幫我掃描 CPO 題材，挑出目前處於『初升段突破』或『VCP 醞釀期』的標的，直接加入 Watchlist。）"):
        st.session_state.chat_history.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))
        target_model = st.secrets.get("GEMINI_MODEL", "gemini-2.5-flash")
        
        if not api_key:
            err_msg = "⚠️ 請在 secrets 中設定 `GEMINI_API_KEY`。"
            with st.chat_message("assistant"):
                st.error(err_msg)
            st.session_state.chat_history.append({"role": "assistant", "content": err_msg})
        else:
            try:
                client = genai.Client(api_key=api_key)
                with st.chat_message("assistant"):
                    with st.spinner("思考與執行工具中..."):
                        response = client.models.generate_content(
                            model=target_model,
                            contents=user_prompt,
                            config={
                                "system_instruction": system_context,
                                "tools": [run_auto_scan_pipeline]
                            }
                        )

                        # 檢驗是否觸發 Tool Calling
                        function_calls = getattr(response, "function_calls", None)
                        if not function_calls and hasattr(response, "candidates") and response.candidates:
                            cand = response.candidates[0]
                            if hasattr(cand, "content") and hasattr(cand.content, "parts"):
                                function_calls = [p.function_call for p in cand.content.parts if getattr(p, "function_call", None)]

                        if function_calls:
                            tool_outputs = []
                            for fc in function_calls:
                                f_name = getattr(fc, "name", "")
                                f_args = getattr(fc, "args", {}) or {}
                                if f_name == "run_auto_scan_pipeline":
                                    tool_res = run_auto_scan_pipeline(**f_args)
                                    tool_outputs.append(tool_res)
                            reply = "\n\n".join(tool_outputs) if tool_outputs else "工具已執行完畢。"
                        else:
                            reply = response.text or ""

                        st.markdown(reply)
                st.session_state.chat_history.append({"role": "assistant", "content": reply})
            except Exception as e:
                with st.chat_message("assistant"):
                    st.error(f"Gemini API 呼叫失敗: {e}")
