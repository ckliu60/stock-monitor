#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (v2.0)
---------------------------------------------------------------------------
每週二、四、六 台灣時間 06:30，抓取美股與台股真實收盤價，計算跨期漲跌、
標示階梯式警示，組成精簡三段式純文字訊息，推播到使用者的 LINE。

資料源：Yahoo Finance（yfinance）。只用真實收盤價，不做任何模擬／推算。
清單來源：環境變數 SHEET_CSV_URL（Google 試算表發布的 CSV）；未設定時讀本機 watchlist.csv。
"""

import os
import sys
import io
import csv
import time
import datetime as dt
from pathlib import Path

import requests
import pandas as pd
import yfinance as yf

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# ---------------------------------------------------------------------------
# 設定區（固定項）
# ---------------------------------------------------------------------------

INDICES = [
    ("^DJI", "道瓊"),
    ("^IXIC", "那斯達克"),
    ("^GSPC", "S&P500"),
    ("^SOX", "費半"),
]

PERIODS = [("週", 5), ("月", 21), ("年", 252)]
MA_WINDOWS = [("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]
TW_INDEX = ("^TWII", "台股大盤")

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
LINE_MAX_LEN = 4800

HERE = Path(__file__).resolve().parent
LOCAL_WATCHLIST = HERE / "watchlist.csv"
OUTPUT_DIR = os.environ.get("REPORT_OUTPUT_DIR", "").strip() or str(HERE)
STATS_SHEET_URL = os.environ.get("STATS_SHEET_URL", "").strip()

# ---------------------------------------------------------------------------
# 清單載入
# ---------------------------------------------------------------------------

def _normalize_row(row):
    market = (row.get("market") or "").strip()
    code = (row.get("code") or "").strip()
    yf_symbol = (row.get("yf_symbol") or "").strip()
    name = (row.get("name") or "").strip()
    group = (row.get("group") or "").strip() or "觀察"
    if not code and not yf_symbol: return None
    if not yf_symbol:
        if market == "台股" or code.replace(".", "").isdigit():
            yf_symbol = code if code.endswith(".TW") else f"{code}.TW"
        else:
            yf_symbol = code
    if not name: name = code or yf_symbol
    return {"market": market, "code": code or yf_symbol,
            "yf_symbol": yf_symbol, "name": name, "group": group}

def load_watchlist():
    url = os.environ.get("SHEET_CSV_URL", "").strip()
    text = None
    if url:
        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            text = resp.text
        except Exception as e:
            print(f"[警告] 讀取試算表 CSV 失敗，改用本機清單：{e}", file=sys.stderr)
    if text is None:
        if not LOCAL_WATCHLIST.exists():
            raise FileNotFoundError(f"找不到清單：{LOCAL_WATCHLIST}，且未設定 SHEET_CSV_URL")
        text = LOCAL_WATCHLIST.read_text(encoding="utf-8")

    rows = []
    for raw in csv.DictReader(io.StringIO(text)):
        norm = _normalize_row({k.strip(): (v or "") for k, v in raw.items()})
        if norm: rows.append(norm)
    return rows

# ---------------------------------------------------------------------------
# 數據擷取與計算
# ---------------------------------------------------------------------------

def _extract_close(data, symbol, n_symbols):
    try: close = data["Close"]
    except Exception: return pd.Series(dtype="float64")
    if isinstance(close, pd.DataFrame):
        if symbol in close.columns: s = close[symbol]
        elif n_symbols == 1 and close.shape[1] == 1: s = close.iloc[:, 0]
        else: return pd.Series(dtype="float64")
    else: s = close
    return s.dropna()

def repair_splits(series, symbol="", low=0.4, high=2.5):
    if len(series) < 2: return series
    s = series.copy()
    for i in range(len(series) - 1, 0, -1):
        prev = float(series.iloc[i - 1])
        if prev == 0: continue
        ratio = float(series.iloc[i]) / prev
        if ratio <= low or ratio >= high:
            s.iloc[:i] = s.iloc[:i] * ratio
    return s

def fetch_data(symbols):
    data = yf.download(symbols, period="2y", interval="1d", auto_adjust=False, progress=False, threads=True, group_by="column")
    closes = {}
    for sym in symbols:
        s = _extract_close(data, sym, len(symbols))
        if len(s) >= 2: closes[sym] = repair_splits(s, symbol=sym)
        else:
            try:
                d1 = yf.download(sym, period="2y", interval="1d", auto_adjust=False, progress=False)
                s = _extract_close(d1, sym, 1)
                if len(s) >= 2: closes[sym] = repair_splits(s, symbol=sym)
            except: pass
    return closes

def pct(new, old): return None if not old else (new / old - 1.0) * 100.0

def compute_period_detail(series):
    last = float(series.iloc[-1])
    out = {}
    for name, days in PERIODS:
        if len(series) >= days + 1:
            ref = float(series.iloc[-(days + 1)])
            out[name] = {"ret": pct(last, ref), "price": ref, "date": series.index[-(days + 1)].date()}
        else:
            out[name] = {"ret": None, "price": None, "date": None}
    return out

def daily_change(series): return pct(float(series.iloc[-1]), float(series.iloc[-2])) if len(series) >= 2 else None
def compute_ma(series): return {name: round(float(series.rolling(w).mean().iloc[-1]), 2) if len(series) >= w else None for name, w in MA_WINDOWS}
def weekly_ma(series, weeks):
    wk = series.resample("W").last().dropna()
    return round(float(wk.rolling(weeks).mean().iloc[-1]), 2) if len(wk) >= weeks else None

def price_level(last, ma):
    if not last or not ma: return ""
    m10, m20, m60, m240 = ma.get("MA10"), ma.get("MA20"), ma.get("MA60"), ma.get("MA240")
    if m10 and last >= m10: return "雙週線上"
    if m20 and last >= m20: return "雙週~月"
    if m60 and last >= m60: return "月~季"
    if m240 and last >= m240: return "季~年"
    if m240: return "年線下"
    return ""

def ma20w_level(last, ma20w): return "" if not last or not ma20w else ("線上" if last >= ma20w else "破線")

def compute_kd(df, n=9):
    if df is None or len(df) < n: return (None, None)
    low = df["Low"].rolling(n).min()
    high = df["High"].rolling(n).max()
    rng = (high - low)
    k = d = 50.0
    for i in range(len(df)):
        denom = rng.iloc[i]
        rsv = 50.0 if pd.isna(denom) or denom == 0 else float((df["Close"].iloc[i] - low.iloc[i]) / denom * 100)
        k = 2 / 3 * k + 1 / 3 * rsv
        d = 2 / 3 * d + 1 / 3 * k
    return (round(k, 1), round(d, 1))

def kd_timeframes(df):
    agg = {"High": "max", "Low": "min", "Close": "last"}
    out = {"日": compute_kd(df)}
    try: out["週"] = compute_kd(df.resample("W").agg(agg).dropna())
    except: out["週"] = (None, None)
    for rule in ("ME", "M"):
        try:
            out["月"] = compute_kd(df.resample(rule).agg(agg).dropna())
            break
        except: out["月"] = (None, None)
    return out

# ---------------------------------------------------------------------------
# 訊息組裝 (v2.0)
# ---------------------------------------------------------------------------

def fmt_pct(v, decimals=1): return "—" if v is None else f"{'▲' if v>=0 else '▼'}{abs(v):.{decimals}f}%"
def fmt_val(v): return "—" if v is None else f"{'▲' if v>=0 else '▼'}{abs(v):.1f}"

def build_message(index_rows, stock_rows, extras, vix_us, vix_tw):
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    lines = [f"📊 美股／台股監控　{today.month}/{today.day:02d}（{wd}）"]

    # 1. 大盤現況
    lines.extend(["", "━━ 大盤 ━━"])
    for idx in index_rows:
        last, dchg = idx["last"], idx["dchg"]
        lines.append(f"{idx['name']}　資料暫缺" if last is None else f"{idx['name']}　{last:,.0f}　{fmt_pct(dchg)}")
        
    twii = extras.get("twii")
    if twii and twii["last"] is not None:
        lines.append(f"{twii['name']}　{twii['last']:,.0f}　{fmt_pct(twii['dchg'])}")

    # 2. 塊 2：位階警示 (v2.0)
    lines.extend(["", "━━ 位階警示 (v2.0) ━━"])
    kd_map = {item['code']: item['tf'] for item in extras.get('kd', [])}
    alerts_found = False

    def get_alerts(item, category="個股"):
        is_tw = (item.get("market") == "台股" or item.get("yf") == "^TWII")
        is_idx = (category == "指數")
        last = item.get("last")
        ma = item.get("ma") or {}
        ma20w = item.get("ma20w")
        rets = item.get("rets") or {}
        kd = kd_map.get(item.get("yf")) or {}
        
        if last is None: return None
        
        # [跌警示] 條件收集
        drop = []
        if ma.get("MA60") and last < ma.get("MA60"): drop.append("破季線")
        if ma.get("MA240") and last < ma.get("MA240"): drop.append("破年線")
        if is_tw and ma20w and last < ma20w: drop.append("破20周")
        
        w_ret, m_ret = rets.get("週"), rets.get("月")
        if (w_ret and w_ret <= -20) or (m_ret and m_ret <= -20):
            drop.append("週或月跌>20%")
            
        vix = vix_tw if is_tw else vix_us
        if is_idx and vix > 30: drop.append("VIX>30")
        elif not is_idx and is_tw and vix > 35: drop.append("VIX>35")
        
        ku20, ko80 = [], []
        for tf in ["日", "週", "月"]:
            if kd.get(tf) and kd[tf][0] is not None:
                if kd[tf][0] < 20: ku20.append(tf)
                if kd[tf][0] > 80: ko80.append(tf)
                
        if is_tw and ku20: drop.append(f"{'/'.join(ku20)}K<20")
        
        res = ""
        if drop: res += f"🟢[跌] {', '.join(drop)}  "
        if is_tw and ko80: res += f"🔴[過熱] {'/'.join(ko80)}K>80  "
        if is_tw and vix > 40: res += f"🔥[落底] VIX>40"
        
        return f"{item['name']}：{res.strip()}" if res else None

    if twii:
        twii_mock = twii.copy()
        twii_mock["market"] = "台股"
        al = get_alerts(twii_mock, "指數")
        if al: lines.append(al); alerts_found = True
        
    for idx in index_rows:
        al = get_alerts(idx, "指數")
        if al: lines.append(al); alerts_found = True
            
    for st in stock_rows:
        al = get_alerts(st, "個股")
        if al: lines.append(al); alerts_found = True
            
    if not alerts_found: lines.append("今日無觸發位階警示")

    # 3. 個股數據
    lines.extend(["", "━━ 個股　週 / 月 / 年 ━━"])
    for m_lbl in ("美股", "台股"):
        subset = [s for s in stock_rows if s["market"] == m_lbl]
        if not subset: continue
        lines.append(f"〔{m_lbl}〕")
        for st in subset:
            r = st["rets"]
            trio = f"{fmt_val(r.get('週'))} / {fmt_val(r.get('月'))} / {fmt_val(r.get('年'))}"
            lines.append(f"{st['code']}　{trio}")

    text = "\n".join(lines).rstrip()
    return text[:LINE_MAX_LEN] + "\n…（訊息過長已截斷）" if len(text) > LINE_MAX_LEN else text

# ---------------------------------------------------------------------------
# 主流程模組
# ---------------------------------------------------------------------------

def gather_data():
    watch = load_watchlist()
    index_syms = [s for s, _ in INDICES]
    stock_syms = [w["yf_symbol"] for w in watch]
    all_syms = list(set(index_syms + stock_syms + ["^TWII"]))
    
    closes = fetch_data(all_syms)

    # 抓 VIX
    try:
        v_data = yf.download(["^VIX", "^VIXTWN"], period="5d", progress=False)
        c = v_data["Close"] if "Close" in v_data else v_data
        us_vix = float(c["^VIX"].dropna().iloc[-1])
        try: tw_vix = float(c["^VIXTWN"].dropna().iloc[-1])
        except: tw_vix = us_vix
    except:
        us_vix, tw_vix = 0.0, 0.0

    index_rows = []
    for sym, name in INDICES:
        s = closes.get(sym)
        if s is None:
            index_rows.append({"name": name, "yf": sym, "last": None, "dchg": None, "asof": None, "detail": {}, "rets": {}, "ma": {}})
        else:
            detail = compute_period_detail(s)
            index_rows.append({"name": name, "yf": sym, "last": float(s.iloc[-1]), "dchg": daily_change(s), "asof": s.index[-1].date(),
                               "detail": detail, "rets": {k: detail[k]["ret"] for k in detail}, "ma": compute_ma(s)})

    stock_rows = []
    for w in watch:
        s = closes.get(w["yf_symbol"])
        if s is None: continue
        is_tw = (w["market"] == "台股" or w["yf_symbol"].endswith(".TW"))
        detail = compute_period_detail(s)
        stock_rows.append({
            "name": w["name"], "code": w["code"], "yf": w["yf_symbol"], "group": w["group"], "market": w["market"] or ("台股" if is_tw else "美股"),
            "last": float(s.iloc[-1]), "asof": s.index[-1].date(), "dchg": daily_change(s), "detail": detail,
            "rets": {k: detail[k]["ret"] for k in detail}, "ma": compute_ma(s), "ma20w": weekly_ma(s, 20) if is_tw else None,
        })
        
    # v2.0 KD 全抓
    tw_targets = [(w["yf_symbol"], w["name"]) for w in watch if w["market"] == "台股" or w["yf_symbol"].endswith(".TW")]
    kd_syms = list(set([sym for sym, _ in tw_targets] + ["^TWII"]))
    kd_names = {w["yf_symbol"]: w["name"] for w in watch}
    kd_names["^TWII"] = "台股大盤"
    
    kd_list = []
    if kd_syms:
        try:
            raw_kd = yf.download(kd_syms, period="2y", interval="1d", auto_adjust=False, progress=False, group_by="ticker")
            for sym in kd_syms:
                df = raw_kd if len(kd_syms) == 1 else raw_kd[sym]
                if "Close" in df:
                    df = df[["High", "Low", "Close"]].dropna()
                    kd_list.append({
                        "code": sym, "name": kd_names.get(sym, sym),
                        "asof": df.index[-1].date() if len(df) else None,
                        "tf": kd_timeframes(df)
                    })
        except: pass

    twii_s = closes.get("^TWII")
    extras_twii = None
    if twii_s is not None:
        detail = compute_period_detail(twii_s)
        extras_twii = {
            "yf": "^TWII", "name": "台股大盤", "last": float(twii_s.iloc[-1]),
            "asof": twii_s.index[-1].date(), "dchg": daily_change(twii_s),
            "ma20w": weekly_ma(twii_s, 20), "detail": detail,
            "rets": {k: detail[k]["ret"] for k in detail}, "ma": compute_ma(twii_s)
        }
        
    extras = {"twii": extras_twii, "kd": kd_list}
    return index_rows, stock_rows, extras, us_vix, tw_vix

def push_line(text, force_dry_run=False):
    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    uid = os.environ.get("LINE_USER_ID", "").strip()
    if force_dry_run or not token or not uid:
        print("[DRY-RUN] 訊息預覽如下：\n\n" + text + "\n")
        return False
    resp = requests.post(LINE_PUSH_URL, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                         json={"to": uid, "messages": [{"type": "text", "text": text}]}, timeout=20)
    resp.raise_for_status()
    print("[OK] 已推播至 LINE。")
    return True

def main():
    force_dry_run = "--dry-run" in sys.argv
    try:
        index_rows, stock_rows, extras, vix_us, vix_tw = gather_data()
        text = build_message(index_rows, stock_rows, extras, vix_us, vix_tw)
        push_line(text, force_dry_run=force_dry_run)
        return 0
    except Exception as e:
        print(f"[嚴重錯誤] {e}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
