#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (v3.3)
---------------------------------------------------------------------------
每週二、四、六 台灣時間 06:30，抓取美股與台股真實收盤價，計算跨期漲跌、
標示階梯式警示，組成精簡三段式純文字訊息，推播到使用者的 LINE。

資料源：Yahoo Finance（yfinance）。只用真實收盤價，不做任何模擬／推算。
"""

import os
import sys
import io
import csv
import time
import datetime as dt
from pathlib import Path

# 強制關閉 curl_cffi 避免 DNS 逾時，改回標準 requests
os.environ["YFINANCE_CACHE_DIR"] = ""
os.environ["CURL_CA_BUNDLE"] = ""

import requests
import pandas as pd
import yfinance as yf

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# ---------------------------------------------------------------------------
# 設定區
# ---------------------------------------------------------------------------

INDICES = [
    ("^DJI", "道瓊"),
    ("^IXIC", "那指"),
    ("^GSPC", "標普"),
    ("^SOX", "費半"),
]

# v3.2/v3.3: Only alert for these specific targets in section 4
ALERT_TARGETS = [
    "道瓊", "那指", "標普", "費半",
    "台積電ADR", "輝達", "特斯拉", "SpaceX",
    "台積電", "聯發科", "元大台灣50"
]

PERIODS = [("週", 5), ("月", 21), ("年", 252)]
MA_WINDOWS = [("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
LINE_MAX_LEN = 4800

HERE = Path(__file__).resolve().parent
LOCAL_WATCHLIST = HERE / "watchlist.csv"

# ---------------------------------------------------------------------------
# 工具函式
# ---------------------------------------------------------------------------

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
            print(f"[警告] 讀取試算表失敗，改用本機：{e}", file=sys.stderr)
    if text is None:
        if not LOCAL_WATCHLIST.exists(): return []
        text = LOCAL_WATCHLIST.read_text(encoding="utf-8")

    rows = []
    for raw in csv.DictReader(io.StringIO(text)):
        code = (raw.get("code") or "").strip()
        yf_symbol = (raw.get("yf_symbol") or "").strip()
        market = (raw.get("market") or "").strip()
        if not yf_symbol:
            yf_symbol = f"{code}.TW" if (market=="台股" or code.isdigit()) else code
        rows.append({
            "market": market, "code": code, "yf_symbol": yf_symbol,
            "name": (raw.get("name") or yf_symbol).strip()
        })
    return rows

def fetch_data(symbols, period="2y"):
    data = yf.download(symbols, period=period, interval="1d", auto_adjust=False, progress=False, group_by="column")
    closes = {}
    for sym in symbols:
        try:
            c = data["Close"]
            s = c[sym] if isinstance(c, pd.DataFrame) else c
            closes[sym] = s.dropna()
        except: pass
    return closes

def pct(new, old): return None if not old else (new / old - 1.0) * 100.0

def compute_ma(series): return {name: round(float(series.rolling(w).mean().iloc[-1]), 2) if len(series) >= w else None for name, w in MA_WINDOWS}

def weekly_ma(series, weeks):
    wk = series.resample("W").last().dropna()
    return round(float(wk.rolling(weeks).mean().iloc[-1]), 2) if len(wk) >= weeks else None

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
    return (round(k, 0), round(d, 0))

def kd_timeframes(df):
    agg = {"High": "max", "Low": "min", "Close": "last"}
    out = {"日": compute_kd(df)}
    try: out["週"] = compute_kd(df.resample("W").agg(agg).dropna())
    except: out["週"] = (None, None)
    try: out["月"] = compute_kd(df.resample("ME").agg(agg).dropna())
    except: out["月"] = (None, None)
    return out

# ---------------------------------------------------------------------------
# 排版與訊息組裝 (v3.3)
# ---------------------------------------------------------------------------

def fmt_pct(v):
    if v is None: return "—"
    return f"{'▲' if v>0 else '▼' if v<0 else ''}{abs(v):.1f}%"

def get_weather(pct_val):
    if pct_val is None: return ""
    if pct_val >= 2.0: return "☀️"
    elif pct_val > 0: return "🌤️"
    elif pct_val == 0: return "☁️"
    elif pct_val > -2.0: return "🌧️"
    else: return "⛈️"

def pad_name(name):
    target_len = 6
    curr_len = len(name)
    if curr_len < target_len:
        return name + "　" * (target_len - curr_len)
    return name

def build_message(closes, watch, kd_data, vix_hist):
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    lines = [f"📊 股市核心快報 {today.month:02d}/{today.day:02d} ({wd})\n"]

    # ---------------- 1. 大盤與核心 ----------------
    lines.append("▌一、大盤與核心狀況 (+/- 2%)")
    
    # 美股
    for i, (sym, name) in enumerate(INDICES, 1):
        s = closes.get(sym)
        if s is not None and len(s) >= 2:
            last, prev = s.iloc[-1], s.iloc[-2]
            dchg = pct(last, prev)
            lines.append(f"　{i}. {pad_name(name)} {last:,.0f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
            
    # 虛線隔開
    lines.append("　" + "-" * 32)
    
    # 台股核心
    tw_cores = [
        ("^TWII", "台股大盤"), 
        ("2330.TW", "台積電"), ("2330", "台積電"),
        ("0050.TW", "0050"), ("0050", "0050"), ("0050.TW", "元大台灣50"), ("0050", "元大台灣50")
    ]
    core_items = []
    seen_core = set()
    
    # 抓 TWII
    s_twii = closes.get("^TWII")
    item_idx = len(INDICES) + 1
    if s_twii is not None and len(s_twii) >= 2:
        last, prev = s_twii.iloc[-1], s_twii.iloc[-2]
        dchg = pct(last, prev)
        lines.append(f"　{item_idx}. {pad_name('台股大盤')} {last:,.0f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
        core_items.append({"name": "台股大盤", "yf": "^TWII", "last": last, "ma20w": weekly_ma(s_twii, 20)})
        item_idx += 1
    
    # 抓台積/0050
    for w in watch:
        sym = w["yf_symbol"]
        if ("2330" in sym or "0050" in sym) and "正2" not in w["name"] and "反1" not in w["name"]:
            if sym in seen_core: continue
            seen_core.add(sym)
            s = closes.get(sym)
            if s is not None and len(s) >= 2:
                last, prev = s.iloc[-1], s.iloc[-2]
                dchg = pct(last, prev)
                display_name = w["name"] if "0050" not in sym or "元大" in w["name"] else "元大台灣50"
                lines.append(f"　{item_idx}. {pad_name(display_name)} {last:,.1f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
                core_items.append({"name": display_name, "yf": sym, "last": last, "ma20w": weekly_ma(s, 20)})
                item_idx += 1

    # ---------------- 2. VIX ----------------
    lines.extend(["\n▌二、VIX 恐慌指數 (<35) (今/昨/前日/上週)"])
    for i, (v_name, v_sym) in enumerate([("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")], 1):
        h = vix_hist.get(v_sym, [0,0,0,0])
        status = "🟢安全" if h[0] < 35 else "🔴恐慌"
        lines.append(f"　{i}. {pad_name(v_name)} {h[0]:.1f} / {h[1]:.1f} / {h[2]:.1f} / {h[3]:.1f} {status}")

    # ---------------- 3. 台股核心指標 ----------------
    lines.extend(["\n▌三、台股核心指標 (K<20) (K值 & 20週均)"])
    for i, item in enumerate(core_items, 1):
        kd = kd_data.get(item["yf"], {})
        kd_d = kd.get("日", (None,))[0]
        kd_w = kd.get("週", (None,))[0]
        kd_m = kd.get("月", (None,))[0]
        
        kd_str = f"日K{int(kd_d) if kd_d else '-'} / 週K{int(kd_w) if kd_w else '-'} / 月K{int(kd_m) if kd_m else '-'}"
        lines.append(f"　{i}. {pad_name(item['name'])} {kd_str}")
        
        last = item["last"]
        ma20w = item["ma20w"]
        if last and ma20w:
            dist = (last - ma20w) / ma20w * 100
            if dist < 0: status = "⚠️破線空手"
            else:
                if kd_d and kd_d < 20: status = "(日K超跌)"
                elif kd_w and kd_w < 20: status = "(週K超跌)"
                elif kd_m and kd_m < 20: status = "(月K超跌)"
                else: status = "(線上安全)"
            # 對齊處理
            lines.append(f"　   {'　'*6}距20週均({int(ma20w)}) : {'+' if dist>0 else ''}{dist:.1f}% {status}")

    # ---------------- 4. 其他位階警示 ----------------
    lines.extend(["\n▌四、其他位階警示"])
    alerts = []
    
    # 掃描所有標的產生警示
    all_targets = [{"name": n, "yf": s, "is_tw": False} for s, n in INDICES] +                   [{"name": w["name"], "yf": w["yf_symbol"], "is_tw": ("台" in w["market"] or ".TW" in w["yf_symbol"])} for w in watch]
    
    alert_count = 1
    for t in all_targets:
        if t["name"] not in ALERT_TARGETS:
            continue
            
        s = closes.get(t["yf"])
        if s is None or len(s) < 60: continue
        
        last = s.iloc[-1]
        ma = compute_ma(s)
        
        drop = []
        if ma.get("MA60") and last < ma.get("MA60"): drop.append("破季線")
        if ma.get("MA240") and last < ma.get("MA240"): drop.append("破年線")
        
        w_ret = pct(last, s.iloc[-6]) if len(s)>=6 else None
        m_ret = pct(last, s.iloc[-22]) if len(s)>=22 else None
        if (w_ret and w_ret <= -20) or (m_ret and m_ret <= -20): drop.append("週或月跌>20%")
        
        kd = kd_data.get(t["yf"], {})
        ko80 = [tf for tf in ["日", "週", "月"] if kd.get(tf) and kd[tf][0] is not None and kd[tf][0] > 80]
        
        res = ""
        if drop: res += f"🟢[跌] {','.join(drop)} "
        if ko80: res += f"🔴[過熱] {'/'.join(ko80)}K>80"
        
        if res: 
            alerts.append(f"　{alert_count}. {t['name']}：{res.strip()}")
            alert_count += 1

    if alerts:
        lines.extend(alerts)
    else:
        lines.append("　*(註：重點監控標的今日無觸發)*")

    # ---------------- 5. 結尾 ----------------
    report_url = os.environ.get("REPORT_URL", "https://your-report-link.com").strip()
    lines.extend([
        f"\n👉 五、詳細個股狀況 (網頁報表)：\n　 {report_url}",
        "\n💡 符號與策略說明：",
        "　 1. ☀️大漲(>2%) 🌤️小漲 🌧️小跌 ⛈️大跌(<-2%)",
        "　 2. 日/週/月 K<20 ＝ 超跌區間 (可關注)",
        "　 3. 破 20 週均線 ＝ 趨勢破壞 (建議離場/空手)",
        "　 4. VIX < 35 ＝ 🟢安全 / VIX > 35 = 🔴恐慌"
    ])

    text = "\n".join(lines).rstrip()
    return text[:LINE_MAX_LEN] + "\n…（訊息截斷）" if len(text) > LINE_MAX_LEN else text


def main():
    force_dry_run = "--dry-run" in sys.argv
    watch = load_watchlist()
    
    # 抓取所有收盤資料
    idx_syms = [s for s, _ in INDICES]
    stk_syms = [w["yf_symbol"] for w in watch]
    all_syms = list(set(idx_syms + stk_syms + ["^TWII"]))
    closes = fetch_data(all_syms)

    # 抓取 VIX
    v_data = fetch_data(["^VIX", "^VIXTWN"], period="1mo")
    vix_hist = {}
    for v_sym in ["^VIX", "^VIXTWN"]:
        s = v_data.get(v_sym)
        if s is not None and len(s) >= 6:
            # 今、昨、前日、上週(約5個交易日前)
            vix_hist[v_sym] = [s.iloc[-1], s.iloc[-2], s.iloc[-3], s.iloc[-6]]
        else:
            vix_hist[v_sym] = [0,0,0,0]

    # 計算 KD (全抓)
    kd_data = {}
    try:
        raw_kd = yf.download(all_syms, period="2y", interval="1d", auto_adjust=False, progress=False, group_by="ticker")
        for sym in all_syms:
            df = raw_kd if len(all_syms) == 1 else raw_kd[sym]
            if "Close" in df:
                df = df[["High", "Low", "Close"]].dropna()
                kd_data[sym] = kd_timeframes(df)
    except: pass

    # 組裝訊息
    text = build_message(closes, watch, kd_data, vix_hist)

    # 推播
    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    uid = os.environ.get("LINE_USER_ID", "").strip()
    if force_dry_run or not token or not uid:
        print("[DRY-RUN] 訊息預覽如下：\n\n" + text + "\n")
        return 0

    resp = requests.post(LINE_PUSH_URL, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                         json={"to": uid, "messages": [{"type": "text", "text": text}]}, timeout=20)
    resp.raise_for_status()
    print("[OK] 已推播至 LINE。")
    return 0

if __name__ == "__main__":
    sys.exit(main())
