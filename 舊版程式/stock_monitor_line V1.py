#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統
---------------------------------------------------------------------------
每週二、四、六 台灣時間 06:30，抓取美股與台股真實收盤價，計算跨期漲跌、
標示階梯式警示，組成精簡三段式純文字訊息，推播到使用者的 LINE。

資料源：Yahoo Finance（yfinance）。只用真實收盤價，不做任何模擬／推算。
清單來源：環境變數 SHEET_CSV_URL（Google 試算表發布的 CSV）；未設定時讀本機 watchlist.csv。

執行方式：
    python stock_monitor_line.py            # 有金鑰就推播，沒金鑰就印預覽（dry-run）
    python stock_monitor_line.py --dry-run  # 強制只印預覽、不推播

環境變數（金鑰，勿寫死在程式）：
    LINE_CHANNEL_ACCESS_TOKEN   LINE Messaging API 的 Channel access token
    LINE_USER_ID                推播目標的 User ID（U 開頭約 33 碼，非加好友 ID）
    SHEET_CSV_URL（可選）        Google 試算表發布的 CSV 網址
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

# 若本機有 python-dotenv 與 .env，自動載入金鑰（雲端用環境變數/Secret，不受影響）
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# ---------------------------------------------------------------------------
# 設定區（固定項）
# ---------------------------------------------------------------------------

# 四大指數：固定排序、不開放增刪
INDICES = [
    ("^DJI", "道瓊"),
    ("^IXIC", "那斯達克"),
    ("^GSPC", "S&P500"),
    ("^SOX", "費半"),
]

# 跨期區間（交易日數）
PERIODS = [("週", 5), ("月", 21), ("年", 252)]

# 階梯式警示門檻（依報酬率「絕對幅度」，漲跌皆適用）
THRESHOLDS = [
    (50, "💥"),  # 跨年倍數級變動
    (30, "⚡"),  # 趨勢主升 / 月爆發
    (20, "🚨"),  # 顯著波段 / 月強勢
    (10, "⚠️"),  # 短線異動 / 週突破
]

# 均線（日）：10 日、月均(20)、季均(60)、年均(240)
MA_WINDOWS = [("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]
MA20W_TARGETS = {"2330.TW", "0050.TW", "2454.TW"}  # 要加抓 20 週均的台股（大盤另外算）
TW_INDEX = ("^TWII", "台股大盤")          # 台股加權指數（加抓進資料檔）
KD_TARGETS = [("0050.TW", "0050 元大台灣50"), ("^TWII", "台股大盤")]  # 要抓 KD 的標的

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
LINE_MAX_LEN = 4800  # LINE 單則文字上限 5000，留安全餘裕

HERE = Path(__file__).resolve().parent
LOCAL_WATCHLIST = HERE / "watchlist.csv"

# 報表輸出資料夾（可用環境變數指到 Google 雲端硬碟同步資料夾）；未設定則用程式所在資料夾
OUTPUT_DIR = os.environ.get("REPORT_OUTPUT_DIR", "").strip() or str(HERE)
# 背後統計表的 Google 試算表網址（設定後，報表的統計表連結會指向它）
STATS_SHEET_URL = os.environ.get("STATS_SHEET_URL", "").strip()


# ---------------------------------------------------------------------------
# 清單載入
# ---------------------------------------------------------------------------

def _normalize_row(row):
    """把一列清單資料正規化成 dict，並補上 yf_symbol（台股自動加 .TW）。"""
    market = (row.get("market") or "").strip()
    code = (row.get("code") or "").strip()
    yf_symbol = (row.get("yf_symbol") or "").strip()
    name = (row.get("name") or "").strip()
    group = (row.get("group") or "").strip() or "觀察"
    if not code and not yf_symbol:
        return None
    if not yf_symbol:
        # 台股或純數字代號 → 加 .TW；其餘沿用代號
        if market == "台股" or code.replace(".", "").isdigit():
            yf_symbol = code if code.endswith(".TW") else f"{code}.TW"
        else:
            yf_symbol = code
    if not name:
        name = code or yf_symbol
    return {"market": market, "code": code or yf_symbol,
            "yf_symbol": yf_symbol, "name": name, "group": group}


def load_watchlist():
    """讀取股票清單。優先讀 SHEET_CSV_URL，失敗或未設定則讀本機 watchlist.csv。"""
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
    reader = csv.DictReader(io.StringIO(text))
    for raw in reader:
        norm = _normalize_row({k.strip(): (v or "") for k, v in raw.items()})
        if norm:
            rows.append(norm)
    return rows


# ---------------------------------------------------------------------------
# 數據擷取與計算
# ---------------------------------------------------------------------------

def _extract_close(data, symbol, n_symbols):
    """從 yf.download 的結果取出某代碼的收盤價 Series（處理單檔/多檔欄位差異）。"""
    try:
        close = data["Close"]
    except Exception:
        return pd.Series(dtype="float64")
    if isinstance(close, pd.DataFrame):
        if symbol in close.columns:
            s = close[symbol]
        elif n_symbols == 1 and close.shape[1] == 1:
            s = close.iloc[:, 0]
        else:
            return pd.Series(dtype="float64")
    else:
        s = close
    return s.dropna()


def repair_splits(series, symbol="", low=0.4, high=2.5):
    """修復 Yahoo 漏記的股票分割。

    auto_adjust 只能還原 Yahoo 有記錄的分割；若某檔分割未被收錄（例：0052.TW），
    會在某日出現「單日剩 1/N」的斷崖，使跨期報酬變成假訊號（甚至反轉正負號）。
    這裡偵測單日跳動超出合理範圍（≤low 或 ≥high，一般股票/ETF/指數不可能的幅度），
    判定為漏記的分割，將斷點之前的歷史價格乘上該比例，使序列連續。
    閾值刻意保守，真實的單日大漲大跌（例 -40%）不會被誤修。
    """
    if len(series) < 2:
        return series
    s = series.copy()
    for i in range(len(series) - 1, 0, -1):
        prev = float(series.iloc[i - 1])
        if prev == 0:
            continue
        ratio = float(series.iloc[i]) / prev
        if ratio <= low or ratio >= high:
            s.iloc[:i] = s.iloc[:i] * ratio
            when = series.index[i]
            when = when.date() if hasattr(when, "date") else when
            print(f"[修復] {symbol} 偵測到未還原的分割 @ {when}（比例 {ratio:.3f}），"
                  f"已重新對齊歷史價格。", file=sys.stderr)
    return s


def fetch_data(symbols):
    """一次下載所有代碼近 2 年日線（原始成交價，未還原配息，貼近看盤軟體），並修復漏記的分割。"""
    data = yf.download(
        symbols, period="2y", interval="1d",
        auto_adjust=False, progress=False, threads=True, group_by="column",
    )
    closes = {}
    for sym in symbols:
        s = _extract_close(data, sym, len(symbols))
        if len(s) >= 2:
            closes[sym] = repair_splits(s, symbol=sym)
    # 批次下載偶爾會漏某幾檔 → 對漏抓的代碼個別重試
    missing = [s for s in symbols if s not in closes]
    for sym in missing:
        try:
            d1 = yf.download(sym, period="2y", interval="1d",
                             auto_adjust=False, progress=False)
            s = _extract_close(d1, sym, 1)
            if len(s) >= 2:
                closes[sym] = repair_splits(s, symbol=sym)
        except Exception as e:
            print(f"[警告] {sym} 重試仍失敗：{e}", file=sys.stderr)
    return closes


def pct(new, old):
    if old is None or old == 0:
        return None
    return (new / old - 1.0) * 100.0


def compute_period_returns(series):
    """回傳 {區間名: 報酬率%}；資料不足的區間為 None。"""
    last = float(series.iloc[-1])
    out = {}
    for name, days in PERIODS:
        if len(series) >= days + 1:
            out[name] = pct(last, float(series.iloc[-(days + 1)]))
        else:
            out[name] = None
    return out


def compute_period_detail(series):
    """回傳 {區間名: {ret, price, date}}：含該區間『N 交易日前的實際收盤價與日期』，
    以利追溯週/月/年報酬是怎麼算出來的。"""
    last = float(series.iloc[-1])
    out = {}
    for name, days in PERIODS:
        if len(series) >= days + 1:
            ref = float(series.iloc[-(days + 1)])
            out[name] = {"ret": pct(last, ref), "price": ref,
                         "date": series.index[-(days + 1)].date()}
        else:
            out[name] = {"ret": None, "price": None, "date": None}
    return out


def daily_change(series):
    """當日漲跌%（最新 vs 前一交易日）。"""
    if len(series) < 2:
        return None
    return pct(float(series.iloc[-1]), float(series.iloc[-2]))


# ---- 均線 / KD（抓進資料檔用；全部由真實收盤價自行計算，算不出來回 None） ----

def compute_ma(series):
    """日均線 MA10/20/60/240。資料不足回 None。"""
    out = {}
    for name, w in MA_WINDOWS:
        out[name] = round(float(series.rolling(w).mean().iloc[-1]), 2) if len(series) >= w else None
    return out


def weekly_ma(series, weeks):
    """N 週均（以週線收盤計算）。資料不足回 None。"""
    wk = series.resample("W").last().dropna()
    if len(wk) < weeks:
        return None
    return round(float(wk.rolling(weeks).mean().iloc[-1]), 2)


def price_level(last, ma):
    """股價位階：最新價相對 雙週線(MA10)/月線(MA20)/季線(MA60)/年線(MA240) 的位置。
    由上而下判斷：站上雙週線 → 雙週~月 → 月~季 → 季~年 → 年線下。"""
    if last is None or not ma:
        return ""
    m10, m20, m60, m240 = ma.get("MA10"), ma.get("MA20"), ma.get("MA60"), ma.get("MA240")
    if m10 is not None and last >= m10:
        return "雙週線上"
    if m20 is not None and last >= m20:
        return "雙週~月"
    if m60 is not None and last >= m60:
        return "月~季"
    if m240 is not None and last >= m240:
        return "季~年"
    if m240 is not None:
        return "年線下"
    return ""


def ma20w_level(last, ma20w):
    """20 週均線位階：最新價在其上或已跌破。"""
    if last is None or ma20w is None:
        return ""
    return "線上" if last >= ma20w else "破線"


def _ohlc_frame(data, symbol):
    """從 yf.download 結果取出 High/Low/Close 乾淨 DataFrame。"""
    if isinstance(data.columns, pd.MultiIndex):
        try:
            df = data.xs(symbol, axis=1, level=-1)
        except Exception:
            df = data.droplevel(-1, axis=1)
    else:
        df = data
    return df[["High", "Low", "Close"]].dropna()


def compute_kd(df, n=9):
    """KD(9,3,3)：回傳最新 (K, D)（小數 1 位）；資料不足回 (None, None)。"""
    if df is None or len(df) < n:
        return (None, None)
    low = df["Low"].rolling(n).min()
    high = df["High"].rolling(n).max()
    rng = (high - low)
    k = d = 50.0
    for i in range(len(df)):
        denom = rng.iloc[i]
        if pd.isna(denom) or denom == 0:
            rsv = 50.0
        else:
            rsv = float((df["Close"].iloc[i] - low.iloc[i]) / denom * 100)
        k = 2 / 3 * k + 1 / 3 * rsv
        d = 2 / 3 * d + 1 / 3 * k
    return (round(k, 1), round(d, 1))


def kd_timeframes(df):
    """日/週/月 KD；回傳 {'日':(K,D),'週':(K,D),'月':(K,D)}。"""
    agg = {"High": "max", "Low": "min", "Close": "last"}
    out = {"日": compute_kd(df)}
    try:
        out["週"] = compute_kd(df.resample("W").agg(agg).dropna())
    except Exception:
        out["週"] = (None, None)
    for rule in ("ME", "M"):
        try:
            out["月"] = compute_kd(df.resample(rule).agg(agg).dropna())
            break
        except Exception:
            out["月"] = (None, None)
    return out


def fetch_extras():
    """抓『資料檔用』的額外資料：台股大盤(^TWII) 與 0050/大盤 的日週月 KD。"""
    extras = {"twii": None, "kd": []}
    try:
        s = _extract_close(yf.download(TW_INDEX[0], period="2y", interval="1d",
                                       auto_adjust=False, progress=False), TW_INDEX[0], 1)
        if len(s) >= 2:
            detail = compute_period_detail(s)
            extras["twii"] = {"yf": TW_INDEX[0], "name": TW_INDEX[1],
                              "last": float(s.iloc[-1]), "asof": s.index[-1].date(),
                              "dchg": daily_change(s), "ma20w": weekly_ma(s, 20),
                              "detail": detail, "rets": {k: detail[k]["ret"] for k in detail},
                              "ma": compute_ma(s)}
    except Exception as e:
        print(f"[警告] {TW_INDEX[0]} 抓取失敗：{e}", file=sys.stderr)
    for sym, name in KD_TARGETS:
        try:
            raw = yf.download(sym, period="2y", interval="1d",
                              auto_adjust=False, progress=False)
            df = _ohlc_frame(raw, sym)
            extras["kd"].append({"code": sym, "name": name,
                                 "asof": df.index[-1].date() if len(df) else None,
                                 "tf": kd_timeframes(df)})
        except Exception as e:
            print(f"[警告] {sym} KD 抓取失敗：{e}", file=sys.stderr)
            extras["kd"].append({"code": sym, "name": name, "asof": None, "tf": {}})
    return extras


def alert_symbol(pct_value):
    """依報酬率絕對幅度回傳警示符號；未達門檻回空字串。"""
    if pct_value is None:
        return ""
    a = abs(pct_value)
    for thr, sym in THRESHOLDS:
        if a >= thr:
            return sym
    return ""


# ---------------------------------------------------------------------------
# 訊息組裝
# ---------------------------------------------------------------------------

def fmt_pct(v, decimals=1):
    """帶 ▲▼ 與 % 的百分比（給大盤、警示用）。"""
    if v is None:
        return "—"
    arrow = "▲" if v >= 0 else "▼"
    return f"{arrow}{abs(v):.{decimals}f}%"


def fmt_val(v):
    """個股區塊用的精簡值：帶 ▲▼、不帶 %（表頭已標週/月/年）。"""
    if v is None:
        return "—"
    arrow = "▲" if v >= 0 else "▼"
    return f"{arrow}{abs(v):.1f}"


def build_message(index_rows, stock_rows):
    """組三段式純文字訊息（為手機 LINE 比例字體優化排版）。
    index_rows: list of (name, last, daily%)
    stock_rows: list of dict(name, code, group, market, rets={週/月/年})
    """
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    lines = [f"📊 美股／台股監控　{today.month}/{today.day:02d}（{wd}）"]

    # 第一段：大盤現況
    lines.append("")
    lines.append("━━ 大盤 ━━")
    for idx in index_rows:
        name, last, dchg = idx["name"], idx["last"], idx["dchg"]
        if last is None:
            lines.append(f"{name}　資料暫缺")
        else:
            lines.append(f"{name}　{last:,.0f}　{fmt_pct(dchg)}")

    # 第二段：異動警示（漲跌都報，下跌特別標示）
    lines.append("")
    lines.append("━━ 異動警示 ━━")
    alerts = []
    for st in stock_rows:
        # 取絕對幅度最大的區間作為該檔代表警示
        best = None
        for pname, _ in PERIODS:
            v = st["rets"].get(pname)
            if v is None:
                continue
            if abs(v) >= 10 and (best is None or abs(v) > abs(best[1])):
                best = (pname, v)
        if best:
            pname, v = best
            sym = alert_symbol(v)
            tag = "　逢低" if v < 0 else ""
            alerts.append((abs(v), f"{sym} {st['code']}　{pname}{fmt_pct(v, 0)}{tag}"))
    if alerts:
        alerts.sort(key=lambda x: x[0], reverse=True)
        for _, line in alerts:
            lines.append(line)
    else:
        lines.append("今日無觸發警示門檻")

    # 第三段：個股數據（週/月/年），美股與台股分段
    lines.append("")
    lines.append("━━ 個股　週 / 月 / 年 ━━")
    for market_label in ("美股", "台股"):
        subset = [s for s in stock_rows if s["market"] == market_label]
        if not subset:
            continue
        lines.append(f"〔{market_label}〕")
        for st in subset:
            r = st["rets"]
            top = max((abs(v) for v in r.values() if v is not None), default=0)
            sym = alert_symbol(top if top >= 10 else None)
            trio = f"{fmt_val(r.get('週'))} / {fmt_val(r.get('月'))} / {fmt_val(r.get('年'))}"
            lines.append(f"{st['code']}　{trio}{('　' + sym) if sym else ''}")

    text = "\n".join(lines).rstrip()
    if len(text) > LINE_MAX_LEN:
        text = text[:LINE_MAX_LEN] + "\n…（訊息過長已截斷）"
    return text


# ---------------------------------------------------------------------------
# LINE 推播
# ---------------------------------------------------------------------------

def push_line(text, force_dry_run=False):
    """推播到 LINE；未設金鑰或 --dry-run 時只印預覽。回傳是否真的送出。"""
    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    uid = os.environ.get("LINE_USER_ID", "").strip()
    if force_dry_run or not token or not uid:
        reason = "指定 --dry-run" if force_dry_run else "未設定 LINE 金鑰"
        print(f"[DRY-RUN：{reason}] 訊息預覽如下：\n")
        print(text)
        print("\n[DRY-RUN] 以上為預覽，未實際推播。")
        return False
    resp = requests.post(
        LINE_PUSH_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"to": uid, "messages": [{"type": "text", "text": text}]},
        timeout=20,
    )
    resp.raise_for_status()
    print("[OK] 已推播至 LINE。")
    return True


# ---------------------------------------------------------------------------
# 主流程（含重試）
# ---------------------------------------------------------------------------

def gather_data():
    """抓取並整理成結構化資料：(index_rows, stock_rows)。"""
    watch = load_watchlist()
    index_syms = [s for s, _ in INDICES]
    stock_syms = [w["yf_symbol"] for w in watch]
    closes = fetch_data(index_syms + stock_syms)

    index_rows = []
    for sym, name in INDICES:
        s = closes.get(sym)
        if s is None:
            index_rows.append({"name": name, "yf": sym, "last": None, "dchg": None,
                               "asof": None, "detail": {}, "rets": {}, "ma": {}})
        else:
            detail = compute_period_detail(s)
            index_rows.append({"name": name, "yf": sym, "last": float(s.iloc[-1]),
                               "dchg": daily_change(s), "asof": s.index[-1].date(),
                               "detail": detail, "rets": {k: detail[k]["ret"] for k in detail},
                               "ma": compute_ma(s)})

    stock_rows = []
    for w in watch:
        s = closes.get(w["yf_symbol"])
        if s is None:
            continue
        detail = compute_period_detail(s)
        stock_rows.append({
            "name": w["name"], "code": w["code"], "yf": w["yf_symbol"], "group": w["group"],
            "market": w["market"] or ("台股" if w["yf_symbol"].endswith(".TW") else "美股"),
            "last": float(s.iloc[-1]), "asof": s.index[-1].date(),
            "dchg": daily_change(s),
            "detail": detail,
            "rets": {k: detail[k]["ret"] for k in detail},
            "ma": compute_ma(s),
            "ma20w": weekly_ma(s, 20) if w["yf_symbol"] in MA20W_TARGETS else None,
        })
    return index_rows, stock_rows


# ---------------------------------------------------------------------------
# 網頁報表（在電腦上看，表格對齊、紅漲綠跌、可點連結）
# ---------------------------------------------------------------------------

OVERSOLD = [(50, "💥"), (30, "⚡")]  # 超跌門檻（跌幅）：只顯示 ≥30%


def oversold_mark(v):
    """跌幅達門檻回符號；否則空字串。v 為報酬率（負值為跌）。"""
    if v is None or v >= 0:
        return ""
    a = -v
    for thr, sym in OVERSOLD:
        if a >= thr:
            return sym
    return ""


def _html_num(v, suffix="%"):
    if v is None:
        return '<td class="num mute">—</td>'
    arrow = "▲" if v >= 0 else "▼"
    cls = "up" if v >= 0 else "down"
    return f'<td class="num {cls}">{arrow}{abs(v):.1f}{suffix}</td>'


NEWS_KEYS = (
    "台股", "美股", "大盤", "加權", "指數", "道瓊", "那斯達克", "納斯達克", "那指",
    "標普", "S&P", "費半", "費城半導體", "半導體", "科技股", "台積", "聯發科",
    "輝達", "Nvidia", "AI", "外資", "台指", "期指", "盤後", "盤中", "華爾街",
    "升息", "降息", "聯準會", "Fed", "財報", "ETF", "選擇權", "2330", "2454",
)


def fetch_news(count=5):
    """抓取 Yahoo 台灣財經中文新聞，優先『美股／台股大盤相關』；失敗回空清單。"""
    import xml.etree.ElementTree as ET
    urls = (
        "https://tw.stock.yahoo.com/rss?category=news",
        "https://tw.news.yahoo.com/rss/finance",
    )
    allitems, seen = [], set()
    for u in urls:
        try:
            r = requests.get(u, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
            if r.status_code != 200 or "<item" not in r.text:
                continue
            root = ET.fromstring(r.content)
            for it in root.findall(".//item"):
                t = (it.findtext("title") or "").strip()
                link = (it.findtext("link") or "").strip()
                if t and link and t not in seen:
                    seen.add(t)
                    allitems.append((t, link))
        except Exception as e:
            print(f"[警告] 新聞抓取失敗（{u}）：{e}", file=sys.stderr)
    if not allitems:
        return []
    matched = [x for x in allitems if any(k in x[0] for k in NEWS_KEYS)]
    rest = [x for x in allitems if x not in matched]
    return (matched + rest)[:count]


def light_for(x):
    """單一指數的燈號：依當日漲跌幅 ±2/4/6% → 平常/小喜(憂)/大喜(憂)/狂喜(憂)。"""
    if x is None:
        return ("", "")
    a = abs(x)
    up = x >= 0
    if a < 2:
        return ("⚪", "平常")
    if a < 4:
        return ("🔴" if up else "🟢", "小喜" if up else "小憂")
    if a < 6:
        return ("🔴" if up else "🟢", "大喜" if up else "大憂")
    return ("🔴🔴" if up else "🟢🟢", "狂喜" if up else "狂憂")


def build_html(index_rows, stock_rows, news=None):
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    now = today.strftime("%Y/%m/%d %H:%M")
    yq = "https://finance.yahoo.com/quote/"
    d = today.date()
    ref_w = (d - dt.timedelta(days=7)).strftime("%m/%d")
    ref_m = (d - dt.timedelta(days=30)).strftime("%m/%d")
    ref_y = (d - dt.timedelta(days=365)).strftime("%Y/%m")
    from collections import Counter
    _dates = [st["asof"] for st in stock_rows if st.get("asof")]
    rep_date = Counter(_dates).most_common(1)[0][0].strftime("%m/%d") if _dates else ""

    # 市場新聞（中文，放最下方）
    news = news or []
    if news:
        items = "".join(f'<li><a href="{lk}" target="_blank">{t}</a></li>' for t, lk in news)
        news_html = f'<ol class="news">{items}</ol>'
    else:
        news_html = '<p class="mute">暫時取不到新聞，稍後再試。</p>'

    stats_link = (f'<a href="{STATS_SHEET_URL}" target="_blank">Google 試算表</a>'
                  if STATS_SHEET_URL else '<a href="股價數據抓取_Yahoo.csv">股價數據抓取_Yahoo.csv</a>')

    # 大盤
    idx_html = []
    for r in index_rows:
        if r["last"] is None:
            idx_html.append(f'<tr><td>{r["name"]}</td><td class="num mute">資料暫缺</td><td></td><td></td></tr>')
        else:
            dc = r["dchg"]
            cls = "" if dc is None else ("up" if dc >= 0 else "down")
            dctxt = "—" if dc is None else f'{"▲" if dc>=0 else "▼"}{abs(dc):.1f}%'
            ldot, llbl = light_for(dc)
            idx_html.append(
                f'<tr><td><a href="{yq}{r["yf"]}" target="_blank">{r["name"]}</a></td>'
                f'<td class="num">{r["last"]:,.0f}</td>'
                f'<td class="num {cls}">{dctxt}</td>'
                f'<td class="light-cell">{ldot} {llbl}</td></tr>')

    # 個股（分群，含序號）
    stk_html = []
    n = 0
    for market_label in ("美股", "台股"):
        subset = [s for s in stock_rows if s["market"] == market_label]
        if not subset:
            continue
        stk_html.append(f'<tr class="grp"><td colspan="7">{market_label}</td></tr>')
        for st in subset:
            n += 1
            r = st["rets"]
            name = f'<a href="{yq}{st["yf"]}" target="_blank">{st["name"]}</a> <span class="code">{st["code"]}</span>'
            stk_html.append(
                f'<tr><td class="num idx">{n}</td><td class="name">{name}</td>'
                f'<td class="num">{st["last"]:,.2f}</td>'
                f'{_html_num(st.get("dchg"))}'
                f'{_html_num(r.get("週"))}{_html_num(r.get("月"))}{_html_num(r.get("年"))}</tr>')

    # 超跌警示（只顯示跌幅 ≥30%）
    over = []
    for st in stock_rows:
        worst = None
        for pname, _ in PERIODS:
            v = st["rets"].get(pname)
            if v is None:
                continue
            if v <= -30 and (worst is None or v < worst[1]):
                worst = (pname, v)
        if worst:
            pname, v = worst
            over.append((v, f'<li>{oversold_mark(v)} <b>{st["name"]}</b> {st["code"]}　{pname}線 ▼{abs(v):.1f}%</li>'))
    if over:
        over.sort(key=lambda x: x[0])
        over_html = "<ul class='over'>" + "".join(h for _, h in over) + "</ul>"
    else:
        over_html = "<p class='mute'>今日無超跌（跌幅未達 30%）。</p>"

    return f"""<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>美股／台股監控 {today.month}/{today.day:02d}</title>
<style>
  :root {{ --up:#d32f2f; --down:#1b8e3a; --line:#e3e6ea; --mute:#8a9099; --head:#1f3b66; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:#f6f7f9; color:#222;
    font-family:"Microsoft JhengHei","PingFang TC",system-ui,sans-serif; font-size:13px; }}
  .wrap {{ max-width:720px; margin:0 auto; padding:16px; }}
  h1 {{ font-size:18px; margin:4px 0; }}
  h2 {{ font-size:14px; margin:18px 0 6px; color:var(--head);
    border-left:4px solid var(--head); padding-left:8px; }}
  .sub {{ color:var(--mute); font-size:12px; margin-bottom:6px; }}
  table {{ width:100%; border-collapse:collapse; background:#fff;
    border:1px solid var(--line); border-radius:8px; overflow:hidden; }}
  th,td {{ padding:6px 8px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }}
  thead th {{ background:#eef1f5; color:#333; font-weight:600; font-size:12px; }}
  thead th.num {{ text-align:center; }}
  thead th small {{ color:var(--mute); font-weight:400; font-size:10px; }}
  tbody tr:last-child td {{ border-bottom:none; }}
  td.num {{ text-align:center; font-variant-numeric:tabular-nums; }}
  td.idx {{ color:var(--mute); width:30px; text-align:center; }}
  td.name {{ white-space:normal; }}
  .code {{ color:var(--mute); font-size:11px; }}
  .date {{ color:var(--mute); }}
  .up {{ color:var(--up); }} .down {{ color:var(--down); }} .mute {{ color:var(--mute); }}
  tr.grp td {{ background:#f0f3f7; font-weight:700; color:var(--head); }}
  .light {{ background:#fff; border:1px solid var(--line); border-radius:8px;
    padding:8px 10px; margin:6px 0 0; font-size:13px; }}
  .light b {{ font-size:14px; }}
  ol.news {{ margin:4px 0 0; padding-left:20px; }}
  ol.news li {{ margin:5px 0; line-height:1.5; }}
  ol.news a {{ color:#1456b8; text-decoration:none; }}
  ol.news a:hover {{ text-decoration:underline; }}
  .foot a {{ color:#1456b8; }}
  td.light-cell {{ white-space:nowrap; text-align:center; }}
  ul.over {{ list-style:none; padding:0; margin:0; }}
  ul.over li {{ background:#fff; border:1px solid var(--line); border-radius:6px;
    padding:6px 10px; margin-bottom:4px; }}
  .foot {{ color:var(--mute); font-size:11px; margin-top:14px; line-height:1.6; }}
</style></head>
<body><div class="wrap">
  <h1>📊 美股／台股監控</h1>
  <div class="sub">{today.year}/{today.month:02d}/{today.day:02d}（{wd}）　產生時間 {now}</div>

  <h2>大盤狀況</h2>
  <table><thead><tr><th>指數</th><th class="num">收盤</th><th class="num">當日</th><th class="num">燈號</th></tr></thead>
  <tbody>{"".join(idx_html)}</tbody></table>

  <h2>個股（最新收盤 {rep_date}／日漲跌／漲跌幅 週·月·年）</h2>
  <table><thead><tr><th class="num">#</th><th>名稱</th><th class="num">收盤</th><th class="num">當日</th>
  <th class="num">週<br><small>≈{ref_w}</small></th><th class="num">月<br><small>≈{ref_m}</small></th>
  <th class="num">年<br><small>≈{ref_y}</small></th></tr></thead>
  <tbody>{"".join(stk_html)}</tbody></table>

  <h2>超跌警示（跌幅 ≥30%）</h2>
  {over_html}

  <h2>市場新聞（Yahoo 財經）</h2>
  {news_html}

  <div class="foot">
  股價來源網頁：<a href="https://finance.yahoo.com/" target="_blank">finance.yahoo.com</a>（點股名可開個股頁）。<br>
  背後統計表（含原始收盤價、日期、來源，可逐筆追溯）：{stats_link}<br>
  紅＝漲、綠＝跌（台股慣例）。週/月/年％ ＝ 最新收盤 ÷（5 / 21 / 252 交易日前還原收盤）− 1。</div>
</div></body></html>"""


def _r1(v):
    return "" if v is None else round(v, 1)


STATS_HEAD = [
    "序號", "類別", "市場", "代碼", "名稱", "股價位階", "20周均線位階",
    "最新收盤", "最新日期", "當日%",
    "週前收盤", "週前日期", "週%", "月前收盤", "月前日期", "月%",
    "年前收盤", "年前日期", "年%",
    "MA10(雙週)", "MA20(月)", "MA60(季)", "MA240(年)", "MA20週", "來源網址",
]


def write_stats_csv(index_rows, stock_rows, path, extras=None):
    """輸出背後統計表：序號首欄全列連號、股價位階、原始收盤價/日期、均線、來源，可逐筆追溯。
    （KD 另存獨立檔 write_kd_csv。）"""
    extras = extras or {}
    yq = "https://finance.yahoo.com/quote/"

    def price(v):
        return "" if v is None else round(v, 2)

    def mav(v):
        return "" if v is None else v

    def row_for(seq, category, market, code, name, r):
        m = r.get("ma") or {}
        d = r.get("detail") or {}
        ma20w = r.get("ma20w")

        def cell(p):
            x = d.get(p) or {}
            return [price(x.get("price")), x.get("date") or "", _r1(x.get("ret"))]

        row = [seq, category, market, code, name,
               price_level(r.get("last"), m), ma20w_level(r.get("last"), ma20w),
               price(r.get("last")), r.get("asof") or "", _r1(r.get("dchg"))]
        row += cell("週") + cell("月") + cell("年")
        row += [mav(m.get("MA10")), mav(m.get("MA20")), mav(m.get("MA60")),
                mav(m.get("MA240")), mav(ma20w)]
        row += [yq + r.get("yf", code)]
        return row

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(STATS_HEAD)
        n = 0
        # 指數（四大 + 台股大盤），全列連號
        idx_rows = list(index_rows)
        if extras.get("twii"):
            idx_rows = idx_rows + [extras["twii"]]
        for r in idx_rows:
            n += 1
            w.writerow(row_for(n, "指數", "", r["yf"], r["name"], r))
        # 個股
        for st in stock_rows:
            n += 1
            w.writerow(row_for(n, "個股", st["market"], st["code"], st["name"], st))


def write_kd_csv(path, extras):
    """KD(9,3,3) 獨立檔：0050 與台股大盤的日/週/月 K、D（序號首欄）。"""
    kd = (extras or {}).get("kd") or []
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序號", "代碼", "名稱", "日K", "日D", "週K", "週D",
                    "月K", "月D", "資料日期"])
        for i, item in enumerate(kd, 1):
            tf = item.get("tf") or {}

            def kv(period, idx):
                t = tf.get(period)
                return "" if not t or t[idx] is None else t[idx]

            w.writerow([i, item["code"], item["name"],
                        kv("日", 0), kv("日", 1), kv("週", 0), kv("週", 1),
                        kv("月", 0), kv("月", 1), item.get("asof") or ""])


XLSX_PRICE_COLORS = {"雙週線上": "C6EFCE", "雙週~月": "E2F0D9", "月~季": "FFF2CC",
                     "季~年": "FCE4D6", "年線下": "FFC7CE"}
XLSX_MA20W_COLORS = {"線上": "C6EFCE", "破線": "FFC7CE"}
XLSX_NUMERIC_HEADERS = {
    "序號", "最新收盤", "當日%", "週前收盤", "週%", "月前收盤", "月%",
    "年前收盤", "年%", "MA10(雙週)", "MA20(月)", "MA60(季)", "MA240(年)", "MA20週",
    "日K", "日D", "週K", "週D", "月K", "月D",
}


def _cjk_width(cells):
    w = 6
    for c in cells:
        s = str(c) if c is not None else ""
        ln = sum(2 if ord(ch) > 0x2E00 else 1 for ch in s)
        w = max(w, ln + 2)
    return min(w, 42)


def _num(s):
    try:
        f = float(s)
        return int(f) if f.is_integer() and abs(f) < 1e15 else f
    except (ValueError, TypeError):
        return s


def write_xlsx(csv_path, kd_path, xlsx_path):
    """由 CSV 產生『排版好的 Excel』：位階上色、數字轉數值、最適欄寬、標題篩選、凍結表頭。"""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    center = Alignment(horizontal="center", vertical="center")
    hfill = PatternFill("solid", fgColor="D9E7FD")

    def add_sheet(title, rows, color=False):
        ws = wb.create_sheet(title)
        if not rows:
            return
        header = rows[0]
        ws.append(header)
        for r in rows[1:]:
            ws.append([_num(v) if (i < len(header) and header[i] in XLSX_NUMERIC_HEADERS) else v
                       for i, v in enumerate(r)])
        ncol, nrow = len(header), len(rows)
        for row in ws.iter_rows(min_row=1, max_row=nrow, max_col=ncol):
            for cell in row:
                cell.alignment = center
        for c in range(1, ncol + 1):
            ws.cell(row=1, column=c).font = Font(bold=True)
            ws.cell(row=1, column=c).fill = hfill
            ws.column_dimensions[get_column_letter(c)].width = _cjk_width(
                [rows[r][c - 1] if c - 1 < len(rows[r]) else "" for r in range(nrow)])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(ncol), nrow)
        if color:
            for name, cmap in (("股價位階", XLSX_PRICE_COLORS), ("20周均線位階", XLSX_MA20W_COLORS)):
                if name in header:
                    ci = header.index(name) + 1
                    for r in range(2, nrow + 1):
                        v = str(ws.cell(row=r, column=ci).value or "").strip()
                        if v in cmap:
                            ws.cell(row=r, column=ci).fill = PatternFill("solid", fgColor=cmap[v])

    main_rows = list(csv.reader(open(csv_path, encoding="utf-8-sig")))
    kd_rows = list(csv.reader(open(kd_path, encoding="utf-8-sig")))
    add_sheet("主表", main_rows, color=True)
    add_sheet("KD", kd_rows)
    wb.save(xlsx_path)


def _write_hist_csv(df, path):
    """寬表歷史收盤：序號, 日期, 各代碼收盤（原始價、已修復分割）。"""
    df = df.dropna(how="all").sort_index()
    cols = list(df.columns)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序號", "日期"] + cols)
        for i, (dt_, row) in enumerate(df.iterrows(), 1):
            vals = ["" if pd.isna(v) else round(float(v), 2) for v in row]
            w.writerow([i, dt_.strftime("%Y-%m-%d")] + vals)


def write_history(outdir):
    """歷史資料庫：全標的（指數＋台股大盤＋個股）日/週/月收盤，存雲端 CSV，日後計算免重抓。"""
    watch = load_watchlist()
    syms = [s for s, _ in INDICES] + [TW_INDEX[0]] + [w["yf_symbol"] for w in watch]
    closes = fetch_data(syms)
    cols = [s for s in syms if s in closes]
    if not cols:
        return
    df = pd.DataFrame({s: closes[s] for s in cols}).sort_index()
    outdir = Path(outdir)
    _write_hist_csv(df, outdir / "歷史_日線.csv")
    _write_hist_csv(df.resample("W").last(), outdir / "歷史_週線.csv")
    for rule in ("ME", "M"):
        try:
            _write_hist_csv(df.resample(rule).last(), outdir / "歷史_月線.csv")
            break
        except Exception:
            continue


def write_html(index_rows, stock_rows, path=None):
    outdir = Path(OUTPUT_DIR)
    try:
        outdir.mkdir(parents=True, exist_ok=True)
    except Exception:
        outdir = HERE
    path = path or str(outdir / "report.html")
    news = fetch_news(5)
    extras = fetch_extras()
    write_stats_csv(index_rows, stock_rows,
                    str(Path(path).with_name("股價數據抓取_Yahoo.csv")), extras)
    write_kd_csv(str(Path(path).with_name("股價KD_Yahoo.csv")), extras)
    try:
        write_xlsx(str(Path(path).with_name("股價數據抓取_Yahoo.csv")),
                   str(Path(path).with_name("股價KD_Yahoo.csv")),
                   str(Path(path).with_name("股價數據抓取_Yahoo.xlsx")))
    except Exception as e:
        print(f"[警告] xlsx 產生失敗：{e}", file=sys.stderr)
    try:
        write_history(Path(path).parent)
    except Exception as e:
        print(f"[警告] 歷史資料輸出失敗：{e}", file=sys.stderr)
    html = build_html(index_rows, stock_rows, news)
    Path(path).write_text(html, encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# 主流程（含重試）
# ---------------------------------------------------------------------------

def main():
    force_dry_run = "--dry-run" in sys.argv
    html_mode = "--html" in sys.argv
    last_err = None
    for attempt, wait in enumerate([5, 10, 20, 0], start=1):
        try:
            index_rows, stock_rows = gather_data()
            if html_mode:
                path = write_html(index_rows, stock_rows)
                print(f"[OK] 網頁報表已產生：{path}")
            else:
                text = build_message(index_rows, stock_rows)
                push_line(text, force_dry_run=force_dry_run)
            return 0
        except Exception as e:
            last_err = e
            print(f"[重試 {attempt}] 失敗：{e}", file=sys.stderr)
            if wait:
                time.sleep(wait)
    # 全部重試失敗
    if html_mode:
        print(f"[失敗] 產生報表失敗：{last_err}", file=sys.stderr)
        return 1
    msg = f"⚠️ 今日數據抓取異常，稍後重試。\n（{type(last_err).__name__}: {last_err}）"
    try:
        push_line(msg, force_dry_run=force_dry_run)
    except Exception as e:
        print(f"[嚴重] 連異常通知都推播失敗：{e}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
