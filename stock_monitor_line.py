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

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
LINE_MAX_LEN = 4800  # LINE 單則文字上限 5000，留安全餘裕

HERE = Path(__file__).resolve().parent
LOCAL_WATCHLIST = HERE / "watchlist.csv"


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
    """一次下載所有代碼近 2 年日線（還原收盤價），並修復漏記的分割。"""
    data = yf.download(
        symbols, period="2y", interval="1d",
        auto_adjust=True, progress=False, threads=True, group_by="column",
    )
    closes = {}
    for sym in symbols:
        s = _extract_close(data, sym, len(symbols))
        if len(s) >= 2:
            closes[sym] = repair_splits(s, symbol=sym)
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


def daily_change(series):
    """當日漲跌%（最新 vs 前一交易日）。"""
    if len(series) < 2:
        return None
    return pct(float(series.iloc[-1]), float(series.iloc[-2]))


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

def fmt_pct(v):
    if v is None:
        return "—"
    arrow = "▲" if v >= 0 else "▼"
    return f"{arrow}{abs(v):.1f}%"


def build_message(index_rows, stock_rows):
    """組三段式純文字訊息。
    index_rows: list of (name, last, daily%)
    stock_rows: list of dict(name, code, group, market, rets={週/月/年})
    """
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    lines = [f"📊 美股／台股監控 {today.month}/{today.day:02d}（{wd}）", ""]

    # 第一段：大盤現況
    lines.append("【大盤現況】")
    for name, last, dchg in index_rows:
        if last is None:
            lines.append(f"{name} 資料暫缺")
        else:
            lines.append(f"{name} {last:,.0f} {fmt_pct(dchg)}")
    lines.append("")

    # 第二段：異動警示（漲跌都報，下跌特別標示）
    lines.append("【異動警示】")
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
            tag = "（逢低）" if v < 0 else ""
            alerts.append((abs(v), f"{sym} {st['code']} {pname}{fmt_pct(v)}{tag}"))
    if alerts:
        alerts.sort(key=lambda x: x[0], reverse=True)
        for _, line in alerts:
            lines.append(line)
    else:
        lines.append("今日無觸發警示門檻")
    lines.append("")

    # 第三段：個股數據（週/月/年），美股與台股分段
    lines.append("【個股數據】週/月/年")
    for market_label in ("美股", "台股"):
        subset = [s for s in stock_rows if s["market"] == market_label]
        if not subset:
            continue
        lines.append(f"—{market_label}—")
        for st in subset:
            r = st["rets"]
            top = max((abs(v) for v in r.values() if v is not None), default=0)
            sym = alert_symbol(top if top >= 10 else None)
            lines.append(
                f"{st['code']} {fmt_pct(r.get('週'))} {fmt_pct(r.get('月'))} {fmt_pct(r.get('年'))}{(' ' + sym) if sym else ''}"
            )
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

def gather(force_dry_run=False):
    watch = load_watchlist()
    index_syms = [s for s, _ in INDICES]
    stock_syms = [w["yf_symbol"] for w in watch]
    closes = fetch_data(index_syms + stock_syms)

    index_rows = []
    for sym, name in INDICES:
        s = closes.get(sym)
        if s is None:
            index_rows.append((name, None, None))
        else:
            index_rows.append((name, float(s.iloc[-1]), daily_change(s)))

    stock_rows = []
    for w in watch:
        s = closes.get(w["yf_symbol"])
        if s is None:
            continue
        stock_rows.append({
            "name": w["name"], "code": w["code"], "group": w["group"],
            "market": w["market"] or ("台股" if w["yf_symbol"].endswith(".TW") else "美股"),
            "rets": compute_period_returns(s),
        })
    return build_message(index_rows, stock_rows)


def main():
    force_dry_run = "--dry-run" in sys.argv
    last_err = None
    for attempt, wait in enumerate([5, 10, 20, 0], start=1):
        try:
            text = gather(force_dry_run=force_dry_run)
            push_line(text, force_dry_run=force_dry_run)
            return 0
        except Exception as e:
            last_err = e
            print(f"[重試 {attempt}] 失敗：{e}", file=sys.stderr)
            if wait:
                time.sleep(wait)
    # 全部重試失敗 → 主動推播異常通知（避免無聲失敗）
    msg = f"⚠️ 今日數據抓取異常，稍後重試。\n（{type(last_err).__name__}: {last_err}）"
    try:
        push_line(msg, force_dry_run=force_dry_run)
    except Exception as e:
        print(f"[嚴重] 連異常通知都推播失敗：{e}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
