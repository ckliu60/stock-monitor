#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (V10.0 落底訊號雙端同步、ETF純數字代碼與三日總數對照版)
---------------------------------------------------------------------------
更新日誌 (V10.0):
1. [新增落底訊號 (10MA/月線)] 嚴格限定僅針對「台大盤 (^TWII)、0050、台積電 (2330)、聯發科 (2454)」四檔計算（美股與其他標的顯示 —）；當 10 日線穿過月線時，於 HTML 與 LINE 卡片同步顯示「🟢 10日穿月線 (落底反轉)」。
2. [台股 ETF 純數字代碼顯示] 凡 00 開頭之台股 ETF（如 0050、0052、00878、00919、00929），於 HTML 網頁與 LINE 卡片一律強制顯示純數字代碼，不顯示中文字。
3. [四-1 與 四-2 三日總數對照] 於提示列直接並列顯示「今 / 昨 / 前」三日之融資餘額、融券餘額、外資期貨淨空口數與借券賣餘額總數。
"""

import os
import sys
import io
import re
import csv
import time
import json
import traceback
import datetime as dt
import xml.etree.ElementTree as ET
from pathlib import Path

# 清除可能導致 curl_cffi / requests 報錯的空字串憑證環境變數
if os.environ.get("CURL_CA_BUNDLE") == "":
    os.environ.pop("CURL_CA_BUNDLE", None)
if os.environ.get("REQUESTS_CA_BUNDLE") == "":
    os.environ.pop("REQUESTS_CA_BUNDLE", None)

import urllib3
import requests
from requests.adapters import HTTPAdapter
import pandas as pd
import yfinance as yf

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 建立共用連線池以加速證交所 / 期交所多日資料查詢
SESSION = requests.Session()
_adapter = HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=1)
SESSION.mount("https://", _adapter)
SESSION.mount("http://", _adapter)

# ---------------------------------------------------------------------------
# 設定區
# ---------------------------------------------------------------------------

INDICES = [
    ("^DJI", "道瓊"),
    ("^IXIC", "那指"),
    ("^GSPC", "標普"),
    ("^SOX", "費半"),
]

# 嚴格限定僅此 4 檔計算「落底訊號 (10MA/月線)」（美股與其他標的皆不適用）
BOTTOM_SIGNAL_TARGETS = {"^TWII", "0050.TW", "2330.TW", "2454.TW"}

DEFAULT_WATCHLIST = [
    {"market": "台股", "code": "2330", "yf_symbol": "2330.TW", "name": "台積", "group": "核心"},
    {"market": "台股", "code": "2454", "yf_symbol": "2454.TW", "name": "聯發科", "group": "核心"},
    {"market": "台股", "code": "0050", "yf_symbol": "0050.TW", "name": "0050", "group": "核心"},
    {"market": "台股", "code": "0052", "yf_symbol": "0052.TW", "name": "0052", "group": "觀察"},
    {"market": "台股", "code": "00919", "yf_symbol": "00919.TW", "name": "00919", "group": "觀察"},
    {"market": "台股", "code": "00929", "yf_symbol": "00929.TW", "name": "00929", "group": "觀察"},
    {"market": "台股", "code": "1303", "yf_symbol": "1303.TW", "name": "南亞", "group": "觀察"},
    {"market": "美股", "code": "TSM", "yf_symbol": "TSM", "name": "台積電R", "group": "核心"},
    {"market": "美股", "code": "TSLA", "yf_symbol": "TSLA", "name": "特斯拉", "group": "核心"},
    {"market": "美股", "code": "SPCX", "yf_symbol": "SPCX", "name": "SPCX", "group": "核心"},
    {"market": "美股", "code": "AMD", "yf_symbol": "AMD", "name": "超微", "group": "核心"},
    {"market": "美股", "code": "GOOG", "yf_symbol": "GOOG", "name": "Alphabet", "group": "核心"},
    {"market": "美股", "code": "MSFT", "yf_symbol": "MSFT", "name": "微軟", "group": "核心"},
    {"market": "美股", "code": "INTC", "yf_symbol": "INTC", "name": "英特爾", "group": "觀察"},
    {"market": "美股", "code": "MRVL", "yf_symbol": "MRVL", "name": "邁威爾", "group": "觀察"},
    {"market": "美股", "code": "PLTR", "yf_symbol": "PLTR", "name": "Palantir", "group": "觀察"},
    {"market": "美股", "code": "META", "yf_symbol": "META", "name": "Meta", "group": "觀察"},
    {"market": "美股", "code": "NOK", "yf_symbol": "NOK", "name": "諾基亞", "group": "觀察"},
    {"market": "美股", "code": "BTC-USD", "yf_symbol": "BTC-USD", "name": "比特幣", "group": "觀察"},
]

MA_WINDOWS = [("MA5", 5), ("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]
LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQbBDIaVplXcSs4Mgn0evbStpMGv3EhI3234VczcBEwD3RuuxoHA8QV0q1ZrM5zzOcFWJ9UKmCuotYR/pub?gid=855107288&single=true&output=csv"

HERE = Path(__file__).resolve().parent
PUBLIC_DIR = HERE / "public"
REPORT_HTML_PATH = PUBLIC_DIR / "index.html"
REPORT_EXCEL_PATH = PUBLIC_DIR / "report_data.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache"
}

# ---------------------------------------------------------------------------
# 工具函式與雲端清單讀取
# ---------------------------------------------------------------------------

def is_tw_etf_symbol(sym):
    """判斷是否為台股 ETF（代碼以 00 開頭，如 0050、0052、00878、00919、00929、00631L）"""
    if not sym:
        return False
    base = str(sym).split(".")[0].strip().upper()
    return base.startswith("00") and len(base) >= 4

def shorten_name(name, code="", market=""):
    """台股 ETF 一律強制顯示純數字代碼（如 00878、00919、0052），其餘依規則精簡"""
    clean_code = str(code).split(".")[0].strip().upper()
    if market == "台股" and is_tw_etf_symbol(clean_code):
        return clean_code
    n = str(name or "").strip()
    if n == "台積電ADR": return "台積電R"
    if n == "台積電": return "台積"
    if n == "台股大盤": return "台大盤"
    return n

def load_watchlist():
    bust_url = f"{SHEET_CSV_URL}&_t={int(time.time())}"
    cloud_rows = []
    seen = set()
    try:
        resp = SESSION.get(bust_url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        for raw in csv.DictReader(io.StringIO(resp.text)):
            code = (raw.get("code") or "").strip().upper()
            yf_symbol = (raw.get("yf_symbol") or "").strip().upper()
            market = (raw.get("market") or "").strip()
            group = (raw.get("長線燈號") or raw.get("group") or "觀察").strip()
            
            if not code and not yf_symbol: continue
            if code in ["^DJI", "^IXIC", "^GSPC", "^SOX", "^TWII"] or group == "大盤": continue
            if market != "台股": market = "美股"

            if market == "台股" and yf_symbol.startswith("00"):
                code = yf_symbol.split(".")[0]
            elif market == "台股" and code.isdigit() and len(code) in (2, 3):
                code = code.zfill(4)

            if not yf_symbol:
                yf_symbol = f"{code}.TW" if market == "台股" else code.replace(".", "-")
            else:
                if market == "美股": yf_symbol = yf_symbol.replace(".", "-")

            if yf_symbol in seen: continue
            seen.add(yf_symbol)

            original_name = (raw.get("name") or yf_symbol).strip()
            display_name = shorten_name(original_name, code=code, market=market)
            cloud_rows.append({
                "market": market, "code": code or yf_symbol, "yf_symbol": yf_symbol,
                "name": display_name, "group": group
            })
    except Exception as e:
        print(f"[提示] 讀取雲端試算表異常: {e}", file=sys.stderr)

    if cloud_rows: return cloud_rows
    out_def = []
    for item in DEFAULT_WATCHLIST:
        cp = item.copy()
        cp["name"] = shorten_name(cp["name"], code=cp["code"], market=cp["market"])
        out_def.append(cp)
    return out_def

# ---------------------------------------------------------------------------
# 台股與 ETF 精準多源採集 & 台指夜盤官方直抓獨立函式
# ---------------------------------------------------------------------------

def fetch_yahoo_chart_quotes(tw_symbols):
    quotes = {}
    if not tw_symbols:
        return quotes

    tw_tz = dt.timezone(dt.timedelta(hours=8))
    hosts = ["https://query1.finance.yahoo.com", "https://query2.finance.yahoo.com"]

    for sym in tw_symbols:
        fetched = False
        for host in hosts:
            if fetched:
                break
            try:
                url = f"{host}/v8/finance/chart/{sym}"
                params = {"interval": "1d", "range": "5d"}
                resp = SESSION.get(url, params=params, headers=HEADERS, timeout=10)
                if resp.status_code != 200:
                    continue
                j = resp.json()
                res_list = (j.get("chart") or {}).get("result") or []
                if not res_list:
                    continue
                res = res_list[0]
                meta = res.get("meta") or {}
                timestamps = res.get("timestamp") or []
                q_block = ((res.get("indicators") or {}).get("quote") or [{}])[0]
                closes_arr = q_block.get("close") or []
                highs_arr = q_block.get("high") or []
                lows_arr = q_block.get("low") or []

                daily_map = {}
                for ts, c_v, h_v, l_v in zip(timestamps, closes_arr, highs_arr, lows_arr):
                    if ts is None or c_v is None:
                        continue
                    try:
                        c_f = float(c_v)
                        if c_f <= 0:
                            continue
                        d_str = dt.datetime.fromtimestamp(int(ts), tw_tz).strftime("%Y-%m-%d")
                        h_f = float(h_v) if (h_v is not None and float(h_v) > 0) else c_f
                        l_f = float(l_v) if (l_v is not None and float(l_v) > 0) else c_f
                        daily_map[d_str] = {"close": c_f, "high": h_f, "low": l_f}
                    except Exception:
                        continue

                rm_p = meta.get("regularMarketPrice")
                rm_t = meta.get("regularMarketTime")
                if rm_p is not None and rm_t is not None:
                    try:
                        rm_pf = float(rm_p)
                        rm_ds = dt.datetime.fromtimestamp(int(rm_t), tw_tz).strftime("%Y-%m-%d")
                        if rm_pf > 0:
                            if rm_ds in daily_map:
                                daily_map[rm_ds]["close"] = rm_pf
                                daily_map[rm_ds]["high"] = max(daily_map[rm_ds]["high"], rm_pf)
                                daily_map[rm_ds]["low"] = min(daily_map[rm_ds]["low"], rm_pf)
                            else:
                                daily_map[rm_ds] = {"close": rm_pf, "high": rm_pf, "low": rm_pf}
                    except Exception:
                        pass

                sorted_dates = sorted(daily_map.keys())
                if not sorted_dates:
                    continue

                latest_d = sorted_dates[-1]
                latest_bar = daily_map[latest_d]
                latest_p = latest_bar["close"]

                prev_p = None
                if len(sorted_dates) >= 2:
                    prev_p = daily_map[sorted_dates[-2]]["close"]
                else:
                    for k in ("chartPreviousClose", "previousClose", "regularMarketPreviousClose"):
                        if meta.get(k) is not None:
                            try:
                                pv = float(meta[k])
                                if pv > 0:
                                    prev_p = pv
                                    break
                            except Exception:
                                pass

                quotes[sym] = {
                    "price": latest_p,
                    "high": latest_bar["high"],
                    "low": latest_bar["low"],
                    "prev_close": prev_p,
                    "date": latest_d,
                    "source": "YahooChartAPI"
                }
                fetched = True
            except Exception:
                continue
        time.sleep(0.04)

    return quotes

def fetch_tw_yahoo_quotes(tw_symbols):
    quotes = {}
    if not tw_symbols: return quotes
    try:
        url = "https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.stockList"
        params = {
            "fields": "symbol,price,regularMarketPrice,regularMarketPreviousClose,regularMarketTime",
            "symbols": ",".join(tw_symbols)
        }
        resp = SESSION.get(url, params=params, headers=HEADERS, timeout=12)
        if resp.status_code == 200:
            data = resp.json()
            items = data if isinstance(data, list) else data.get("data", [])
            tw_tz = dt.timezone(dt.timedelta(hours=8))
            for item in items:
                sym = item.get("symbol")
                rm_raw = item.get("regularMarketPrice")
                pr_raw = item.get("price")
                price_raw = rm_raw if rm_raw not in (None, "-", "", "0", 0) else pr_raw
                prev_raw = item.get("regularMarketPreviousClose")
                mkt_time = item.get("regularMarketTime")
                if sym and price_raw not in (None, "-", ""):
                    try:
                        price = float(str(price_raw).replace(",", ""))
                    except Exception:
                        continue
                    prev = None
                    if prev_raw not in (None, "-", ""):
                        try: prev = float(str(prev_raw).replace(",", ""))
                        except Exception: pass

                    date_str = None
                    if mkt_time:
                        try:
                            s_mt = str(mkt_time).strip()
                            if s_mt.isdigit():
                                ts = int(s_mt)
                                if ts > 1e11: ts //= 1000
                                date_str = dt.datetime.fromtimestamp(ts, tw_tz).strftime("%Y-%m-%d")
                            elif len(s_mt) >= 10 and s_mt[4] == "-":
                                date_str = s_mt[:10]
                        except Exception:
                            pass

                    if price > 0 and date_str:
                        quotes[sym] = {
                            "price": price,
                            "prev_close": prev,
                            "date": date_str,
                            "source": "台灣奇摩API"
                        }
    except Exception as e:
        print(f"[稽核提示] 台灣奇摩備援略過: {e}")
    return quotes

def fetch_twse_official_quotes(tw_symbols=None):
    twse_map = {}
    latest_tw_date = None
    try:
        r = SESSION.get("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", headers=HEADERS, timeout=12, verify=False)
        if r.status_code == 200:
            for item in r.json():
                code = (item.get("Code") or "").strip()
                close_s = (item.get("ClosingPrice") or "").replace(",", "").strip()
                date_roc = (item.get("Date") or "").strip()
                if code and close_s:
                    try:
                        p = float(close_s)
                        d_str = None
                        if len(date_roc) == 7 and date_roc.isdigit():
                            y = int(date_roc[:3]) + 1911
                            d_str = f"{y:04d}-{date_roc[3:5]}-{date_roc[5:7]}"
                            if not latest_tw_date or d_str > latest_tw_date:
                                latest_tw_date = d_str
                        if p > 0 and d_str:
                            twse_map[f"{code}.TW"] = {"price": p, "prev_close": None, "date": d_str, "source": "證交所TWSE"}
                    except Exception:
                        pass
    except Exception as e:
        print(f"[證交所提示] STOCK_DAY_ALL 略過: {e}")

    if tw_symbols:
        tw_tz = dt.timezone(dt.timedelta(hours=8))
        ym_str = dt.datetime.now(tw_tz).strftime("%Y%m") + "01"
        for sym in tw_symbols:
            if not is_tw_etf_symbol(sym) or sym in twse_map:
                continue
            code = sym.split(".")[0].strip()
            try:
                url_etf = f"https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={ym_str}&stockNo={code}&response=json"
                r_etf = SESSION.get(url_etf, headers=HEADERS, timeout=8, verify=False)
                if r_etf.status_code == 200:
                    j_etf = r_etf.json()
                    rows_etf = j_etf.get("data") or []
                    if rows_etf:
                        last_row = rows_etf[-1]
                        raw_d = str(last_row[0]).strip()
                        parts = raw_d.split("/")
                        d_str = None
                        if len(parts) == 3 and parts[0].isdigit():
                            d_str = f"{int(parts[0]) + 1911:04d}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
                        close_val = float(str(last_row[6]).replace(",", "").strip())
                        prev_val = None
                        if len(rows_etf) >= 2:
                            try:
                                prev_val = float(str(rows_etf[-2][6]).replace(",", "").strip())
                            except Exception:
                                pass
                        if close_val > 0 and d_str:
                            twse_map[sym] = {
                                "price": close_val,
                                "prev_close": prev_val,
                                "date": d_str,
                                "source": "證交所ETF日線"
                            }
                time.sleep(0.10)
            except Exception:
                pass

    try:
        r2 = SESSION.get("https://openapi.twse.com.tw/v1/exchangeReport/MI_INDEX", headers=HEADERS, timeout=12, verify=False)
        if r2.status_code == 200:
            for item in r2.json():
                idx_name = item.get("指數", "")
                if "發行量加權股價指數" in idx_name:
                    close_s = (item.get("收盤指數") or "").replace(",", "").strip()
                    chg_s = (item.get("漲跌點數") or "").replace(",", "").strip()
                    dir_s = (item.get("漲跌") or "").strip()
                    date_roc = (item.get("日期") or "").strip()
                    d_str = latest_tw_date
                    if len(date_roc) == 7 and date_roc.isdigit():
                        d_str = f"{int(date_roc[:3]) + 1911:04d}-{date_roc[3:5]}-{date_roc[5:7]}"
                    if close_s and d_str:
                        p = float(close_s)
                        prev_p = None
                        try:
                            if chg_s:
                                chg_val = float(chg_s)
                                if "-" in dir_s or "跌" in dir_s: chg_val = -abs(chg_val)
                                prev_p = round(p - chg_val, 2)
                        except Exception:
                            pass
                        if p > 0:
                            twse_map["^TWII"] = {"price": p, "prev_close": prev_p, "date": d_str, "source": "證交所TWSE"}
                    break
    except Exception:
        pass
    return twse_map

def fetch_taifex_night_quote(day_close_ref=None):
    """
    【獨立函式】確保早上 06:30 穩定抓取昨夜（清晨 05:00 收盤）台指期夜盤數據
    """
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)
    date_label = now_tw.strftime("%m-%d")

    def _to_float(val):
        if val in (None, "-", "", "NaN", "--"): return None
        try: return float(str(val).replace(",", "").replace("%", "").strip())
        except Exception: return None

    # 來源 1：台灣期交所官方夜盤專用頻道 (MarketType="1" 盤後交易時段, CID="TXF" 臺股期貨)
    try:
        mis_url = "https://mis.taifex.com.tw/futures/api/getQuoteList"
        mis_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Content-Type": "application/json;charset=UTF-8",
            "Referer": "https://mis.taifex.com.tw/futures/AfterHoursSession/EquityIndices/FuturesDomestic/",
            "Origin": "https://mis.taifex.com.tw"
        }
        payload = {
            "MarketType": "1",
            "SymbolType": "F",
            "KindID": "1",
            "CID": "TXF",
            "ExpireMonth": "",
            "RowSize": "全部",
            "PageNo": "",
            "SortColumn": "",
            "AscDesc": "A"
        }
        r_mis = SESSION.post(mis_url, json=payload, headers=mis_headers, timeout=12, verify=False)
        if r_mis.status_code == 200:
            j_mis = r_mis.json()
            quote_list = (j_mis.get("RtData") or {}).get("QuoteList") or []
            for q in quote_list:
                sym_id = str(q.get("SymbolID") or "").strip()
                disp_name = str(q.get("DispCName") or "").strip()
                if not sym_id or "-S" in sym_id or "/" in sym_id or "現貨" in disp_name or "價差" in disp_name:
                    continue

                last_p = _to_float(q.get("CLastPrice"))
                ref_p = _to_float(q.get("CReferencePrice"))
                diff_v = _to_float(q.get("CDiff"))
                pct_v = _to_float(q.get("CDiffRate"))
                c_date = str(q.get("CDate") or "").strip()
                if len(c_date) == 8 and c_date.isdigit():
                    date_label = f"{c_date[4:6]}-{c_date[6:8]}"

                if last_p and last_p > 0:
                    base = ref_p if (ref_p and ref_p > 0) else (day_close_ref if (day_close_ref and day_close_ref > 0) else last_p)
                    if diff_v is None:
                        diff_v = (last_p - base) if base != last_p else 0.0
                    if pct_v is None:
                        pct_v = (diff_v / base * 100.0) if base > 0 else 0.0
                    print(f"[台指夜盤] 期交所夜盤 MIS ({sym_id}) 取得: {last_p} ({diff_v:+.0f}, {pct_v:+.2f}%)")
                    return {"price": last_p, "diff": diff_v, "pct": pct_v, "date": date_label, "source": "期交所夜盤MIS"}
                break
    except Exception as e:
        print(f"[夜盤提示] 期交所 MIS 查詢略過: {e}")

    # 來源 2：直接讀取台灣奇摩期貨網頁 (https://tw.stock.yahoo.com/future/WTX%26)
    def _extract_yahoo_num(txt_block, field_names):
        for fn in field_names:
            m1 = re.search(rf'"{fn}"\s*:\s*\{{\s*"raw"\s*:\s*([-0-9.]+)', txt_block)
            if m1:
                v = _to_float(m1.group(1))
                if v is not None:
                    return v
            m2 = re.search(rf'"{fn}"\s*:\s*"?([-0-9.,]+)"?', txt_block)
            if m2:
                v = _to_float(m2.group(1))
                if v is not None:
                    return v
        return None

    try:
        r_web = SESSION.get("https://tw.stock.yahoo.com/future/WTX%26", headers=HEADERS, timeout=12)
        if r_web.status_code == 200:
            txt = r_web.text
            idx_wtx = txt.find('"WTX&"')
            sub_txt = txt[max(0, idx_wtx - 600): idx_wtx + 4000] if idx_wtx != -1 else txt

            p_cur = _extract_yahoo_num(sub_txt, ["regularMarketPrice", "price"])
            p_prev = _extract_yahoo_num(sub_txt, ["regularMarketPreviousClose", "previousClose"])
            chg_cur = _extract_yahoo_num(sub_txt, ["regularMarketChange", "change"])

            if p_cur and p_cur > 0:
                if chg_cur is not None:
                    diff = chg_cur
                    base = p_cur - diff
                elif p_prev and p_prev > 0 and abs(p_cur - p_prev) > 0.1:
                    base = p_prev
                    diff = p_cur - base
                elif day_close_ref and day_close_ref > 0:
                    base = day_close_ref
                    diff = p_cur - base
                else:
                    diff, base = 0.0, p_cur
                pct_val = (diff / base * 100.0) if base > 0 else 0.0
                print(f"[台指夜盤] 奇摩期貨網頁取得: {p_cur} ({diff:+.0f})")
                return {"price": p_cur, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨網頁"}

            if p_prev and p_prev > 0:
                base = day_close_ref if (day_close_ref and day_close_ref > 0) else p_prev
                diff = (p_prev - base) if (day_close_ref and day_close_ref > 0) else 0.0
                pct_val = (diff / base * 100.0) if base > 0 else 0.0
                print(f"[台指夜盤] 奇摩期貨網頁盤前結算欄取得: {p_prev} ({diff:+.0f})")
                return {"price": p_prev, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨網頁(盤後結算)"}
    except Exception as e:
        print(f"[夜盤提示] 奇摩網頁直抓略過: {e}")

    # 來源 3：台灣奇摩新版期貨 API (WTX&)
    try:
        url = "https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.stockList"
        params = {
            "fields": "symbol,price,change,changePercent,regularMarketPrice,regularMarketChange,regularMarketPreviousClose,regularMarketTime",
            "symbols": "WTX&"
        }
        resp = SESSION.get(url, params=params, headers=HEADERS, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            items = data if isinstance(data, list) else data.get("data", [])
            for item in items:
                p_cur = _to_float(item.get("price")) or _to_float(item.get("regularMarketPrice"))
                p_prev = _to_float(item.get("regularMarketPreviousClose"))
                chg_cur = _to_float(item.get("change")) or _to_float(item.get("regularMarketChange"))
                mkt_time = item.get("regularMarketTime")

                if mkt_time:
                    try:
                        if str(mkt_time).isdigit():
                            ts = int(mkt_time)
                            if ts > 1e11: ts //= 1000
                            date_label = dt.datetime.fromtimestamp(ts, tw_tz).strftime("%m-%d")
                        else:
                            date_label = str(mkt_time)[5:10]
                    except Exception:
                        pass

                if p_cur and p_cur > 0:
                    if chg_cur is not None:
                        diff = chg_cur
                        base = p_cur - diff
                    elif p_prev and p_prev > 0:
                        base = p_prev
                        diff = p_cur - base
                    elif day_close_ref and day_close_ref > 0:
                        base = day_close_ref
                        diff = p_cur - base
                    else:
                        diff, base = 0.0, p_cur
                    pct_val = (diff / base * 100.0) if base > 0 else 0.0
                    print(f"[台指夜盤] 奇摩 API 取得: {p_cur} ({diff:+.0f})")
                    return {"price": p_cur, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨API"}

                if (not p_cur or p_cur <= 0) and p_prev and p_prev > 0:
                    base = day_close_ref if (day_close_ref and day_close_ref > 0) else p_prev
                    diff = (p_prev - base) if (day_close_ref and day_close_ref > 0) else 0.0
                    pct_val = (diff / base * 100.0) if base > 0 else 0.0
                    print(f"[台指夜盤] 奇摩 API 盤前結算欄取得: {p_prev} ({diff:+.0f})")
                    return {"price": p_prev, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨API(盤後結算)"}
    except Exception as e:
        print(f"[夜盤提示] 奇摩 API 查詢略過: {e}")

    print("[夜盤警告] 無法取得台指夜盤報價，將顯示暫無盤後。")
    return None

# ---------------------------------------------------------------------------
# 【獨立新模組 V1】三、台美匯率與外資動向 (USD/TWD) [V1]
# ---------------------------------------------------------------------------

def fetch_usdtwd_analysis():
    try:
        df = yf.download("TWD=X", period="2y", interval="1d", auto_adjust=False, progress=False)
        if df.empty:
            return None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.droplevel(1)
        s = df["Close"].dropna()
        s.index = pd.to_datetime(s.index).tz_localize(None).normalize()
        s = s[~s.index.duplicated(keep="last")].sort_index()
        if len(s) < 3:
            return None

        today_val = float(s.iloc[-1])
        today_date = s.index[-1].strftime("%m-%d")

        streak_days = 0
        streak_dir = None
        for idx in range(len(s) - 1, 0, -1):
            curr_v = float(s.iloc[idx])
            prev_v = float(s.iloc[idx - 1])
            if curr_v < prev_v:
                cur_dir = "appreciate"
            elif curr_v > prev_v:
                cur_dir = "depreciate"
            else:
                break
            if streak_dir is None:
                streak_dir = cur_dir
                streak_days = 1
            elif cur_dir == streak_dir:
                streak_days += 1
            else:
                break

        ref_idx = max(0, len(s) - 1 - streak_days)
        ref_start_val = float(s.iloc[ref_idx])
        streak_pct = abs((today_val - ref_start_val) / ref_start_val * 100.0) if ref_start_val > 0 else 0.0

        if streak_dir == "appreciate" and streak_days >= 1:
            tip_text = f"台幣連 {streak_days} 升（較 {streak_days} 日前升 {streak_pct:.2f}%｜外資匯入偏多）"
            tip_color = "#d32f2f"
        elif streak_dir == "depreciate" and streak_days >= 1:
            tip_text = f"台幣連 {streak_days} 貶（較 {streak_days} 日前貶 {streak_pct:.2f}%｜外資匯出偏空）"
            tip_color = "#1b8e3a"
        else:
            tip_text = "台幣匯率今日持平（資金觀望盤整）"
            tip_color = "#666666"

        offsets = [
            ("較昨日", 1),
            ("較前日", 2),
            ("較 1 週前", 5),
            ("較 1 月前", 21),
            ("較 1 季前", 63),
            ("較半年前", 126),
            ("較 1 年前", 252),
        ]

        cols = []
        for label, off in offsets:
            if len(s) > off:
                past_val = float(s.iloc[-1 - off])
                twd_chg_pct = (past_val - today_val) / past_val * 100.0 if past_val > 0 else 0.0
                if twd_chg_pct > 0.005:
                    dir_str = f"升 {abs(twd_chg_pct):.2f}%"
                    flow_str = "匯入"
                    mkt_str = "易漲"
                    cls = "up"
                elif twd_chg_pct < -0.005:
                    dir_str = f"貶 {abs(twd_chg_pct):.2f}%"
                    flow_str = "匯出"
                    mkt_str = "易跌"
                    cls = "down"
                else:
                    dir_str = "持平 0.00%"
                    flow_str = "持平"
                    mkt_str = "盤整"
                    cls = "mute"
                cols.append({
                    "label": label,
                    "rate": f"{past_val:.3f}",
                    "twd_chg": dir_str,
                    "flow": flow_str,
                    "mkt": mkt_str,
                    "cls": cls
                })
            else:
                cols.append({
                    "label": label, "rate": "—", "twd_chg": "—", "flow": "—", "mkt": "—", "cls": "mute"
                })

        return {
            "today_rate": f"{today_val:.3f}",
            "today_date": today_date,
            "tip_text": tip_text,
            "tip_color": tip_color,
            "cols": cols
        }
    except Exception as e:
        print(f"[匯率模組提示] 抓取 TWD=X 異常: {e}")
        return None

# ---------------------------------------------------------------------------
# 【獨立新模組 V1】二、外資與投信買賣超（單位：億元）[V1]
# ---------------------------------------------------------------------------

def fetch_institutional_flow():
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)

    daily_records = []
    seen_dates = set()

    def _parse_num(s):
        if s is None: return 0.0
        try: return float(str(s).replace(",", "").strip())
        except Exception: return 0.0

    def _norm_date_yyyymmdd(raw_d):
        s = str(raw_d or "").strip().replace("/", "").replace("-", "")
        if len(s) == 8 and s.isdigit():
            return s
        if len(s) == 7 and s.isdigit():
            return f"{int(s[:3]) + 1911:04d}{s[3:]}"
        return ""

    def _classify_inst(unit_name):
        n = str(unit_name or "").strip()
        if "外資" in n:
            if "外資自營商" in n and "不含" not in n:
                return None
            return "foreign"
        if "投信" in n:
            return "trust"
        return None

    try:
        r_open = SESSION.get("https://openapi.twse.com.tw/v1/fund/BFI82U", headers=HEADERS, timeout=10, verify=False)
        if r_open.status_code == 200:
            j_list = r_open.json()
            if isinstance(j_list, list) and len(j_list) > 0:
                f_net, t_net, d_str = 0.0, 0.0, ""
                for item in j_list:
                    name = item.get("單位名稱") or item.get("Name") or ""
                    diff_raw = item.get("買賣差額") or item.get("Difference") or "0"
                    raw_d = item.get("日期") or item.get("Date") or ""
                    norm_d = _norm_date_yyyymmdd(raw_d)
                    if norm_d:
                        d_str = norm_d
                    role = _classify_inst(name)
                    if role == "foreign":
                        f_net += _parse_num(diff_raw) / 1e8
                    elif role == "trust":
                        t_net += _parse_num(diff_raw) / 1e8
                if d_str and d_str not in seen_dates and (f_net != 0.0 or t_net != 0.0):
                    seen_dates.add(d_str)
                    daily_records.append({"date": d_str, "foreign": round(f_net, 1), "trust": round(t_net, 1)})
    except Exception as e:
        print(f"[法人 OpenAPI 提示] {e}")

    for d_off in range(0, 96):
        if len(daily_records) >= 60:
            break
        target_dt = now_tw - dt.timedelta(days=d_off)
        if target_dt.weekday() >= 5:
            continue
        q_date = target_dt.strftime("%Y%m%d")
        if q_date in seen_dates:
            continue
        try:
            url = f"https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&dayDate={q_date}&type=day"
            resp = SESSION.get(url, headers=HEADERS, timeout=8, verify=False)
            if resp.status_code == 200:
                j = resp.json()
                if j.get("stat") == "OK" and j.get("data"):
                    f_net, t_net = 0.0, 0.0
                    for row in j["data"]:
                        if len(row) >= 4:
                            role = _classify_inst(row[0])
                            diff_val = _parse_num(row[3]) / 1e8
                            if role == "foreign":
                                f_net += diff_val
                            elif role == "trust":
                                t_net += diff_val
                    seen_dates.add(q_date)
                    daily_records.append({"date": q_date, "foreign": round(f_net, 1), "trust": round(t_net, 1)})
            time.sleep(0.08)
        except Exception:
            continue

    daily_records.sort(key=lambda x: x["date"], reverse=True)

    if not daily_records:
        return None

    f_streak = 0
    f_dir = None
    f_streak_sum = 0.0
    for rec in daily_records:
        val = rec["foreign"]
        if val > 0:
            cur_d = "buy"
        elif val < 0:
            cur_d = "sell"
        else:
            break
        if f_dir is None:
            f_dir = cur_d
            f_streak = 1
            f_streak_sum = val
        elif cur_d == f_dir:
            f_streak += 1
            f_streak_sum += val
        else:
            break

    if f_dir == "buy" and f_streak >= 1:
        tip_text = f"外資連 {f_streak} 買（累計 +{abs(f_streak_sum):,.1f} 億｜籌碼偏多）"
        tip_color = "#d32f2f"
    elif f_dir == "sell" and f_streak >= 1:
        tip_text = f"外資連 {f_streak} 賣（累計 -{abs(f_streak_sum):,.1f} 億｜籌碼偏空）"
        tip_color = "#1b8e3a"
    else:
        tip_text = "外資今日買賣持平（籌碼觀望）"
        tip_color = "#666666"

    windows = [
        ("今日", 0, 1),
        ("昨日", 1, 2),
        ("前日", 2, 3),
        ("近 1 週", 0, 5),
        ("近 1 月", 0, 20),
        ("近 1 季", 0, 60),
    ]

    cols = []
    for label, start_i, end_i in windows:
        sub = daily_records[start_i:end_i]
        if sub:
            f_sum = round(sum(x["foreign"] for x in sub), 1)
            t_sum = round(sum(x["trust"] for x in sub), 1)
            tot_sum = round(f_sum + t_sum, 1)
            if tot_sum > 0:
                mkt_str, mkt_cls = "易漲", "up"
            elif tot_sum < 0:
                mkt_str, mkt_cls = "易跌", "down"
            else:
                mkt_str, mkt_cls = "盤整", "mute"
            cols.append({
                "label": label,
                "foreign": f_sum,
                "trust": t_sum,
                "total": tot_sum,
                "mkt": mkt_str,
                "mkt_cls": mkt_cls
            })
        else:
            cols.append({
                "label": label, "foreign": None, "trust": None, "total": None, "mkt": "—", "mkt_cls": "mute"
            })

    latest_d = daily_records[0]["date"]
    latest_d_fmt = f"{latest_d[4:6]}-{latest_d[6:8]}" if len(latest_d) == 8 else latest_d

    return {
        "latest_date": latest_d_fmt,
        "tip_text": tip_text,
        "tip_color": tip_color,
        "cols": cols
    }

# ---------------------------------------------------------------------------
# 【獨立新模組 V1】四-1 散戶信用交易 & 四-2 法人空單力道（含三日總數對照）
# ---------------------------------------------------------------------------

def fetch_margin_and_shorts():
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)

    def _pnum(s):
        if s is None: return None
        txt = re.sub(r"<[^>]+>", "", str(s)).replace(",", "").replace("%", "").strip()
        if txt in ("", "-", "--", "NaN"): return None
        try: return float(txt)
        except Exception: return None

    margin_recs = []   # [{"date": "YYYYMMDD", "m_bal": 億, "m_chg": 億, "s_bal": 張, "s_chg": 張}]
    sbl_recs = []      # [{"date": "YYYYMMDD", "sbl_bal": 萬張, "sbl_chg": 萬張}]
    fut_recs = []      # [{"date": "YYYYMMDD", "net_oi": 口}]

    # 1. 抓取證交所信用交易統計 (MI_MARGN, selectType=MS)
    for d_off in range(0, 96):
        if len(margin_recs) >= 60:
            break
        t_dt = now_tw - dt.timedelta(days=d_off)
        if t_dt.weekday() >= 5:
            continue
        q_d = t_dt.strftime("%Y%m%d")
        try:
            url_m = f"https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?response=json&date={q_d}&selectType=MS"
            rm = SESSION.get(url_m, headers=HEADERS, timeout=8, verify=False)
            if rm.status_code == 200:
                jm = rm.json()
                if jm.get("stat") == "OK":
                    rows = []
                    if jm.get("tables") and len(jm["tables"]) > 0:
                        rows = jm["tables"][0].get("data") or []
                    elif jm.get("creditList"):
                        rows = jm["creditList"]
                    elif jm.get("data"):
                        rows = jm["data"]

                    m_bal_yi, m_chg_yi, s_bal_zhang, s_chg_zhang = None, None, None, None
                    for r in rows:
                        if not isinstance(r, list) or len(r) < 6:
                            continue
                        label_0 = str(r[0]).strip()
                        prev_v = _pnum(r[4])
                        curr_v = _pnum(r[5])
                        if prev_v is None or curr_v is None:
                            continue
                        if "融資金額" in label_0:
                            m_bal_yi = round(curr_v / 1e5, 1)
                            m_chg_yi = round((curr_v - prev_v) / 1e5, 1)
                        elif "融券" in label_0 and "金額" not in label_0:
                            s_bal_zhang = int(round(curr_v))
                            s_chg_zhang = int(round(curr_v - prev_v))

                    if m_bal_yi is not None and s_bal_zhang is not None:
                        margin_recs.append({
                            "date": q_d,
                            "m_bal": m_bal_yi,
                            "m_chg": m_chg_yi,
                            "s_bal": s_bal_zhang,
                            "s_chg": s_chg_zhang
                        })
            time.sleep(0.06)
        except Exception:
            continue

    # 2. 抓取證交所借券賣出餘額彙總表 (TWT93U)
    for d_off in range(0, 96):
        if len(sbl_recs) >= 60:
            break
        t_dt = now_tw - dt.timedelta(days=d_off)
        if t_dt.weekday() >= 5:
            continue
        q_d = t_dt.strftime("%Y%m%d")
        try:
            url_s = f"https://www.twse.com.tw/rwd/zh/marginTrading/TWT93U?response=json&date={q_d}"
            rs = SESSION.get(url_s, headers=HEADERS, timeout=8, verify=False)
            if rs.status_code == 200:
                js = rs.json()
                if js.get("stat") == "OK":
                    rows = js.get("data") or []
                    if not rows and js.get("tables"):
                        rows = js["tables"][0].get("data") or []
                    for r in reversed(rows):
                        if not isinstance(r, list) or len(r) < 13:
                            continue
                        c0 = str(r[0]).strip()
                        c1 = str(r[1]).strip()
                        if "合計" in c0 or "合計" in c1:
                            prev_sbl = _pnum(r[8])
                            curr_sbl = _pnum(r[12])
                            if prev_sbl is not None and curr_sbl is not None and curr_sbl > 0:
                                div_unit = 1e7 if curr_sbl > 1e8 else 1e4
                                sbl_recs.append({
                                    "date": q_d,
                                    "sbl_bal": round(curr_sbl / div_unit, 2),
                                    "sbl_chg": round((curr_sbl - prev_sbl) / div_unit, 2)
                                })
                            break
            time.sleep(0.06)
        except Exception:
            continue

    # 3. 抓取期交所三大法人臺股期貨未平倉餘額 (futContractsDate)
    for d_off in range(0, 98):
        if len(fut_recs) >= 61:
            break
        t_dt = now_tw - dt.timedelta(days=d_off)
        if t_dt.weekday() >= 5:
            continue
        q_d_slash = t_dt.strftime("%Y/%m/%d")
        q_d_plain = t_dt.strftime("%Y%m%d")
        try:
            url_f = "https://www.taifex.com.tw/cht/3/futContractsDate"
            params_f = {
                "queryType": "1",
                "goDay": "",
                "doQuery": "1",
                "dateaddcnt": "",
                "queryDate": q_d_slash,
                "commodityId": "TXF"
            }
            rf = SESSION.get(url_f, params=params_f, headers=HEADERS, timeout=8, verify=False)
            if rf.status_code == 200 and "外資" in rf.text:
                tr_list = re.findall(r"<tr[^>]*>(.*?)</tr>", rf.text, flags=re.S | re.I)
                in_txf = False
                for tr_html in tr_list:
                    cells = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", td)).strip()
                             for td in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr_html, flags=re.S | re.I)]
                    if not cells:
                        continue
                    row_txt = " ".join(cells)
                    if "臺股期貨" in row_txt or "台股期貨" in row_txt:
                        in_txf = True
                    if in_txf and ("外資" in row_txt):
                        if len(cells) >= 12:
                            net_oi_val = _pnum(cells[-2])
                            if net_oi_val is not None:
                                fut_recs.append({"date": q_d_plain, "net_oi": int(round(net_oi_val))})
                                break
            time.sleep(0.06)
        except Exception:
            continue

    margin_recs.sort(key=lambda x: x["date"], reverse=True)
    sbl_recs.sort(key=lambda x: x["date"], reverse=True)
    fut_recs.sort(key=lambda x: x["date"], reverse=True)

    windows = [
        ("今日", 0, 1),
        ("昨日", 1, 2),
        ("前日", 2, 3),
        ("近 1 週", 0, 5),
        ("近 1 月", 0, 20),
        ("近 1 季", 0, 60),
    ]

    # 組裝【四-1、散戶信用交易（融資 / 融券） [V1]】
    retail_data = None
    if margin_recs:
        latest_m = margin_recs[0]
        m_streak = 0
        m_dir = None
        for rec in margin_recs:
            c = rec["m_chg"]
            cur_d = "dec" if c < 0 else "inc" if c > 0 else None
            if not cur_d:
                break
            if m_dir is None:
                m_dir = cur_d
                m_streak = 1
            elif cur_d == m_dir:
                m_streak += 1
            else:
                break

        streak_str = f"（連 {m_streak} {'減' if m_dir == 'dec' else '增'}）" if m_streak >= 1 else ""
        m_t0 = f"{margin_recs[0]['m_bal']:,.0f}"
        m_t1 = f"{margin_recs[1]['m_bal']:,.0f}" if len(margin_recs) >= 2 else "—"
        m_t2 = f"{margin_recs[2]['m_bal']:,.0f}" if len(margin_recs) >= 3 else "—"

        s_t0 = f"{margin_recs[0]['s_bal'] / 10000.0:,.1f}"
        s_t1 = f"{margin_recs[1]['s_bal'] / 10000.0:,.1f}" if len(margin_recs) >= 2 else "—"
        s_t2 = f"{margin_recs[2]['s_bal'] / 10000.0:,.1f}" if len(margin_recs) >= 3 else "—"

        tip_retail = f"融資餘額：今 {m_t0} / 昨 {m_t1} / 前 {m_t2} 億{streak_str}｜融券餘額：今 {s_t0} / 昨 {s_t1} / 前 {s_t2} 萬張"
        tip_retail_color = "#d32f2f" if m_dir == "dec" else "#1b8e3a" if m_dir == "inc" else "#1f3b66"

        m_cols = []
        s_cols = []
        for label, st_i, ed_i in windows:
            sub = margin_recs[st_i:ed_i]
            if sub:
                m_sum = round(sum(x["m_chg"] for x in sub), 1)
                s_sum = int(sum(x["s_chg"] for x in sub))

                if m_sum < 0:
                    m_judge, m_jcls = "看多", "up"
                elif m_sum > 0:
                    m_judge, m_jcls = "看空", "down"
                else:
                    m_judge, m_jcls = "持平", "mute"

                if s_sum > 0:
                    s_judge, s_jcls = "看多", "up"
                elif s_sum < 0:
                    s_judge, s_jcls = "看空", "down"
                else:
                    s_judge, s_jcls = "持平", "mute"

                m_cols.append({"label": label, "val": m_sum, "judge": m_judge, "jcls": m_jcls})
                s_cols.append({"label": label, "val": s_sum, "judge": s_judge, "jcls": s_jcls})
            else:
                m_cols.append({"label": label, "val": None, "judge": "—", "jcls": "mute"})
                s_cols.append({"label": label, "val": None, "judge": "—", "jcls": "mute"})

        ld_m = latest_m["date"]
        retail_data = {
            "latest_date": f"{ld_m[4:6]}-{ld_m[6:8]}" if len(ld_m) == 8 else ld_m,
            "tip_text": tip_retail,
            "tip_color": tip_retail_color,
            "m_cols": m_cols,
            "s_cols": s_cols
        }

    # 組裝【四-2、法人空單力道（期貨 / 借券） [V1]】
    inst_short_data = None
    if fut_recs or sbl_recs:
        tip_parts = []
        tip_short_color = "#1f3b66"
        latest_d_short = ""

        f_cols = []
        if fut_recs:
            latest_oi = fut_recs[0]["net_oi"]
            latest_d_short = fut_recs[0]["date"]
            f_streak = 0
            f_sdir = None
            for idx in range(len(fut_recs) - 1):
                diff_day = fut_recs[idx]["net_oi"] - fut_recs[idx + 1]["net_oi"]
                cur_sd = "add_short" if diff_day < 0 else "cover" if diff_day > 0 else None
                if not cur_sd:
                    break
                if f_sdir is None:
                    f_sdir = cur_sd
                    f_streak = 1
                elif cur_sd == f_sdir:
                    f_streak += 1
                else:
                    break

            def _fmt_oi_num(v):
                return f"{abs(v):,}" if v <= 0 else f"淨多{v:,}"

            oi_label = "外資期貨淨空" if latest_oi <= 0 else "外資期貨部位"
            f_t0 = _fmt_oi_num(fut_recs[0]["net_oi"])
            f_t1 = _fmt_oi_num(fut_recs[1]["net_oi"]) if len(fut_recs) >= 2 else "—"
            f_t2 = _fmt_oi_num(fut_recs[2]["net_oi"]) if len(fut_recs) >= 3 else "—"
            st_txt = f"（連 {f_streak} {'加空' if f_sdir == 'add_short' else '回補'}）" if f_streak >= 1 else ""
            tip_parts.append(f"{oi_label}：今 {f_t0} / 昨 {f_t1} / 前 {f_t2} 口{st_txt}")
            if f_sdir == "add_short":
                tip_short_color = "#1b8e3a"
            elif f_sdir == "cover":
                tip_short_color = "#d32f2f"

            for label, st_i, ed_i in windows:
                if len(fut_recs) > ed_i:
                    delta_oi = fut_recs[st_i]["net_oi"] - fut_recs[ed_i]["net_oi"]
                    if delta_oi > 0:
                        vcls, judge, jcls = "up", "看多", "up"
                    elif delta_oi < 0:
                        vcls, judge, jcls = "down", "看空", "down"
                    else:
                        vcls, judge, jcls = "mute", "持平", "mute"
                    f_cols.append({"label": label, "val": delta_oi, "vcls": vcls, "judge": judge, "jcls": jcls})
                else:
                    f_cols.append({"label": label, "val": None, "vcls": "mute", "judge": "—", "jcls": "mute"})
        else:
            for label, _, _ in windows:
                f_cols.append({"label": label, "val": None, "vcls": "mute", "judge": "—", "jcls": "mute"})

        sbl_cols = []
        if sbl_recs:
            if not latest_d_short:
                latest_d_short = sbl_recs[0]["date"]
            sbl_t0 = f"{sbl_recs[0]['sbl_bal']:,.0f}"
            sbl_t1 = f"{sbl_recs[1]['sbl_bal']:,.0f}" if len(sbl_recs) >= 2 else "—"
            sbl_t2 = f"{sbl_recs[2]['sbl_bal']:,.0f}" if len(sbl_recs) >= 3 else "—"
            tip_parts.append(f"借券賣餘額：今 {sbl_t0} / 昨 {sbl_t1} / 前 {sbl_t2} 萬張")
            for label, st_i, ed_i in windows:
                sub = sbl_recs[st_i:ed_i]
                if sub:
                    sbl_sum = round(sum(x["sbl_chg"] for x in sub), 2)
                    if sbl_sum < 0:
                        vcls, judge, jcls = "up", "看多", "up"
                    elif sbl_sum > 0:
                        vcls, judge, jcls = "down", "看空", "down"
                    else:
                        vcls, judge, jcls = "mute", "持平", "mute"
                    sbl_cols.append({"label": label, "val": sbl_sum, "vcls": vcls, "judge": judge, "jcls": jcls})
                else:
                    sbl_cols.append({"label": label, "val": None, "vcls": "mute", "judge": "—", "jcls": "mute"})
        else:
            for label, _, _ in windows:
                sbl_cols.append({"label": label, "val": None, "vcls": "mute", "judge": "—", "jcls": "mute"})

        inst_short_data = {
            "latest_date": f"{latest_d_short[4:6]}-{latest_d_short[6:8]}" if len(latest_d_short) == 8 else latest_d_short,
            "tip_text": "｜".join(tip_parts),
            "tip_color": tip_short_color,
            "f_cols": f_cols,
            "sbl_cols": sbl_cols
        }

    return retail_data, inst_short_data

# ---------------------------------------------------------------------------
# 主抓取與自動稽核引擎（Yahoo Chart API 主力 + 真實交易日期防呆鎖）
# ---------------------------------------------------------------------------

def fetch_and_audit_all(symbols, tw_symbols_set, period="2y"):
    closes, kd_data, audit_logs = {}, {}, []
    corrected_count = 0

    tw_list = list(tw_symbols_set)
    yahoo_chart_map = fetch_yahoo_chart_quotes(tw_list)
    tw_yahoo_map = fetch_tw_yahoo_quotes(tw_list)
    twse_map = fetch_twse_official_quotes(tw_list)

    for sym in symbols:
        try:
            df = yf.download(sym, period=period, interval="1d", auto_adjust=False, progress=False)
            if df.empty and sym.endswith(".TW"):
                df = yf.download(sym[:-3] + ".TWO", period=period, interval="1d", auto_adjust=False, progress=False)
            if df.empty:
                audit_logs.append(f"⚠️ {sym}: 無資料")
                continue
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.droplevel(1)
            df = df[["High", "Low", "Close"]].dropna()
            if df.empty:
                continue
            df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
            df = df[~df.index.duplicated(keep="last")].sort_index()

            yf_last_date = df.index[-1].strftime("%Y-%m-%d")
            yf_last_price = float(df["Close"].iloc[-1])

            alt_info = None
            if sym in tw_symbols_set:
                candidates = [
                    yahoo_chart_map.get(sym),
                    twse_map.get(sym),
                    tw_yahoo_map.get(sym),
                ]
                valid_cands = [c for c in candidates if c and c.get("price", 0) > 0 and c.get("date")]
                if valid_cands:
                    valid_cands.sort(key=lambda x: x["date"], reverse=True)
                    alt_info = valid_cands[0]

            if alt_info and alt_info["price"] > 0 and alt_info.get("date"):
                alt_p = float(alt_info["price"])
                alt_d = str(alt_info["date"]).strip()
                alt_prev = alt_info.get("prev_close")
                alt_h = float(alt_info.get("high") or alt_p)
                alt_l = float(alt_info.get("low") or alt_p)

                if alt_d > yf_last_date:
                    new_ts = pd.Timestamp(alt_d)
                    if alt_prev and alt_prev > 0 and abs(yf_last_price - alt_prev) / alt_prev > 0.001:
                        df.loc[df.index[-1], "Close"] = float(alt_prev)
                    df.loc[new_ts] = {"High": alt_h, "Low": alt_l, "Close": alt_p}
                    corrected_count += 1
                    print(f"[自動校正] {sym} 補入 {alt_info['source']} 最新收盤 ({alt_d}): {alt_p} (前收: {alt_prev})")
                elif alt_d == yf_last_date:
                    if abs(alt_p - yf_last_price) / yf_last_price > 0.0005:
                        df.loc[df.index[-1], "Close"] = alt_p
                        corrected_count += 1
                        print(f"[自動校正] {sym} 更新當日收盤 ({alt_d}): {alt_p}")
                    if alt_prev and alt_prev > 0 and len(df) >= 2:
                        prev_df_p = float(df["Close"].iloc[-2])
                        if abs(prev_df_p - float(alt_prev)) / float(alt_prev) > 0.001:
                            df.loc[df.index[-2], "Close"] = float(alt_prev)

            closes[sym] = df["Close"]
            kd_data[sym] = kd_timeframes(df)
        except Exception as e:
            print(f"[錯誤] {sym}: {e}")

    def is_crypto(s): return "-USD" in s.upper()

    tw_dates = [closes[s].index[-1] for s in tw_symbols_set if s in closes and len(closes[s]) > 0]
    us_equity_dates = [closes[s].index[-1] for s in symbols if (s not in tw_symbols_set) and (not is_crypto(s)) and s in closes and len(closes[s]) > 0]
    max_tw_date = max(tw_dates) if tw_dates else None
    max_us_date = max(us_equity_dates) if us_equity_dates else None

    stale_syms = set()
    for sym, s in closes.items():
        if len(s) == 0 or is_crypto(sym): continue
        last_d = s.index[-1]
        ref_d = max_tw_date if sym in tw_symbols_set else max_us_date
        if ref_d is not None and (ref_d - last_d).days >= 1:
            stale_syms.add(sym)
            audit_logs.append(f"⚠️ {sym} 落後 ({last_d.strftime('%m/%d')})")

    audit_summary = {"corrected": corrected_count, "stale_syms": stale_syms, "logs": audit_logs, "tw_date": max_tw_date.strftime("%Y-%m-%d") if max_tw_date else "—", "us_date": max_us_date.strftime("%Y-%m-%d") if max_us_date else "—"}
    return closes, kd_data, audit_summary

# ---------------------------------------------------------------------------
# 技術指標與輔助（含限定 4 檔之落底訊號計算）
# ---------------------------------------------------------------------------

def pct(new, old): return None if not old else (new / old - 1.0) * 100.0
def compute_ma(series): return {name: round(float(series.rolling(w).mean().iloc[-1]), 2) if len(series) >= w else None for name, w in MA_WINDOWS}
def weekly_ma(series, weeks):
    wk = series.resample("W").last().dropna()
    return round(float(wk.rolling(weeks).mean().iloc[-1]), 2) if len(wk) >= weeks else None

def compute_bottom_signal(series, sym):
    """
    計算「落底訊號 (10MA/月線)」：
    嚴格限定僅針對 台灣大盤 (^TWII)、0050 (0050.TW)、台積電 (2330.TW)、聯發科 (2454.TW) 四檔。
    美股與其他標的皆不適用，回傳 ("—", "mute", False)。
    """
    if sym not in BOTTOM_SIGNAL_TARGETS or series is None or len(series) < 21:
        return ("—", "mute", False)

    ma10_s = series.rolling(10).mean()
    ma20_s = series.rolling(20).mean()
    if pd.isna(ma10_s.iloc[-1]) or pd.isna(ma20_s.iloc[-1]) or pd.isna(ma10_s.iloc[-2]) or pd.isna(ma20_s.iloc[-2]):
        return ("—", "mute", False)

    ma10_now = float(ma10_s.iloc[-1])
    ma20_now = float(ma20_s.iloc[-1])
    ma10_prev = float(ma10_s.iloc[-2])
    ma20_prev = float(ma20_s.iloc[-2])

    if ma10_prev < ma20_prev and ma10_now >= ma20_now:
        return ("🟢 10日穿月線 (落底反轉)", "up", True)
    elif ma10_now >= ma20_now:
        return ("✓ 10MA月線上", "safe", False)
    else:
        return ("⏸️ 10MA月線下", "mute", False)

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
    except Exception: out["週"] = (None, None)
    try: out["月"] = compute_kd(df.resample("ME").agg(agg).dropna())
    except Exception: out["月"] = (None, None)
    return out

def get_weather_line(pct_val):
    if pct_val is None: return ""
    if pct_val >= 2.0: return "☀️"
    elif pct_val > 0: return "🌤️"
    elif pct_val == 0: return "☁️"
    elif pct_val > -2.0: return "🌧️"
    else: return "⛈️"

def fmt_pct(v): return "—" if v is None else f"{'▲' if v>0 else '▼' if v<0 else ''}{abs(v):.1f}%"

def get_short_term_status(last, ma5, ma10):
    if pd.isna(last) or pd.isna(ma5) or pd.isna(ma10): return "—"
    if last > ma5 and last > ma10: return "🔥站雙線"
    if last > ma5 and last < ma10: return "⚡站回五日"
    return "⚠️雙線破"

def fetch_yahoo_news(limit=5):
    url = "https://tw.news.yahoo.com/rss/finance"
    news = []
    try:
        resp = SESSION.get(url, headers=HEADERS, timeout=10)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        for item in root.findall('.//item')[:limit]:
            title = item.find('title')
            link = item.find('link')
            if title is not None and link is not None: news.append({"title": title.text, "link": link.text})
    except Exception:
        pass
    if not news: news.append({"title": "暫時無法取得新聞。", "link": "#"})
    return news

# ---------------------------------------------------------------------------
# 匯出資料
# ---------------------------------------------------------------------------
def export_all_data(closes, watch, kd_data, audit_summary):
    all_targets = [{"market": "美股", "code": s, "name": n, "yf": s, "group": "大盤"} for s, n in INDICES] + \
                  [{"market": "台股", "code": "^TWII", "name": "台大盤", "yf": "^TWII", "group": "大盤"}] + \
                  [{"market": w["market"], "code": w["code"], "name": shorten_name(w["name"], w["code"], w["market"]), "yf": w["yf_symbol"], "group": w["group"]} for w in watch]
    
    rows = []
    stale_set = audit_summary.get("stale_syms", set())
    for i, t in enumerate(all_targets, 1):
        sym = t["yf"]
        s = closes.get(sym)
        if s is None or len(s) < 1: continue
        
        last = float(s.iloc[-1])
        pct_1d = pct(last, s.iloc[-2]) if len(s) >= 2 else None
        net_1d = round(last - float(s.iloc[-2]), 2) if len(s) >= 2 else None
        if net_1d == 0:
            net_1d = 0.0
        pct_1w = pct(last, s.iloc[-6]) if len(s) >= 6 else None
        pct_1m = pct(last, s.iloc[-22]) if len(s) >= 22 else None
        pct_1y = pct(last, s.iloc[-253]) if len(s) >= 253 else None
        
        ma = compute_ma(s)
        ma20w = weekly_ma(s, 20)
        kd = kd_data.get(sym, {})
        kd_d, kd_w, kd_m = kd.get("日", (None, None)), kd.get("週", (None, None)), kd.get("月", (None, None))
        
        audit_flag = "⚠️日期延遲" if sym in stale_set else "✅正常"
        short_status = get_short_term_status(last, ma.get("MA5"), ma.get("MA10"))
        bot_text, bot_cls, bot_trig = compute_bottom_signal(s, sym)
        
        rows.append({
            "序號": i, "長線燈號": t["group"], "短線燈號": short_status,
            "落底訊號": bot_text, "落底樣式": bot_cls, "落底觸發": bot_trig,
            "市場": t["market"], "代碼": t["code"],
            "名稱": t["name"], "yf_symbol": sym, "最新收盤": round(last, 2),
            "資料日期": s.index[-1].strftime("%Y-%m-%d"), "稽核狀態": audit_flag,
            "當日漲跌": net_1d, "當日%": round(pct_1d, 2) if pct_1d is not None else None,
            "週%": round(pct_1w, 2) if pct_1w is not None else None, "月%": round(pct_1m, 2) if pct_1m is not None else None,
            "年%": round(pct_1y, 2) if pct_1y is not None else None,
            "日K": kd_d[0], "日D": kd_d[1], "週K": kd_w[0], "週D": kd_w[1], "月K": kd_m[0], "月D": kd_m[1],
            "MA5": ma.get("MA5"), "MA10": ma.get("MA10"), "MA20": ma.get("MA20"), "MA60": ma.get("MA60"), "MA240": ma.get("MA240"), "MA20週": ma20w
        })
        
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    
    try:
        excel_df = df.drop(columns=["落底樣式", "落底觸發"], errors="ignore")
        with pd.ExcelWriter(REPORT_EXCEL_PATH, engine='xlsxwriter') as writer:
            excel_df.to_excel(writer, sheet_name='股市監控報表', index=False)
            workbook, worksheet = writer.book, writer.sheets['股市監控報表']
            white_bg = workbook.add_format({'bg_color': '#FFFFFF'})
            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'center', 'align': 'center', 'fg_color': '#D7E4BC', 'border': 1})
            for col_num, value in enumerate(excel_df.columns.values): worksheet.write(0, col_num, value, header_format)
            worksheet.set_column('A:F', 12, white_bg)
            worksheet.set_column('G:H', 12, white_bg)  
            worksheet.set_column('I:AA', 10, white_bg)  
            worksheet.freeze_panes(1, 7)
            worksheet.autofilter(0, 0, len(excel_df), len(excel_df.columns) - 1)
            tw_format = workbook.add_format({'bg_color': '#FFF9DB'})
            worksheet.conditional_format(1, 0, len(excel_df), len(excel_df.columns) - 1, {'type': 'formula', 'criteria': '=$E2="台股"', 'format': tw_format})
    except Exception as e:
        print(f"[提示] 略過 Excel 匯出: {e}", file=sys.stderr)
        
    return df

def generate_html_report(df, vix_hist, news_data, audit_summary, night_quote=None, fx_data=None, inst_data=None, retail_data=None, inst_short_data=None):
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_str = dt.datetime.now(tw_tz).strftime("%Y/%m/%d %H:%M")
    
    idx_df = df[df['長線燈號'] == '大盤']
    us_stocks = df[(df['市場'] == '美股') & (df['長線燈號'] != '大盤')]
    tw_stocks = df[(df['市場'] == '台股') & (df['長線燈號'] != '大盤')]
    
    corr_n = audit_summary.get("corrected", 0)
    stale_n = len(audit_summary.get("stale_syms", set()))
    tw_d, us_d = audit_summary.get("tw_date", "—"), audit_summary.get("us_date", "—")
    if stale_n == 0: audit_banner = f"<div style='background:#e8f5e9; color:#1b8e3a; padding:8px 12px; border-radius:6px; font-size:12px; margin:8px 0 14px 0; font-weight:bold;'>✅ 自動稽核通過｜美股基準日：{us_d}｜台股基準日：{tw_d}（多源資料已即時驗證，自動校正 {corr_n} 檔收盤價）</div>"
    else: audit_banner = f"<div style='background:#fff3e0; color:#ef6c00; padding:8px 12px; border-radius:6px; font-size:12px; margin:8px 0 14px 0; font-weight:bold;'>⚠️ 自動稽核提醒｜美股基準日：{us_d}｜台股基準日：{tw_d}｜已校正 {corr_n} 檔，仍有 {stale_n} 檔日期異常：{'；'.join(audit_summary.get('logs', [])[:3])}</div>"

    def fmt_net(val, is_idx=False):
        if pd.isna(val): return '<td class="num mute">—</td>'
        if abs(float(val)) < 1e-6:
            return f'<td class="num mute">{0:,.{0 if is_idx else 2}f}</td>'
        return f'<td class="num {"up" if val>0 else "down" if val<0 else "mute"}">{"+" if val>0 else ""}{val:,.{0 if is_idx else 2}f}</td>'

    def fmt_pct_td(val):
        if pd.isna(val): return '<td class="num mute">—</td>'
        if abs(float(val)) < 0.05: return '<td class="num mute">0.0%</td>'
        return f'<td class="num {"up" if val>0 else "down" if val<0 else "mute"}">{"▲" if val>0 else "▼" if val<0 else ""}{abs(val):.1f}%</td>'

    def get_weather_icon(pct_val, threshold):
        if pd.isna(pct_val): return "—"
        if pct_val >= threshold: return "☀️"
        elif pct_val > 0: return "🌤️"
        elif pct_val == 0: return "☁️"
        elif pct_val <= -threshold: return "⛈️"
        else: return "🌧️"

    def get_position(last, ma20, ma60, ma240):
        if pd.isna(last): return "—"
        if pd.notna(ma20) and last >= ma20: return "月線上"
        if pd.notna(ma60) and last >= ma60: return "月～季"
        if pd.notna(ma240) and last >= ma240: return "季～年"
        if pd.notna(ma240) and last < ma240: return "年線下"
        return "—"

    def get_status_badge(last, ma20w, kd_d, kd_w, kd_m):
        if pd.isna(last) or pd.isna(ma20w) or ma20w == 0: return '<span class="mute">—</span>'
        if last < ma20w: return '<span class="badge b-warn">⚠️ 破線空手</span>'
        kd_list = [v for v in [kd_d, kd_w, kd_m] if pd.notna(v)]
        if any(k < 20 for k in kd_list): return '<span class="badge b-down"><span class="s-dot">●</span> 日K超跌</span>'
        if any(k > 80 for k in kd_list): return '<span class="badge b-up"><span class="s-dot">●</span> 熱K>80</span>'
        return '<span class="badge b-safe">✓ 線上安全</span>'
    
    def get_short_badge(status_text):
        if status_text == "🔥站雙線": return '<span class="badge b-up">🔥站雙線</span>'
        if status_text == "⚡站回五日": return '<span class="badge" style="background:#e3f2fd; color:#1976d2;">⚡站回五日</span>'
        if status_text == "⚠️雙線破": return '<span class="badge b-down">⚠️雙線破</span>'
        return '<span class="mute">—</span>'

    def get_bottom_badge(bot_text, bot_cls):
        if not bot_text or bot_text == "—":
            return '<span class="mute">—</span>'
        if bot_cls == "up":
            return f'<span class="badge b-up">{bot_text}</span>'
        if bot_cls == "safe":
            return f'<span class="badge b-safe">{bot_text}</span>'
        return f'<span class="badge" style="background:#f5f5f5; color:#777;">{bot_text}</span>'
        
    def get_dist_pct(last, ma20w): return "—" if pd.isna(last) or pd.isna(ma20w) or ma20w == 0 else f"{'+' if ((last - ma20w) / ma20w * 100)>0 else ''}{(last - ma20w) / ma20w * 100:.1f}%"

    html = f"""<!doctype html><html lang='zh-Hant'><head><meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no'>
<title>股市監控報表</title>
<style>
:root {{ --up:#d32f2f; --down:#1b8e3a; --line:#e3e6ea; --mute:#8a9099; --head:#1f3b66; }} 
* {{ box-sizing:border-box; }} 
body {{ margin:0; background:#f6f7f9; color:#222; font-family:'Microsoft JhengHei',system-ui,sans-serif; font-size:13px; padding-bottom:40px; overflow-x:hidden; }} 
.wrap {{ max-width:960px; margin:0 auto; padding:16px; overflow-x:hidden; }} 
.table-container {{ width:100%; overflow-x:auto; -webkit-overflow-scrolling:touch; background:#fff; border:1px solid var(--line); border-radius:8px; margin-bottom: 20px; }}
table {{ width:100%; border-collapse:collapse; background:#fff; }} 
th,td {{ padding:8px 10px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }} 
thead th {{ background:#eef1f5; color:#333; font-weight:600; font-size:12px; }} 
thead th.num {{ text-align:center; }} 
tbody tr:last-child td {{ border-bottom:none; }} 
td.num {{ text-align:center; font-variant-numeric:tabular-nums; }} 
td.idx {{ color:var(--mute); text-align:center; }} 
.up {{ color:var(--up); font-weight:bold; }} .down {{ color:var(--down); font-weight:bold; }} .mute {{ color:var(--mute); }} 
tr.grp td {{ background:#f0f3f7; font-weight:700; color:var(--head); position:static; box-shadow:none; }} 
tr.sub-judge td {{ background:#fafbfd; padding-top:6px; padding-bottom:7px; }}
tr.sub-judge td.freeze-idx, tr.sub-judge td.freeze-name {{ background:#fafbfd; }}
a {{ color:#1456b8; text-decoration:none; font-weight:bold; }} a:hover {{ text-decoration:underline; }}
.badge {{ padding: 3px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block; }}
.b-up {{ background: #ffebee; color: #d32f2f; }}
.b-down {{ background: #e8f5e9; color: #1b8e3a; }}
.b-warn {{ background: #fff3e0; color: #ef6c00; }}
.b-safe {{ background: #f0f3f7; color: #333; }}
.attr-tag {{ display:inline-block; background:#f0f3f7; color:#1f3b66; font-size:11px; font-weight:bold; padding:2px 6px; border-radius:4px; border:1px solid #d5dce6; }}
.foot {{ color:var(--mute); font-size:11px; margin-top:14px; text-align:center; }}
.action-btn {{ display:inline-block; padding:10px 20px; border-radius:4px; text-decoration:none; font-size:14px; font-weight:bold; text-align:center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); transition: opacity 0.2s; margin-top: 10px; }}
.action-btn:hover {{ opacity: 0.8; text-decoration:none; }}
.btn-download {{ background:#1b8e3a; color:#fff; width: 100%; max-width: 300px; }}

/* 獨立模組版本標籤與趨勢提示樣式 */
.ver-tag {{ display:inline-block; background:#eef1f5; color:#1f3b66; font-size:11px; font-weight:bold; padding:2px 7px; border-radius:4px; margin-left:6px; vertical-align:2px; border:1px solid #d5dce6; }}
.trend-tip {{ font-size:13px; margin:-6px 0 10px 0; padding:6px 10px; background:#fff; border-left:4px solid #1f3b66; border-radius:4px; border-top:1px solid var(--line); border-right:1px solid var(--line); border-bottom:1px solid var(--line); line-height:1.6; }}

/* 精緻小圓點樣式 */
.s-dot {{ font-size: 0.65em; vertical-align: 1px; margin-right: 1px; display: inline-block; }}

/* HTML 第六區位階警示完全取消粗體 (減壓) */
ol.alert-list {{ line-height:1.8; font-size:13px; font-weight:400 !important; padding-left:24px; margin-bottom:24px; color:#333; }}
ol.alert-list li, ol.alert-list span {{ font-weight:400 !important; }}

.freeze-single {{ position: -webkit-sticky; position: sticky; left: 0; z-index: 2; background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2); }}
.freeze-idx {{ position: -webkit-sticky; position: sticky; left: 0; z-index: 2; background: #fff; min-width: 25px; max-width: 25px; padding-left: 4px; padding-right: 4px; }}
.freeze-name {{ position: -webkit-sticky; position: sticky; left: 25px; z-index: 2; background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2); min-width: 78px; max-width: 105px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding-left: 6px; padding-right: 6px; }}
thead th.freeze-single, thead th.freeze-idx, thead th.freeze-name {{ background: #eef1f5; z-index: 3; }}
tr.tw-row td {{ background-color: #fff9db !important; }}
tr.tw-row td.freeze-single, tr.tw-row td.freeze-idx, tr.tw-row td.freeze-name {{ background-color: #fff9db !important; }}
</style></head><body><div class='wrap'>
<h1>📊 股市核心監控報表</h1><div class='sub'>產生時間：{now_str} (台灣時間)</div>
{audit_banner}

<h2>一、大盤與 VIX 狀況 (+/- 2%)</h2>
<div class='table-container'><table>
<thead><tr>
    <th class='freeze-single'>指數</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>當日燈號</th><th class='num'>股價位階</th><th class='num'>短線燈號</th><th class='num'>落底訊號 (10MA/月線)</th><th class='num'>5MA/10MA</th><th class='num'>長線燈號</th><th class='num'>資料日期</th>
</tr></thead><tbody>
"""
    
    for _, row in idx_df.iterrows():
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma5, ma10, ma20, ma60, ma240, ma20w = row['MA5'], row['MA10'], row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        pos = get_position(last, ma20, ma60, ma240)
        weather = get_weather_icon(pct_val, 2.0)
        short_status_text = get_short_term_status(last, ma5, ma10)
        short_badge = get_short_badge(short_status_text)
        bot_badge = get_bottom_badge(row['落底訊號'], row['落底樣式'])
        ma5_10_str = f"{ma5:,.0f} / {ma10:,.0f}" if pd.notna(ma5) and pd.notna(ma10) else "—"
        status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        d_tag = row['資料日期'][5:] + ("⚠️" if "⚠️" in row['稽核狀態'] else "")
        
        html += f"<tr class='{bg_cls}'><td class='freeze-single'><a href='https://finance.yahoo.com/quote/{row['yf_symbol']}' target='_blank'>{row['名稱']}</a></td>"
        html += f"<td class='num'>{last:,.0f}</td>"
        html += fmt_net(net, is_idx=True) + fmt_pct_td(pct_val)
        html += f"<td class='num'>{weather}</td><td class='num'>{pos}</td><td class='num'>{short_badge}</td><td class='num'>{bot_badge}</td><td class='num'>{ma5_10_str}</td><td class='num'>{status}</td><td class='num mute'>{d_tag}</td></tr>"

        # 緊接在「台大盤」下方插入「台指夜盤」
        if row['yf_symbol'] == '^TWII':
            if night_quote and night_quote.get("price", 0) > 0:
                np = night_quote["price"]
                ndiff = night_quote["diff"]
                npct = night_quote["pct"]
                n_weather = get_weather_icon(npct, 2.0)
                n_date_tag = f"{night_quote.get('date', '')} 盤後"
                html += f"<tr class='tw-row'><td class='freeze-single'><a href='https://tw.stock.yahoo.com/future/WTX%26' target='_blank'>台指夜盤</a></td>"
                html += f"<td class='num'>{np:,.0f}</td>"
                html += fmt_net(ndiff, is_idx=True) + fmt_pct_td(npct)
                html += f"<td class='num'>{n_weather}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num'><span class='badge b-safe'>夜盤收盤</span></td><td class='num mute'>{n_date_tag}</td></tr>"
            else:
                html += f"<tr class='tw-row'><td class='freeze-single'>台指夜盤</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>暫無盤後</td><td class='num mute'>—</td></tr>"

    for v_name, v_sym in [("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")]:
        h = vix_hist.get(v_sym, [0])
        val = h[0]
        if val is None or val <= 0:
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>暫無報價</td><td class='num mute'>—</td></tr>"
        else:
            status = "<span class='badge b-down'><span class='s-dot'>●</span> 安全</span>" if val < 35 else "<span class='badge b-up'><span class='s-dot'>●</span> 恐慌</span>"
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num'>{val:.2f}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num'>{status}</td><td class='num mute'>—</td></tr>"
            
    html += "</tbody></table></div>"

    # -----------------------------------------------------------------------
    # 【獨立模組 V1】二、外資與投信買賣超（單位：億元）[V1]
    # -----------------------------------------------------------------------
    html += "<h2>二、外資與投信買賣超（單位：億元） <span class='ver-tag'>V1</span></h2>"
    if inst_data and inst_data.get("cols"):
        html += f"<div class='trend-tip' style='border-left-color:{inst_data['tip_color']};'>💡 <span style='color:{inst_data['tip_color']}; font-weight:bold;'>{inst_data['tip_text']}</span> <span class='mute' style='font-size:11px;'>(最新交易日：{inst_data['latest_date']})</span></div>"
        html += "<div class='table-container'><table><thead><tr>"
        html += "<th class='num freeze-idx'>項次</th><th class='freeze-name'>項目</th>"
        for c in inst_data["cols"]:
            html += f"<th class='num'>{c['label']}</th>"
        html += "</tr></thead><tbody>"

        def _fmt_inst_td(val, bold=False):
            if val is None: return '<td class="num mute">—</td>'
            cls = "up" if val > 0 else "down" if val < 0 else "mute"
            txt = f"{'+' if val > 0 else ''}{val:,.1f}"
            return f'<td class="num {cls}">{"<b>" + txt + "</b>" if bold else txt}</td>'

        html += "<tr><td class='num idx freeze-idx'>1</td><td class='freeze-name'><b>外資</b></td>"
        for c in inst_data["cols"]:
            html += _fmt_inst_td(c["foreign"])
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>2</td><td class='freeze-name'><b>投信</b></td>"
        for c in inst_data["cols"]:
            html += _fmt_inst_td(c["trust"])
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>3</td><td class='freeze-name'><b>合計</b></td>"
        for c in inst_data["cols"]:
            html += _fmt_inst_td(c["total"], bold=True)
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>4</td><td class='freeze-name'><b>股市漲跌</b></td>"
        for c in inst_data["cols"]:
            html += f"<td class='num {c['mkt_cls']}'>{c['mkt']}</td>"
        html += "</tr>"

        html += "</tbody></table></div>"
    else:
        html += "<div class='trend-tip'>⚠️ 暫時無法取得證交所三大法人買賣超資料。</div>"

    # -----------------------------------------------------------------------
    # 【獨立模組 V1】三、台美匯率與外資動向 (USD/TWD) [V1]
    # -----------------------------------------------------------------------
    html += "<h2>三、台美匯率與外資動向 (USD/TWD) <span class='ver-tag'>V1</span></h2>"
    if fx_data and fx_data.get("cols"):
        html += f"<div class='trend-tip' style='border-left-color:{fx_data['tip_color']};'>💡 <span style='color:{fx_data['tip_color']}; font-weight:bold;'>{fx_data['tip_text']}</span> <span class='mute' style='font-size:11px;'>(基準日：{fx_data['today_date']})</span></div>"
        html += "<div class='table-container'><table><thead><tr>"
        html += "<th class='num freeze-idx'>項次</th><th class='freeze-name'>項目</th><th class='num'>今日基準</th>"
        for c in fx_data["cols"]:
            html += f"<th class='num'>{c['label']}</th>"
        html += "</tr></thead><tbody>"

        html += f"<tr><td class='num idx freeze-idx'>1</td><td class='freeze-name'><b>當時匯率</b></td><td class='num'><b>{fx_data['today_rate']}</b></td>"
        for c in fx_data["cols"]:
            html += f"<td class='num'>{c['rate']}</td>"
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>2</td><td class='freeze-name'><b>台幣升貶</b></td><td class='num mute'>—</td>"
        for c in fx_data["cols"]:
            html += f"<td class='num {c['cls']}'>{c['twd_chg']}</td>"
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>3</td><td class='freeze-name'><b>外資動向</b></td><td class='num mute'>—</td>"
        for c in fx_data["cols"]:
            html += f"<td class='num {c['cls']}'>{c['flow']}</td>"
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>4</td><td class='freeze-name'><b>股市漲跌</b></td><td class='num mute'>—</td>"
        for c in fx_data["cols"]:
            html += f"<td class='num {c['cls']}'>{c['mkt']}</td>"
        html += "</tr>"

        html += "</tbody></table></div>"
    else:
        html += "<div class='trend-tip'>⚠️ 暫時無法取得台美匯率歷史資料。</div>"

    # -----------------------------------------------------------------------
    # 【獨立模組 V1】四-1、散戶信用交易（融資 / 融券）[V1]
    # -----------------------------------------------------------------------
    html += "<h2>四-1、散戶信用交易（融資 / 融券） <span class='ver-tag'>V1</span></h2>"
    if retail_data and retail_data.get("m_cols"):
        html += f"<div class='trend-tip' style='border-left-color:{retail_data['tip_color']};'>💡 <span style='color:{retail_data['tip_color']}; font-weight:bold;'>{retail_data['tip_text']}</span> <span class='mute' style='font-size:11px;'>(最新交易日：{retail_data['latest_date']})</span></div>"
        html += "<div class='table-container'><table><thead><tr>"
        html += "<th class='num freeze-idx'>項次</th><th class='freeze-name'>項目</th><th class='num'>屬性</th>"
        for c in retail_data["m_cols"]:
            html += f"<th class='num'>{c['label']}</th>"
        html += "</tr></thead><tbody>"

        html += "<tr><td class='num idx freeze-idx'>1</td><td class='freeze-name'><b>融資(億)</b></td><td class='num'><span class='attr-tag'>散多標</span></td>"
        for c in retail_data["m_cols"]:
            v = c["val"]
            txt = "—" if v is None else f"{'+' if v > 0 else ''}{v:,.1f}"
            html += f"<td class='num'>{txt}</td>"
        html += "</tr>"

        html += "<tr class='sub-judge'><td class='num idx freeze-idx'>↳</td><td class='freeze-name mute'>判讀</td><td class='num mute'>—</td>"
        for c in retail_data["m_cols"]:
            html += f"<td class='num {c['jcls']}'>{c['judge']}</td>"
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>2</td><td class='freeze-name'><b>融券(張)</b></td><td class='num'><span class='attr-tag'>散空標</span></td>"
        for c in retail_data["s_cols"]:
            v = c["val"]
            txt = "—" if v is None else f"{'+' if v > 0 else ''}{v:,d}"
            html += f"<td class='num'>{txt}</td>"
        html += "</tr>"

        html += "<tr class='sub-judge'><td class='num idx freeze-idx'>↳</td><td class='freeze-name mute'>判讀</td><td class='num mute'>—</td>"
        for c in retail_data["s_cols"]:
            html += f"<td class='num {c['jcls']}'>{c['judge']}</td>"
        html += "</tr>"

        html += "</tbody></table></div>"
    else:
        html += "<div class='trend-tip'>⚠️ 暫時無法取得證交所信用交易（融資／融券）資料。</div>"

    # -----------------------------------------------------------------------
    # 【獨立模組 V1】四-2、法人空單力道（期貨 / 借券）[V1]
    # -----------------------------------------------------------------------
    html += "<h2>四-2、法人空單力道（期貨 / 借券） <span class='ver-tag'>V1</span></h2>"
    if inst_short_data and inst_short_data.get("f_cols"):
        html += f"<div class='trend-tip' style='border-left-color:{inst_short_data['tip_color']};'>💡 <span style='color:{inst_short_data['tip_color']}; font-weight:bold;'>{inst_short_data['tip_text']}</span> <span class='mute' style='font-size:11px;'>(最新交易日：{inst_short_data['latest_date']})</span></div>"
        html += "<div class='table-container'><table><thead><tr>"
        html += "<th class='num freeze-idx'>項次</th><th class='freeze-name'>項目</th><th class='num'>屬性</th>"
        for c in inst_short_data["f_cols"]:
            html += f"<th class='num'>{c['label']}</th>"
        html += "</tr></thead><tbody>"

        html += "<tr><td class='num idx freeze-idx'>1</td><td class='freeze-name'><b>期貨空單(口)</b></td><td class='num'><span class='attr-tag'>外空標</span></td>"
        for c in inst_short_data["f_cols"]:
            v = c["val"]
            txt = "—" if v is None else f"{'+' if v > 0 else ''}{v:,d}"
            html += f"<td class='num {c['vcls']}'>{txt}</td>"
        html += "</tr>"

        html += "<tr class='sub-judge'><td class='num idx freeze-idx'>↳</td><td class='freeze-name mute'>判讀</td><td class='num mute'>—</td>"
        for c in inst_short_data["f_cols"]:
            html += f"<td class='num {c['jcls']}'>{c['judge']}</td>"
        html += "</tr>"

        html += "<tr><td class='num idx freeze-idx'>2</td><td class='freeze-name'><b>借券賣(萬張)</b></td><td class='num'><span class='attr-tag'>外空標</span></td>"
        for c in inst_short_data["sbl_cols"]:
            v = c["val"]
            txt = "—" if v is None else f"{'+' if v > 0 else ''}{v:,.2f}"
            html += f"<td class='num {c['vcls']}'>{txt}</td>"
        html += "</tr>"

        html += "<tr class='sub-judge'><td class='num idx freeze-idx'>↳</td><td class='freeze-name mute'>判讀</td><td class='num mute'>—</td>"
        for c in inst_short_data["sbl_cols"]:
            html += f"<td class='num {c['jcls']}'>{c['judge']}</td>"
        html += "</tr>"

        html += "</tbody></table></div>"
    else:
        html += "<div class='trend-tip'>⚠️ 暫時無法取得期交所外資期貨未平倉或借券賣出資料。</div>"
    
    # -----------------------------------------------------------------------
    # 五、個股追蹤清單 (支援短線/落底/長線三燈號)
    # -----------------------------------------------------------------------
    html += "<h2>五、個股追蹤清單 (支援短線/長線雙燈號)</h2><div class='table-container'><table><thead><tr>"
    html += "<th class='num freeze-idx'>#</th><th class='freeze-name'>名稱</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>當日燈號</th><th class='num'>股價位階</th><th class='num'>短線燈號</th><th class='num'>落底訊號 (10MA/月線)</th><th class='num'>5MA/10MA</th><th class='num'>長線燈號</th><th class='num'>日/週/月 K值</th><th class='num'>MA20週均</th><th class='num'>距20週均(%)</th><th class='num'>資料日期</th></tr></thead><tbody>"

    def render_stock_row(row, idx):
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma5, ma10, ma20, ma60, ma240, ma20w = row['MA5'], row['MA10'], row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        pos = get_position(last, ma20, ma60, ma240)
        weather = get_weather_icon(pct_val, 5.0)
        long_status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        short_status_text = get_short_term_status(last, ma5, ma10)
        short_badge = get_short_badge(short_status_text)
        bot_badge = get_bottom_badge(row['落底訊號'], row['落底樣式'])
        dist_str = get_dist_pct(last, ma20w)
        
        kd_str = f"{int(kd_d) if pd.notna(kd_d) else '-'}/{int(kd_w) if pd.notna(kd_w) else '-'}/{int(kd_m) if pd.notna(kd_m) else '-'}"
        ma20w_str = f"{ma20w:,.2f}" if pd.notna(ma20w) else "—"
        ma5_10_str = f"{ma5:,.2f} / {ma10:,.2f}" if pd.notna(ma5) and pd.notna(ma10) else "—"
        
        url = f"https://finance.yahoo.com/quote/{row['yf_symbol']}"
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        d_tag = row['資料日期'][5:] + ("⚠️" if "⚠️" in row['稽核狀態'] else "")
        
        res = f"<tr class='{bg_cls}'><td class='num idx freeze-idx'>{idx}</td><td class='name freeze-name'><a href='{url}' target='_blank'>{row['名稱']}</a></td><td class='num'>{last:,.2f}</td>"
        res += fmt_net(net, is_idx=False) + fmt_pct_td(pct_val)
        res += f"<td class='num'>{weather}</td><td class='num'>{pos}</td><td class='num'>{short_badge}</td><td class='num'>{bot_badge}</td><td class='num'>{ma5_10_str}</td><td class='num'>{long_status}</td><td class='num'>{kd_str}</td><td class='num'>{ma20w_str}</td><td class='num { 'up' if '+' in dist_str else 'down' }'>{dist_str}</td><td class='num mute'>{d_tag}</td></tr>"
        return res

    html += "<tr class='grp'><td colspan='15'>美股觀察區</td></tr>"
    idx = 1
    for _, row in us_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "<tr class='grp'><td colspan='15'>台股觀察區</td></tr>"
    for _, row in tw_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "</tbody></table></div>"

    # -----------------------------------------------------------------------
    # 六、其他位階警示
    # -----------------------------------------------------------------------
    html += "<h2>六、其他位階警示</h2><ol class='alert-list'>"
    alert_df = df[df['長線燈號'].isin(['大盤', '核心'])]
    alert_count_html = 0
    for _, r in alert_df.iterrows():
        last, ma60, ma240 = r['最新收盤'], r['MA60'], r['MA240']
        drop = []
        if pd.notna(ma60) and last < ma60: drop.append("<span style='color:#1b8e3a; font-weight:400;'><span class='s-dot'>●</span>[跌]破季線</span>")
        if pd.notna(ma240) and last < ma240: drop.append("<span style='color:#d32f2f; font-weight:400;'><span class='s-dot'>●</span>[跌]破年線</span>")
        w_ret, m_ret = r['週%'], r['月%']
        if (pd.notna(w_ret) and w_ret <= -20) or (pd.notna(m_ret) and m_ret <= -20): drop.append("<span style='color:#1b8e3a; font-weight:400;'><span class='s-dot'>●</span>週/月跌>20%</span>")
            
        ko80 = []
        if pd.notna(r['日K']) and r['日K'] > 80: ko80.append("日")
        if pd.notna(r['週K']) and r['週K'] > 80: ko80.append("週")
        if pd.notna(r['月K']) and r['月K'] > 80: ko80.append("月")

        bot_trig = bool(r.get('落底觸發', False))
        
        if drop or ko80 or bot_trig:
            short_s = r['短線燈號']
            short_html = ""
            if short_s != "—":
                s_color = "#d32f2f" if "站" in short_s else "#1b8e3a" if "破" in short_s else "#1976d2"
                short_html = f" ｜ <span style='color:{s_color}; font-weight:400;'>{short_s}</span>"

            li_html = f"<li style='font-weight:400;'>{r['名稱']}: "
            if bot_trig:
                li_html += f"<span style='color:#d32f2f; font-weight:400;'>{r['落底訊號']}</span> "
            if drop: li_html += " ".join(drop) + " "
            if ko80: li_html += f"<span style='color:#d32f2f; font-weight:400;'><span class='s-dot'>●</span>[熱]{'/'.join(ko80)}K>80</span>"
            li_html += short_html
            li_html += "</li>"
            html += li_html
            alert_count_html += 1
            
    if alert_count_html == 0: html += "<li style='font-weight:400;'>*(今日無觸發任何位階警示)*</li>"
    html += "</ol>"

    # -----------------------------------------------------------------------
    # 七、市場新聞 ( Yahoo 財經 )
    # -----------------------------------------------------------------------
    html += "<h2>七、市場新聞 ( Yahoo 財經 )</h2><ol style='line-height:1.8; font-size:13px; padding-left:24px; margin-bottom:24px;'>"
    for item in news_data: html += f"<li><a href='{item['link']}' target='_blank' style='font-weight:normal;'>{item['title']}</a></li>"
    html += "</ol>"

    html += """
    <div class="notes" style="margin-top: 24px; font-size: 12px; color: #666; background: #fcfcfc; padding: 16px; border-radius: 6px; border: 1px solid var(--line);">
        <b style="font-size: 13px; color: #333;">📝 燈號與位階判斷標準</b>
        <div style="display: flex; flex-wrap: wrap; gap: 16px; margin-top: 10px; line-height: 1.7;">
            <div style="flex: 1 1 210px;">
                <b style="color:#1f3b66;">1. 短線與落底訊號</b><br>
                🔥 <b>站雙線</b>：收盤 ＞ 5MA 且 ＞ 10MA (短線強勢)<br>
                ⚡ <b>站回五日</b>：收盤 ＞ 5MA 但 ＜ 10MA (止跌轉強)<br>
                ⚠️ <b>雙線破</b>：收盤 ＜ 5MA 且 ＜ 10MA (短線弱勢)<br>
                🟢 <b>10日穿月線 (落底反轉)</b>：限台大盤、0050、台積電、聯發科，10MA 站上月線視為落底反轉，開始分批投入子彈。
            </div>
            <div style="flex: 1 1 205px;">
                <b style="color:#1f3b66;">2. 長線燈號 (20週均線)</b><br>
                ⚠️ <b>破線空手</b>：跌破 20 週均線 (長線趨勢轉弱)<br>
                <span style="color:#1b8e3a;"><span class="s-dot">●</span></span> <b>K值超跌</b>：20 週線上，且任一 K值 ＜ 20<br>
                <span style="color:#d32f2f;"><span class="s-dot">●</span></span> <b>熱K>80</b>：20 週線上，且任一 K值 ＞ 80<br>
                ✓ <b>線上安全</b>：20 週線上，且 KD 無極端狀況
            </div>
            <div style="flex: 1 1 205px;">
                <b style="color:#1f3b66;">3. 匯率與三大法人籌碼判讀</b><br>
                <span style="color:#d32f2f;"><b>紅字 (升值/買超)</b></span>：台幣升值＝外資匯入；法人合計買超 → <b>易漲</b><br>
                <span style="color:#1b8e3a;"><b>綠字 (貶值/賣超)</b></span>：台幣貶值＝外資匯出；法人合計賣超 → <b>易跌</b>
            </div>
            <div style="flex: 1 1 240px;">
                <b style="color:#1f3b66;">4. 融資融券與空單定義</b><br>
                • <b>融資 (散多標)</b>：散戶借錢買股。大減＝籌碼洗淨 (<span style="color:#d32f2f;"><b>看多</b></span>)；大增＝散戶搶短 (<span style="color:#1b8e3a;"><b>看空</b></span>)。<br>
                • <b>融券 (散空標)</b>：散戶借股放空。大增＝具軋空燃料 (<span style="color:#d32f2f;"><b>看多</b></span>)；大減＝空單回補 (<span style="color:#1b8e3a;"><b>看空</b></span>)。<br>
                • <b>期貨空單 (外空標)</b>：外資期貨避險。負值 (-)＝加空 (<span style="color:#1b8e3a;"><b>看空</b></span>)；正值 (+)＝回補 (<span style="color:#d32f2f;"><b>看多</b></span>)，逾 3.5 萬口警戒。<br>
                • <b>借券賣 (外空標)</b>：外資借股現貨放空。增加 (+)＝賣壓增 (<span style="color:#1b8e3a;"><b>看空</b></span>)；減少 (-)＝還券回補 (<span style="color:#d32f2f;"><b>看多</b></span>)。
            </div>
        </div>
    </div>
    <div class='foot' style='margin-top:20px; font-size:12px;'>
        資料來源：台灣期交所 MIS / Yahoo Finance / 台灣奇摩股市 / 證交所 ｜ 報表產生器：V10.0 落底訊號雙端同步與三日總數對照版<br><br>
        <div style="display:flex; justify-content:center; align-items:center;">
            <a href='report_data.xlsx' download='stock_monitor.xlsx' class='action-btn btn-download'>📥 下載完整 Excel (.xlsx) 報表</a>
        </div>
    </div>
    </div></body></html>
    """
    with open(REPORT_HTML_PATH, "w", encoding="utf-8") as f: f.write(html)

# ---------------------------------------------------------------------------
# 組裝 LINE Flex Message 卡片（同步顯示限定 4 檔之「落底訊號 (10MA/月線)」）
# ---------------------------------------------------------------------------
def build_flex_message(closes, watch, kd_data, vix_hist, audit_summary, df_report, night_quote=None):
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    today = dt.datetime.now(tw_tz)
    wd = "一二三四五六日"[today.weekday()]
    
    C_UP, C_DOWN, C_FLAT = "#d32f2f", "#1b8e3a", "#8a9099"
    stale_set = audit_summary.get("stale_syms", set())
    
    contents = [
        { "type": "text", "text": f"📊 股市核心快報 {today.month:02d}/{today.day:02d} ({wd})", "weight": "bold", "size": "md", "color": "#1f3b66" },
        {"type": "separator", "margin": "md"}
    ]
    
    sec1 = [{"type": "text", "text": "▌一、大盤與核心狀況 (+/- 2%)", "size": "sm", "color": "#333333"}]
    item_idx = 1
    
    def add_sec1_item(name, sym, s, is_idx=False):
        nonlocal item_idx
        if s is not None and len(s) >= 2:
            last, prev = float(s.iloc[-1]), float(s.iloc[-2])
            diff = last - prev
            dchg = pct(last, prev)
            weather = get_weather_line(dchg)
            stale_mark = " ⚠️待查" if sym in stale_set else ""
            val_color = C_UP if diff > 0 else C_DOWN if diff < 0 else C_FLAT
            
            sec1.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": f"{item_idx}. {name} {f'{last:,.0f}' if is_idx else f'{last:,.1f}'} ("},
                    {"type": "span", "text": f"{f'{diff:+,.0f}' if is_idx else f'{diff:+,.1f}'} ; {fmt_pct(dchg)}", "color": val_color, "weight": "bold"},
                    {"type": "span", "text": f") {weather}{stale_mark}"} 
                ]
            })
            item_idx += 1
            return True
        return False

    for sym, name in INDICES: add_sec1_item(name, sym, closes.get(sym), True)
    sec1.append({"type": "text", "text": "-" * 15, "size": "sm", "color": "#8a9099"})
    
    core_items = []
    seen_core = set()
    s_twii = closes.get("^TWII")
    if add_sec1_item("台大盤", "^TWII", s_twii, True):
        core_items.append({"name": "台大盤", "yf": "^TWII", "last": float(s_twii.iloc[-1]), "ma20w": weekly_ma(s_twii, 20)})

    if night_quote and night_quote.get("price", 0) > 0:
        np = night_quote["price"]
        ndiff = night_quote["diff"]
        npct = night_quote["pct"]
        n_weather = get_weather_line(npct)
        n_color = C_UP if ndiff > 0 else C_DOWN if ndiff < 0 else C_FLAT
        sec1.append({
            "type": "text", "size": "sm", "wrap": True,
            "contents": [
                {"type": "span", "text": f"{item_idx}. 台指夜 {np:,.0f} ("},
                {"type": "span", "text": f"{ndiff:+,.0f} ; {fmt_pct(npct)}", "color": n_color, "weight": "bold"},
                {"type": "span", "text": f") {n_weather}"}
            ]
        })
        item_idx += 1
    else:
        sec1.append({
            "type": "text", "size": "sm", "color": "#8a9099",
            "text": f"{item_idx}. 台指夜 — (暫無盤後)"
        })
        item_idx += 1
        
    for w in watch:
        sym = w["yf_symbol"]
        if w["market"] == "台股" and (w["group"] == "核心" or "2330" in sym or "0050" in sym):
            if "正2" in w["name"] or "反1" in w["name"] or sym.endswith(("L.TW", "R.TW")) or sym in seen_core: continue
            seen_core.add(sym)
            disp_n = shorten_name(w["name"], w["code"], w["market"])
            if add_sec1_item(disp_n, sym, closes.get(sym), False):
                core_items.append({"name": disp_n, "yf": sym, "last": float(closes.get(sym).iloc[-1]), "ma20w": weekly_ma(closes.get(sym), 20)})
                
    contents.extend(sec1)
    contents.append({"type": "separator", "margin": "md"})
    
    sec2 = [{"type": "text", "text": "▌二、VIX 恐慌指數 (<35)", "size": "sm", "color": "#333333"}]
    for i, (v_name, v_sym) in enumerate([("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")], 1):
        h = vix_hist.get(v_sym, [0, 0, 0, 0])
        v = h[0]
        if v is None or v <= 0: sec2.append({"type": "text", "text": f"{i}. {v_name} — (暫無報價)", "size": "sm", "color": "#8a9099"})
        else:
            c_color = C_DOWN if v < 35 else C_UP
            yest_str = f"昨日{h[1]:.1f}" if h[1] > 0 else "昨日—"
            prev_str = f"前一週{h[3]:.1f}" if h[3] > 0 else "前一週—"
            
            sec2.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": f"{i}. {v_name} {v:.1f} ("},
                    {"type": "span", "text": "●", "color": c_color, "size": "xxs"},
                    {"type": "span", "text": "安全" if v < 35 else "恐慌", "color": c_color, "weight": "bold"},
                    {"type": "span", "text": f"), {yest_str}, {prev_str}"}
                ]
            })
    contents.extend(sec2)
    contents.append({"type": "separator", "margin": "md"})
    
    sec3 = [{"type": "text", "text": "▌三、台股核心指標 (長/短線/落底燈號)", "size": "sm", "color": "#333333"}]
    for i, item in enumerate(core_items, 1):
        last, ma20w = item["last"], item["ma20w"]
        df_row = df_report[df_report['yf_symbol'] == item["yf"]]
        if df_row.empty: continue
        r = df_row.iloc[0]
        kd_d, kd_w, kd_m, short_status = r['日K'], r['週K'], r['月K'], r['短線燈號']
        short_color = C_UP if "站" in short_status else C_DOWN if "破" in short_status else "#1976d2"
        
        sec3.append({"type": "text", "text": f"{i}. {item['name']} ({f'{last:,.0f}' if item['yf']=='^TWII' else f'{last:,.1f}'})", "size": "sm", "color": "#333333", "margin": "sm"})
        sec3.append({
            "type": "text", "size": "sm", "wrap": True,
            "contents": [
                {"type": "span", "text": f"   K值: 日{int(kd_d) if pd.notna(kd_d) else '-'}/週{int(kd_w) if pd.notna(kd_w) else '-'}/月{int(kd_m) if pd.notna(kd_m) else '-'} 短線: ", "color": "#666666"},
                {"type": "span", "text": short_status, "color": short_color}
            ]
        })

        # 顯示「落底訊號 (10MA/月線)」（限定台大盤、0050、台積電、聯發科）
        bot_txt = r.get('落底訊號', '—')
        if bot_txt and bot_txt != "—":
            bot_cls = r.get('落底樣式', 'mute')
            bot_color = C_UP if bot_cls == "up" else "#1f3b66" if bot_cls == "safe" else C_FLAT
            bot_weight = "bold" if bot_cls == "up" else "regular"
            sec3.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": "   落底訊號 (10MA/月線): ", "color": "#666666"},
                    {"type": "span", "text": bot_txt, "color": bot_color, "weight": bot_weight}
                ]
            })
        
        if last and ma20w:
            dist = (last - ma20w) / ma20w * 100
            dist_color = C_UP if dist > 0 else C_DOWN if dist < 0 else C_FLAT
            line_spans = [
                {"type": "span", "text": f"   長線距20週({f'{ma20w:,.0f}' if item['yf']=='^TWII' else f'{ma20w:,.1f}'}): "},
                {"type": "span", "text": f"{dist:+.1f}%", "color": dist_color, "weight": "bold"},
                {"type": "span", "text": " ("}
            ]
            if dist < 0: line_spans.append({"type": "span", "text": "⚠️破線空手)"})
            else:
                if pd.notna(kd_d) and kd_d < 20: line_spans.extend([{"type": "span", "text": "●", "color": C_DOWN, "size": "xxs"}, {"type": "span", "text": "日K超跌)"}])
                elif pd.notna(kd_w) and kd_w < 20: line_spans.extend([{"type": "span", "text": "●", "color": C_DOWN, "size": "xxs"}, {"type": "span", "text": "週K超跌)"}])
                elif pd.notna(kd_m) and kd_m < 20: line_spans.extend([{"type": "span", "text": "●", "color": C_DOWN, "size": "xxs"}, {"type": "span", "text": "月K超跌)"}])
                else: line_spans.append({"type": "span", "text": "✓線上安全)"})
            sec3.append({"type": "text", "size": "sm", "wrap": True, "contents": line_spans})
    contents.extend(sec3)
    contents.append({"type": "separator", "margin": "md"})
    
    sec4 = [{"type": "text", "text": "▌四、其他位階警示", "size": "sm", "color": "#333333"}]
    alert_df = df_report[df_report['長線燈號'].isin(['大盤', '核心'])]
    
    idx_alerts, tw_alerts, us_alerts = [], [], []
    for _, r in alert_df.iterrows():
        last, ma60, ma240 = r['最新收盤'], r['MA60'], r['MA240']
        drop_q = (pd.notna(ma60) and last < ma60)
        drop_y = (pd.notna(ma240) and last < ma240)
        w_ret, m_ret = r['週%'], r['月%']
        drop_20 = (pd.notna(w_ret) and w_ret <= -20) or (pd.notna(m_ret) and m_ret <= -20)
        ko80 = [k for k, val in zip(["日", "週", "月"], [r['日K'], r['週K'], r['月K']]) if pd.notna(val) and val > 80]
        bot_trig = bool(r.get('落底觸發', False))
        
        if drop_q or drop_y or drop_20 or ko80 or bot_trig:
            spans = [{"type": "span", "text": f"📍 {r['名稱']}: ", "weight": "regular"}]
            if bot_trig: spans.append({"type": "span", "text": f"{r['落底訊號']} ", "color": C_UP, "weight": "bold"})
            if drop_q: spans.extend([{"type": "span", "text": "●", "color": C_DOWN, "size": "xxs", "weight": "regular"}, {"type": "span", "text": "[跌]破季線 ", "color": C_DOWN, "weight": "regular"}])
            if drop_y: spans.extend([{"type": "span", "text": "●", "color": C_UP, "size": "xxs", "weight": "regular"}, {"type": "span", "text": "[跌]破年線 ", "color": C_UP, "weight": "regular"}])
            if drop_20: spans.extend([{"type": "span", "text": "●", "color": C_DOWN, "size": "xxs", "weight": "regular"}, {"type": "span", "text": "週/月跌>20% ", "color": C_DOWN, "weight": "regular"}])
            if ko80: spans.extend([{"type": "span", "text": "●", "color": C_UP, "size": "xxs", "weight": "regular"}, {"type": "span", "text": f"[熱]{'/'.join(ko80)}K>80", "color": C_UP, "weight": "regular"}])
            
            short_s = r['短線燈號']
            if short_s != "—":
                s_color = C_UP if "站" in short_s else C_DOWN if "破" in short_s else "#1976d2"
                spans.extend([{"type": "span", "text": " ｜ ", "weight": "regular"}, {"type": "span", "text": short_s, "color": s_color, "weight": "regular"}])
                
            block = {"type": "text", "size": "sm", "weight": "regular", "wrap": True, "contents": spans, "margin": "sm"}
            
            if r['名稱'] == "台大盤": tw_alerts.insert(0, block)
            elif r['長線燈號'] == "大盤": idx_alerts.append(block)
            elif r['市場'] == "台股": tw_alerts.append(block)
            else: us_alerts.append(block)

    has_alerts = False
    if idx_alerts:
        sec4.extend(idx_alerts)
        has_alerts = True
    if tw_alerts:
        if has_alerts: sec4.append({"type": "separator", "margin": "sm", "color": "#000000"})
        sec4.extend(tw_alerts)
        has_alerts = True
    if us_alerts:
        if has_alerts: sec4.append({"type": "separator", "margin": "sm", "color": "#000000"})
        sec4.extend(us_alerts)
        has_alerts = True

    if not has_alerts: sec4.append({"type": "text", "text": "*(今日無觸發)*", "size": "sm", "color": "#8a9099"})
    contents.extend(sec4)
    
    contents.append({"type": "separator", "margin": "md"})
    contents.append({
        "type": "text", "size": "xxs", "color": "#8a9099", "wrap": True,
        "contents": [
            {"type": "span", "text": "💡 短線: 🔥站雙線(>5/10MA)｜⚡站回五日(>5MA)｜⚠️雙線破\n💡 落底(限台大盤/0050/2330/2454): 🟢10日穿月線 (落底反轉)\n💡 長線: 20週均上為✓安全，破線為⚠️空手｜"},
            {"type": "span", "text": "●", "color": C_DOWN, "size": "xxs"}, {"type": "span", "text": "破季線 "},
            {"type": "span", "text": "●", "color": C_UP, "size": "xxs"}, {"type": "span", "text": "破年線"}
        ]
    })
    
    contents.append({
        "type": "button", "style": "primary", "color": "#1456b8", "height": "sm", "margin": "sm",
        "action": { "type": "uri", "label": "👉 開啟詳細網頁報表", "uri": os.environ.get("GITHUB_PAGES_URL", "").strip() or "https://ckliu60.github.io/stock-monitor/" }
    })

    return {
        "type": "flex",
        "altText": f"📊 股市快報 {today.month:02d}/{today.day:02d} 已送達",
        "contents": {
            "type": "bubble", "size": "mega",
            "body": { "type": "box", "layout": "vertical", "spacing": "md", "contents": contents }
        }
    }

def main():
    try:
        force_dry_run = "--dry-run" in sys.argv
        is_manual_run = os.environ.get("GITHUB_EVENT_NAME") in ["workflow_dispatch", "repository_dispatch"]
        
        tw_tz = dt.timezone(dt.timedelta(hours=8))
        today_tw = dt.datetime.now(tw_tz)
        weekday = today_tw.weekday() # 0=週一, 6=週日
        
        skip_line_push = False
        if not force_dry_run and not is_manual_run and weekday in [0, 6]:
            skip_line_push = True
            print("[智慧免打擾] 今日為週日或週一，自動暫停 LINE 推播，僅更新報表。")

        watch = load_watchlist()
        idx_syms = [s for s, _ in INDICES]
        stk_syms = [w["yf_symbol"] for w in watch]
        all_syms = list(dict.fromkeys(idx_syms + ["^TWII"] + stk_syms))
        tw_symbols_set = set(["^TWII"] + [w["yf_symbol"] for w in watch if w["market"] == "台股"])
        
        closes, kd_data, audit_summary = fetch_and_audit_all(all_syms, tw_symbols_set)
        
        s_twii = closes.get("^TWII")
        twii_last = float(s_twii.iloc[-1]) if (s_twii is not None and len(s_twii) > 0) else None
        night_quote = fetch_taifex_night_quote(day_close_ref=twii_last)

        fx_data = fetch_usdtwd_analysis()
        inst_data = fetch_institutional_flow()
        retail_data, inst_short_data = fetch_margin_and_shorts()

        vix_hist = {}
        v_closes, _, _ = fetch_and_audit_all(["^VIX", "^VIXTWN"], set(), period="1mo")
        for v_sym in ["^VIX", "^VIXTWN"]:
            s = v_closes.get(v_sym)
            if s is not None and len(s) >= 6: vix_hist[v_sym] = [float(s.iloc[-1]), float(s.iloc[-2]), float(s.iloc[-3]), float(s.iloc[-6])]
            elif s is not None and len(s) > 0: vix_hist[v_sym] = [float(s.iloc[-1]), 0, 0, 0]
            else: vix_hist[v_sym] = [0, 0, 0, 0]

        news_data = fetch_yahoo_news(limit=5)
        df_report = export_all_data(closes, watch, kd_data, audit_summary)
        generate_html_report(df_report, vix_hist, news_data, audit_summary, night_quote, fx_data, inst_data, retail_data, inst_short_data)
        
        if skip_line_push:
            return 0
        
        flex_message = build_flex_message(closes, watch, kd_data, vix_hist, audit_summary, df_report, night_quote)

        token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
        uid = os.environ.get("LINE_USER_ID", "").strip()
        
        if force_dry_run or not token or not uid:
            print("[DRY-RUN] 訊息預覽 JSON：\n" + json.dumps(flex_message, indent=2, ensure_ascii=False) + "\n")
            return 0

        resp = SESSION.post(LINE_PUSH_URL, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                            json={"to": uid, "messages": [flex_message]}, timeout=20)
        try:
            resp.raise_for_status()
            print("[OK] 已成功推播卡片訊息至 LINE。")
        except Exception as e:
            print(f"[錯誤] LINE 推播失敗: {e}\n{resp.text}", file=sys.stderr)
            
        return 0
    except Exception:
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
