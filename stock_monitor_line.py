#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (V9.7 Yahoo Chart 全球實價引擎與真實日期防呆鎖版)
---------------------------------------------------------------------------
更新日誌 (V9.7):
1. [徹底修復台股與 ETF 今日收盤及漲跌幅] 新增 Yahoo Finance 全球 CDN Chart API (v8/finance/chart) 作為主力實價來源，精準解析 regularMarketTime 真實交易日與昨日收盤，徹底解決 2330、2454 及 0050、0052、00919、00929 停留在昨日收盤或漲跌幅顯示 0.00/0.0% 的問題（HTML 與 LINE 卡片同步生效）。
2. [真實交易日期防呆鎖] 移除 fast_info 強制覆寫今日日期的缺陷，杜絕重複貼上舊價。
3. [證交所 SSL 容錯] 強化證交所 OpenAPI 與 RWD 接口在雲端 Linux 主機的憑證連線穩定度。
4. [完整保留 V9.4~V9.6 功能] 保留「二、外資與投信買賣超 [V1]」、「三、台美匯率與外資動向 [V1]」、極速台指夜盤與原始 LINE 卡片排版。
"""

import os
import sys
import io
import re
import csv
import time
import json
import urllib3
import datetime as dt
import xml.etree.ElementTree as ET
from pathlib import Path

# 關閉憑證備援警告並強制關閉 curl_cffi 避免 DNS 逾時
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
os.environ["YFINANCE_CACHE_DIR"] = ""
os.environ["CURL_CA_BUNDLE"] = ""

import requests
import pandas as pd
import yfinance as yf

# ---------------------------------------------------------------------------
# 設定區
# ---------------------------------------------------------------------------

INDICES = [
    ("^DJI", "道瓊"),
    ("^IXIC", "那指"),
    ("^GSPC", "標普"),
    ("^SOX", "費半"),
]

DEFAULT_WATCHLIST = [
    {"market": "台股", "code": "2330", "yf_symbol": "2330.TW", "name": "台積", "group": "核心"},
    {"market": "台股", "code": "2454", "yf_symbol": "2454.TW", "name": "聯發科", "group": "核心"},
    {"market": "台股", "code": "0050", "yf_symbol": "0050.TW", "name": "0050", "group": "核心"},
    {"market": "台股", "code": "0052", "yf_symbol": "0052.TW", "name": "富邦科技", "group": "觀察"},
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

def _safe_get(url, params=None, timeout=10):
    """具備 SSL 中繼憑證容錯的 HTTP GET 請求"""
    try:
        return requests.get(url, params=params, headers=HEADERS, timeout=timeout)
    except requests.exceptions.SSLError:
        return requests.get(url, params=params, headers=HEADERS, timeout=timeout, verify=False)

def shorten_name(name):
    n = name.strip()
    if "群益台灣精選高息" in n: return "00919"
    if "復華台灣科技優息" in n: return "00929"
    if "元大台灣50" in n: return "0050"
    if n == "台積電ADR": return "台積電R"
    if n == "台積電": return "台積"
    if n == "台股大盤": return "台大盤"
    return n

def get_latest_tw_trading_date_str():
    """取得台灣最新已收盤交易日 (YYYY-MM-DD)：平日 13:35 後為今日，否則為前一交易日"""
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)
    ref_dt = now_tw
    if now_tw.weekday() < 5 and (now_tw.hour < 13 or (now_tw.hour == 13 and now_tw.minute < 35)):
        ref_dt -= dt.timedelta(days=1)
    while ref_dt.weekday() >= 5:
        ref_dt -= dt.timedelta(days=1)
    return ref_dt.strftime("%Y-%m-%d")

def is_tw_etf_symbol(sym):
    """判斷是否為台股 ETF（代碼以 00 開頭，如 0050.TW, 0052.TW, 00919.TW, 00929.TW）"""
    if not sym: return False
    base = sym.split(".")[0].strip()
    return base.startswith("00") and base.isdigit()

def load_watchlist():
    bust_url = f"{SHEET_CSV_URL}&_t={int(time.time())}"
    cloud_rows = []
    seen = set()
    try:
        resp = requests.get(bust_url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        for raw in csv.DictReader(io.StringIO(resp.text)):
            code = (raw.get("code") or "").strip()
            yf_symbol = (raw.get("yf_symbol") or "").strip()
            market = (raw.get("market") or "").strip()
            group = (raw.get("長線燈號") or raw.get("group") or "觀察").strip()
            
            if not code and not yf_symbol: continue
            if code in ["^DJI", "^IXIC", "^GSPC", "^SOX", "^TWII"] or group == "大盤": continue
            if market != "台股": market = "美股"

            if market == "台股" and code.isdigit() and len(code) in (2, 3): code = code.zfill(4)

            if not yf_symbol:
                yf_symbol = f"{code}.TW" if market == "台股" else code.replace(".", "-")
            else:
                if market == "美股": yf_symbol = yf_symbol.replace(".", "-")

            if yf_symbol in seen: continue
            seen.add(yf_symbol)

            original_name = (raw.get("name") or yf_symbol).strip()
            cloud_rows.append({
                "market": market, "code": code or yf_symbol, "yf_symbol": yf_symbol,
                "name": shorten_name(original_name), "group": group
            })
    except Exception as e:
        print(f"[提示] 讀取雲端試算表異常: {e}", file=sys.stderr)

    if cloud_rows: return cloud_rows
    return [item.copy() for item in DEFAULT_WATCHLIST]

# ---------------------------------------------------------------------------
# 多重實價採集引擎（Yahoo Global Chart API + 台灣奇摩 + 證交所）& 台指夜盤直抓
# ---------------------------------------------------------------------------

def fetch_yf_chart_quotes(symbols):
    """
    透過 Yahoo Finance 全球 CDN Chart API (v8/finance/chart) 取得最新收盤價與真實交易時間戳。
    100% 暢通於雲端主機，完整支援所有台股普通股 (2330.TW) 與 ETF (0050.TW, 0052.TW, 00919.TW, 00929.TW)，
    並以 regularMarketTime 換算真實交易日，絕不誤貼昨日舊價。
    """
    quotes = {}
    if not symbols:
        return quotes
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    us_tz = dt.timezone(dt.timedelta(hours=-4))

    for sym in symbols:
        if "-USD" in sym.upper():
            continue
        for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
            try:
                url = f"https://{host}/v8/finance/chart/{sym}"
                params = {"interval": "1d", "range": "5d"}
                resp = requests.get(url, params=params, headers=HEADERS, timeout=10)
                if resp.status_code != 200:
                    continue
                data = resp.json()
                res_list = (data.get("chart") or {}).get("result") or []
                if not res_list:
                    continue
                meta = res_list[0].get("meta") or {}
                rm_price = meta.get("regularMarketPrice")
                rm_time = meta.get("regularMarketTime")
                prev_close = meta.get("chartPreviousClose") or meta.get("previousClose")

                timestamps = res_list[0].get("timestamp") or []
                q_closes = (((res_list[0].get("indicators") or {}).get("quote") or [{}])[0]).get("close") or []
                valid_pairs = [(ts, c) for ts, c in zip(timestamps, q_closes) if ts is not None and c is not None and float(c) > 0]

                is_tw = sym.endswith(".TW") or sym.endswith(".TWO") or sym == "^TWII"
                tz_use = tw_tz if is_tw else us_tz

                date_str = None
                if rm_time:
                    date_str = dt.datetime.fromtimestamp(int(rm_time), tz_use).strftime("%Y-%m-%d")
                elif valid_pairs:
                    date_str = dt.datetime.fromtimestamp(int(valid_pairs[-1][0]), tz_use).strftime("%Y-%m-%d")

                if valid_pairs and len(valid_pairs) >= 2:
                    last_bar_date = dt.datetime.fromtimestamp(int(valid_pairs[-1][0]), tz_use).strftime("%Y-%m-%d")
                    if date_str and last_bar_date == date_str:
                        prev_close = float(valid_pairs[-2][1])
                    elif date_str and last_bar_date < date_str:
                        prev_close = float(valid_pairs[-1][1])

                if rm_price and float(rm_price) > 0 and date_str:
                    quotes[sym] = {
                        "price": float(rm_price),
                        "prev_close": float(prev_close) if prev_close and float(prev_close) > 0 else None,
                        "date": date_str,
                        "source": "YahooChartAPI"
                    }
                    break
            except Exception:
                continue
    return quotes

def fetch_tw_yahoo_quotes(tw_symbols):
    """抓取台灣奇摩即時/盤後收盤報價（僅採信帶有真實日期或明確已變動之收盤價）"""
    quotes = {}
    if not tw_symbols: return quotes
    try:
        url = "https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.stockList"
        params = {
            "fields": "symbol,price,regularMarketPrice,regularMarketPreviousClose,regularMarketTime",
            "symbols": ",".join(tw_symbols)
        }
        resp = requests.get(url, params=params, headers=HEADERS, timeout=12)
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
                    except:
                        continue
                    prev = None
                    if prev_raw not in (None, "-", ""):
                        try: prev = float(str(prev_raw).replace(",", ""))
                        except: pass

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
                        except:
                            pass

                    # 僅當具備真實日期戳記時才納入，避免無日期舊價誤貼今日
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
        r = _safe_get("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", timeout=12)
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
                    except: pass
    except Exception as e:
        print(f"[證交所提示] STOCK_DAY_ALL 略過: {e}")

    if not latest_tw_date:
        latest_tw_date = get_latest_tw_trading_date_str()

    # 針對 STOCK_DAY_ALL 未收錄的 00 開頭 ETF，單獨向證交所 STOCK_DAY 官方月行情接口補抓最新收盤價
    if tw_symbols:
        ym_str = latest_tw_date.replace("-", "")[:6] + "01"
        for sym in tw_symbols:
            if not is_tw_etf_symbol(sym) or sym in twse_map:
                continue
            code = sym.split(".")[0].strip()
            try:
                url_etf = f"https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={ym_str}&stockNo={code}&response=json"
                r_etf = _safe_get(url_etf, timeout=8)
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
                            except:
                                pass
                        if close_val > 0 and d_str:
                            twse_map[sym] = {
                                "price": close_val,
                                "prev_close": prev_val,
                                "date": d_str,
                                "source": "證交所ETF官方日線"
                            }
                            print(f"[證交所ETF補齊] {sym} ({d_str}) 收盤: {close_val}")
                time.sleep(0.15)
            except Exception:
                pass

    try:
        r2 = _safe_get("https://openapi.twse.com.tw/v1/exchangeReport/MI_INDEX", timeout=12)
        if r2.status_code == 200:
            for item in r2.json():
                idx_name = item.get("指數", "")
                if "發行量加權股價指數" in idx_name:
                    close_s = (item.get("收盤指數") or "").replace(",", "").strip()
                    chg_s = (item.get("漲跌點數") or "").replace(",", "").strip()
                    dir_s = (item.get("漲跌") or "").strip()
                    if close_s:
                        p = float(close_s)
                        prev_p = None
                        try:
                            if chg_s:
                                chg_val = float(chg_s)
                                if "-" in dir_s or "跌" in dir_s: chg_val = -abs(chg_val)
                                prev_p = round(p - chg_val, 2)
                        except: pass
                        if p > 0:
                            twse_map["^TWII"] = {"price": p, "prev_close": prev_p, "date": latest_tw_date, "source": "證交所TWSE"}
                    break
    except: pass
    return twse_map

def fetch_taifex_night_quote(day_close_ref=None):
    """
    【獨立函式】極速直抓台指期夜盤（清晨 05:00 收盤）現成數值，不下載 K 線、不影響其他指數。
    針對清晨 05:00~08:45 盤前清空時段，支援讀取網頁保留之 regularMarketPreviousClose（夜盤收盤價）。
    """
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)
    date_label = now_tw.strftime("%m-%d")

    def _to_float(val):
        if val in (None, "-", "", "NaN"): return None
        try: return float(str(val).replace(",", "").strip())
        except: return None

    # 來源 1：台灣奇摩新版期貨 API (WTX& 台指期近月)
    try:
        url = "https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.stockList"
        params = {
            "fields": "symbol,price,change,changePercent,regularMarketPrice,regularMarketChange,regularMarketPreviousClose,regularMarketTime",
            "symbols": "WTX&"
        }
        resp = requests.get(url, params=params, headers=HEADERS, timeout=10)
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
                    except: pass

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
                    print(f"[台指夜盤] 奇摩 API 即時/盤後欄取得: {p_cur} ({diff:+.0f})")
                    return {"price": p_cur, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨API"}

                if (not p_cur or p_cur <= 0) and p_prev and p_prev > 0:
                    base = day_close_ref if (day_close_ref and day_close_ref > 0) else p_prev
                    diff = (p_prev - base) if (day_close_ref and day_close_ref > 0) else 0.0
                    pct_val = (diff / base * 100.0) if base > 0 else 0.0
                    print(f"[台指夜盤] 奇摩 API 盤前結算欄取得: {p_prev} ({diff:+.0f})")
                    return {"price": p_prev, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨API(盤後結算)"}
    except Exception as e:
        print(f"[夜盤提示] 奇摩 API 查詢略過: {e}")

    # 來源 2：直接讀取台灣奇摩期貨網頁現成數字 (https://tw.stock.yahoo.com/future/WTX%26)
    try:
        r_web = requests.get("https://tw.stock.yahoo.com/future/WTX%26", headers=HEADERS, timeout=10)
        if r_web.status_code == 200:
            txt = r_web.text
            m_price = re.search(r'"(?:regularMarketPrice|price)"\s*:\s*\{\s*"raw"\s*:\s*([0-9.]+)', txt)
            m_chg = re.search(r'"(?:regularMarketChange|change)"\s*:\s*\{\s*"raw"\s*:\s*([-0-9.]+)', txt)
            m_prev = re.search(r'"regularMarketPreviousClose"\s*:\s*\{\s*"raw"\s*:\s*([0-9.]+)', txt)

            p_cur = float(m_price.group(1)) if m_price else None
            p_prev = float(m_prev.group(1)) if m_prev else None
            chg_cur = float(m_chg.group(1)) if m_chg else None

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
                print(f"[台指夜盤] 奇摩期貨網頁取得: {p_cur} ({diff:+.0f})")
                return {"price": p_cur, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨網頁"}

            if (not p_cur or p_cur <= 0) and p_prev and p_prev > 0:
                base = day_close_ref if (day_close_ref and day_close_ref > 0) else p_prev
                diff = (p_prev - base) if (day_close_ref and day_close_ref > 0) else 0.0
                pct_val = (diff / base * 100.0) if base > 0 else 0.0
                print(f"[台指夜盤] 奇摩期貨網頁盤前結算欄取得: {p_prev} ({diff:+.0f})")
                return {"price": p_prev, "diff": diff, "pct": pct_val, "date": date_label, "source": "奇摩期貨網頁(盤後結算)"}
    except Exception as e:
        print(f"[夜盤提示] 奇摩網頁直抓略過: {e}")

    print("[夜盤警告] 無法取得台指夜盤報價，將顯示暫無盤後。")
    return None

# ---------------------------------------------------------------------------
# 【獨立新模組 V1】台美匯率與外資動向 (USD/TWD) [V1]
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
# 【獨立新模組 V1】外資與投信買賣超（單位：億元）[V1]
# ---------------------------------------------------------------------------

def fetch_institutional_flow():
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_tw = dt.datetime.now(tw_tz)

    daily_records = []
    seen_dates = set()

    def _parse_num(s):
        if s is None: return 0.0
        try: return float(str(s).replace(",", "").strip())
        except: return 0.0

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
        r_open = _safe_get("https://openapi.twse.com.tw/v1/fund/BFI82U", timeout=10)
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
            resp = _safe_get(url, timeout=8)
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
            time.sleep(0.12)
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
# 主抓取與自動稽核引擎（真實交易日期防呆鎖，確保漲跌幅 100% 精準）
# ---------------------------------------------------------------------------

def fetch_and_audit_all(symbols, tw_symbols_set, period="2y"):
    closes, kd_data, audit_logs = {}, {}, []
    corrected_count = 0

    tw_list = list(tw_symbols_set)
    # 1. 全球 CDN Yahoo Chart API（主攻台股普通股 + 所有 00 開頭 ETF，附帶真實交易時間戳記）
    chart_map = fetch_yf_chart_quotes(symbols)
    # 2. 台灣奇摩股市 API（帶有真實日期之報價備援）
    tw_yahoo_map = fetch_tw_yahoo_quotes(tw_list)
    # 3. 證交所官方行情（含 STOCK_DAY_ALL 與 ETF 專屬 STOCK_DAY 月行情備援）
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

            # 依優先順序挑選「具備真實日期」的實價來源（絕不使用無日期之 fast_info 偽造今日資料）
            alt_info = None
            if sym in tw_symbols_set:
                c_q = chart_map.get(sym)
                y_q = tw_yahoo_map.get(sym)
                t_q = twse_map.get(sym)
                candidates = [q for q in (c_q, y_q, t_q) if q and q.get("price", 0) > 0 and q.get("date")]
                if candidates:
                    # 優先採用日期最新的來源
                    candidates.sort(key=lambda x: x["date"], reverse=True)
                    alt_info = candidates[0]
            else:
                alt_info = chart_map.get(sym)

            if alt_info and alt_info["price"] > 0 and alt_info.get("date"):
                alt_p = float(alt_info["price"])
                alt_d = alt_info["date"]
                alt_prev = alt_info.get("prev_close")

                if alt_d > yf_last_date:
                    # 真實交易日確實比歷史日線新：先校準昨日收盤，再補入今日收盤
                    new_ts = pd.Timestamp(alt_d)
                    if alt_prev and alt_prev > 0 and abs(yf_last_price - alt_prev) / alt_prev > 0.001:
                        df.loc[df.index[-1], "Close"] = alt_prev
                    df.loc[new_ts] = {"High": alt_p, "Low": alt_p, "Close": alt_p}
                    corrected_count += 1
                    print(f"[自動校正] {sym} 補入 {alt_info['source']} 最新收盤 ({alt_d}): {alt_p} (昨收: {alt_prev})")
                elif alt_d == yf_last_date:
                    # 同一交易日：僅當收盤價有微調時更新當日收盤，並校準前一日收盤價
                    if abs(alt_p - yf_last_price) / yf_last_price > 0.001:
                        df.loc[df.index[-1], "Close"] = alt_p
                        corrected_count += 1
                        print(f"[自動校正] {sym} 更新當日收盤價 ({alt_d}): {alt_p}")
                    if alt_prev and alt_prev > 0 and len(df) >= 2:
                        df_prev = float(df["Close"].iloc[-2])
                        if abs(df_prev - alt_prev) / alt_prev > 0.001:
                            df.loc[df.index[-2], "Close"] = alt_prev

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
# 技術指標與輔助
# ---------------------------------------------------------------------------

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
        resp = requests.get(url, headers=HEADERS, timeout=10)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        for item in root.findall('.//item')[:limit]:
            title = item.find('title')
            link = item.find('link')
            if title is not None and link is not None: news.append({"title": title.text, "link": link.text})
    except Exception: pass
    if not news: news.append({"title": "暫時無法取得新聞。", "link": "#"})
    return news

# ---------------------------------------------------------------------------
# 匯出資料
# ---------------------------------------------------------------------------
def export_all_data(closes, watch, kd_data, audit_summary):
    all_targets = [{"market": "美股", "code": s, "name": n, "yf": s, "group": "大盤"} for s, n in INDICES] + \
                  [{"market": "台股", "code": "^TWII", "name": "台大盤", "yf": "^TWII", "group": "大盤"}] + \
                  [{"market": w["market"], "code": w["code"], "name": w["name"], "yf": w["yf_symbol"], "group": w["group"]} for w in watch]
    
    rows = []
    stale_set = audit_summary.get("stale_syms", set())
    for i, t in enumerate(all_targets, 1):
        sym = t["yf"]
        s = closes.get(sym)
        if s is None or len(s) < 1: continue
        
        last = float(s.iloc[-1])
        pct_1d = pct(last, s.iloc[-2]) if len(s) >= 2 else None
        net_1d = (last - s.iloc[-2]) if len(s) >= 2 else None
        pct_1w = pct(last, s.iloc[-6]) if len(s) >= 6 else None
        pct_1m = pct(last, s.iloc[-22]) if len(s) >= 22 else None
        pct_1y = pct(last, s.iloc[-253]) if len(s) >= 253 else None
        
        ma = compute_ma(s)
        ma20w = weekly_ma(s, 20)
        kd = kd_data.get(sym, {})
        kd_d, kd_w, kd_m = kd.get("日", (None, None)), kd.get("週", (None, None)), kd.get("月", (None, None))
        
        audit_flag = "⚠️日期延遲" if sym in stale_set else "✅正常"
        short_status = get_short_term_status(last, ma.get("MA5"), ma.get("MA10"))
        
        rows.append({
            "序號": i, "長線燈號": t["group"], "短線燈號": short_status, "市場": t["market"], "代碼": t["code"],
            "名稱": t["name"], "yf_symbol": sym, "最新收盤": round(last, 2),
            "資料日期": s.index[-1].strftime("%Y-%m-%d"), "稽核狀態": audit_flag,
            "當日漲跌": round(net_1d, 2) if net_1d is not None else None, "當日%": round(pct_1d, 2) if pct_1d is not None else None,
            "週%": round(pct_1w, 2) if pct_1w is not None else None, "月%": round(pct_1m, 2) if pct_1m is not None else None,
            "年%": round(pct_1y, 2) if pct_1y is not None else None,
            "日K": kd_d[0], "日D": kd_d[1], "週K": kd_w[0], "週D": kd_w[1], "月K": kd_m[0], "月D": kd_m[1],
            "MA5": ma.get("MA5"), "MA10": ma.get("MA10"), "MA20": ma.get("MA20"), "MA60": ma.get("MA60"), "MA240": ma.get("MA240"), "MA20週": ma20w
        })
        
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    
    try:
        with pd.ExcelWriter(REPORT_EXCEL_PATH, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='股市監控報表', index=False)
            workbook, worksheet = writer.book, writer.sheets['股市監控報表']
            white_bg = workbook.add_format({'bg_color': '#FFFFFF'})
            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'center', 'align': 'center', 'fg_color': '#D7E4BC', 'border': 1})
            for col_num, value in enumerate(df.columns.values): worksheet.write(0, col_num, value, header_format)
            worksheet.set_column('A:E', 10, white_bg)
            worksheet.set_column('F:G', 12, white_bg)  
            worksheet.set_column('H:Z', 10, white_bg)  
            worksheet.freeze_panes(1, 6)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            tw_
