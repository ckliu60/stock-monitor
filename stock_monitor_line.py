#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (V8.2 清爽字體版)
---------------------------------------------------------------------------
更新日誌 (V8.2):
1. [視覺減壓] 移除 LINE 圖卡與 HTML 網頁報表中「四、其他位階警示」的所有粗體設定，改為清爽一般字體。
2. [長短線雙燈號架構] 完美對接試算表新增的 5MA, 10MA 與短線燈號欄位。
3. [自訂警示顏色] 網頁與 LINE 圖卡同步：跌破季線標綠色 🟢，跌破年線標紅色 🔴。
"""

import os
import sys
import io
import csv
import time
import json
import datetime as dt
import xml.etree.ElementTree as ET
from pathlib import Path

# 強制關閉 curl_cffi 避免 DNS 逾時
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

# 斷線時的緊急備援清單 (平時以雲端試算表 100% 為主，不強制補回已刪除股票)
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

# 老爺最新發布的 Google 試算表 CSV 網址
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
# 工具函式與雲端清單讀取 (100% 雲端試算表驅動)
# ---------------------------------------------------------------------------

def shorten_name(name):
    """極致縮短名稱，避免手機版 LINE 斷行"""
    n = name.strip()
    if "群益台灣精選高息" in n: return "00919"
    if "復華台灣科技優息" in n: return "00929"
    if "元大台灣50" in n: return "0050"
    if n == "台積電ADR": return "台積電R"
    if n == "台積電": return "台積"
    if n == "台股大盤": return "台大盤"
    return n

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
            
            if not code and not yf_symbol:
                continue
            if code in ["^DJI", "^IXIC", "^GSPC", "^SOX", "^TWII"] or group == "大盤":
                continue
            if market != "台股":
                market = "美股"

            if market == "台股" and code.isdigit() and len(code) in (2, 3):
                code = code.zfill(4)

            if not yf_symbol:
                yf_symbol = f"{code}.TW" if market == "台股" else code.replace(".", "-")
            else:
                if market == "美股":
                    yf_symbol = yf_symbol.replace(".", "-")

            if yf_symbol in seen:
                continue
            seen.add(yf_symbol)

            original_name = (raw.get("name") or yf_symbol).strip()
            cloud_rows.append({
                "market": market,
                "code": code or yf_symbol,
                "yf_symbol": yf_symbol,
                "name": shorten_name(original_name),
                "group": group
            })
    except Exception as e:
        print(f"[提示] 讀取雲端試算表異常，啟用緊急保底清單: {e}", file=sys.stderr)

    if cloud_rows:
        return cloud_rows

    return [item.copy() for item in DEFAULT_WATCHLIST]

# ---------------------------------------------------------------------------
# 第二資料源：台灣奇摩股市 API & 台灣證交所官方 API
# ---------------------------------------------------------------------------

def fetch_tw_yahoo_quotes(tw_symbols):
    quotes = {}
    if not tw_symbols:
        return quotes
    try:
        url = "https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.stockList"
        params = {"fields": "symbol,price,regularMarketPreviousClose,regularMarketTime", "symbols": ",".join(tw_symbols)}
        resp = requests.get(url, params=params, headers=HEADERS, timeout=12)
        if resp.status_code == 200:
            data = resp.json()
            items = data if isinstance(data, list) else data.get("data", [])
            tw_tz = dt.timezone(dt.timedelta(hours=8))
            for item in items:
                sym = item.get("symbol")
                price_raw = item.get("price")
                prev_raw = item.get("regularMarketPreviousClose")
                mkt_time = item.get("regularMarketTime")
                if sym and price_raw not in (None, "-", ""):
                    price = float(str(price_raw).replace(",", ""))
                    prev = float(str(prev_raw).replace(",", "")) if prev_raw not in (None, "-", "") else None
                    date_str = None
                    if mkt_time:
                        try:
                            if str(mkt_time).isdigit():
                                ts = int(mkt_time)
                                if ts > 1e11: ts //= 1000
                                date_str = dt.datetime.fromtimestamp(ts, tw_tz).strftime("%Y-%m-%d")
                            else:
                                date_str = str(mkt_time)[:10]
                        except:
                            pass
                    if price > 0:
                        quotes[sym] = {"price": price, "prev_close": prev, "date": date_str, "source": "台灣奇摩API"}
    except Exception as e:
        print(f"[稽核提示] 台灣奇摩 API 備援查詢略過: {e}")
    return quotes

def fetch_twse_official_quotes():
    twse_map = {}
    try:
        r = requests.get("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", headers=HEADERS, timeout=12)
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
                        if p > 0:
                            twse_map[f"{code}.TW"] = {"price": p, "prev_close": None, "date": d_str, "source": "證交所TWSE"}
                    except:
                        pass
    except Exception as e:
        print(f"[稽核提示] 證交所 STOCK_DAY_ALL 查詢略過: {e}")

    try:
        r2 = requests.get("https://openapi.twse.com.tw/v1/exchangeReport/MI_INDEX", headers=HEADERS, timeout=12)
        if r2.status_code == 200:
            for item in r2.json():
                idx_name = item.get("指數", "")
                if "發行量加權股價指數" in idx_name:
                    close_s = (item.get("收盤指數") or "").replace(",", "").strip()
                    if close_s:
                        p = float(close_s)
                        if p > 0:
                            twse_map["^TWII"] = {"price": p, "prev_close": None, "date": None, "source": "證交所TWSE"}
                    break
    except:
        pass
    return twse_map

# ---------------------------------------------------------------------------
# 主資料抓取與三道自動稽核引擎
# ---------------------------------------------------------------------------

def fetch_and_audit_all(symbols, tw_symbols_set, period="2y"):
    closes = {}
    kd_data = {}
    audit_logs = []
    corrected_count = 0

    tw_list = list(tw_symbols_set)
    tw_yahoo_map = fetch_tw_yahoo_quotes(tw_list)
    twse_map = fetch_twse_official_quotes()

    for sym in symbols:
        try:
            df = yf.download(sym, period=period, interval="1d", auto_adjust=False, progress=False)
            if df.empty and sym.endswith(".TW"):
                alt_sym = sym[:-3] + ".TWO"
                df = yf.download(alt_sym, period=period, interval="1d", auto_adjust=False, progress=False)

            if df.empty:
                audit_logs.append(f"⚠️ {sym}: 歷史庫無資料")
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
                alt_info = tw_yahoo_map.get(sym) or twse_map.get(sym)

            if not alt_info:
                try:
                    fi = yf.Ticker(sym).fast_info
                    lp = fi.get("lastPrice") or fi.get("regularMarketPrice")
                    pc = fi.get("previousClose")
                    if lp and float(lp) > 0:
                        alt_info = {"price": float(lp), "prev_close": float(pc) if pc else None, "date": None, "source": "fast_info"}
                except:
                    pass

            if alt_info and alt_info["price"] > 0:
                alt_p = alt_info["price"]
                alt_d = alt_info.get("date")
                alt_prev = alt_info.get("prev_close")

                if alt_d and alt_d > yf_last_date:
                    new_ts = pd.Timestamp(alt_d)
                    if alt_prev and alt_prev > 0 and abs(yf_last_price - alt_prev) / alt_prev > 0.002:
                        df.loc[df.index[-1], "Close"] = alt_prev
                    df.loc[new_ts] = {"High": alt_p, "Low": alt_p, "Close": alt_p}
                    corrected_count += 1
                    print(f"[自動校正] {sym} 補入 {alt_info['source']} 最新收盤 ({alt_d}): {alt_p}")

                elif alt_prev and alt_prev > 0 and abs(yf_last_price - alt_prev) / alt_prev < 0.001 and abs(alt_p - yf_last_price) / yf_last_price > 0.001:
                    tw_today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d")
                    new_ts = pd.Timestamp(tw_today if tw_today > yf_last_date else (df.index[-1] + pd.Timedelta(days=1)))
                    df.loc[new_ts] = {"High": alt_p, "Low": alt_p, "Close": alt_p}
                    corrected_count += 1
                    print(f"[自動校正] {sym} 偵測到停留於昨收，已由 {alt_info['source']} 補入最新收盤: {alt_p}")

                elif alt_d == yf_last_date and abs(alt_p - yf_last_price) / yf_last_price > 0.003:
                    df.loc[df.index[-1], "Close"] = alt_p
                    corrected_count += 1
                    print(f"[自動校正] {sym} 修正同日收盤價差: {yf_last_price} -> {alt_p}")

            closes[sym] = df["Close"]
            kd_data[sym] = kd_timeframes(df)

        except Exception as e:
            print(f"[錯誤] 抓取或稽核 {sym} 失敗: {e}")

    def is_crypto(s):
        return "-USD" in s.upper()

    tw_dates = [closes[s].index[-1] for s in tw_symbols_set if s in closes and len(closes[s]) > 0]
    us_equity_dates = [closes[s].index[-1] for s in symbols if (s not in tw_symbols_set) and (not is_crypto(s)) and s in closes and len(closes[s]) > 0]
    
    max_tw_date = max(tw_dates) if tw_dates else None
    max_us_date = max(us_equity_dates) if us_equity_dates else None

    stale_syms = set()
    for sym, s in closes.items():
        if len(s) == 0 or is_crypto(sym):
            continue
        last_d = s.index[-1]
        ref_d = max_tw_date if sym in tw_symbols_set else max_us_date
        if ref_d is not None and (ref_d - last_d).days >= 1:
            stale_syms.add(sym)
            audit_logs.append(f"⚠️ {sym} 日期落後 ({last_d.strftime('%m/%d')})")

    audit_summary = {
        "corrected": corrected_count,
        "stale_syms": stale_syms,
        "logs": audit_logs,
        "tw_date": max_tw_date.strftime("%Y-%m-%d") if max_tw_date else "—",
        "us_date": max_us_date.strftime("%Y-%m-%d") if max_us_date else "—"
    }
    return closes, kd_data, audit_summary

# ---------------------------------------------------------------------------
# 技術指標與輔助函式
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

def fmt_pct(v):
    if v is None: return "—"
    return f"{'▲' if v>0 else '▼' if v<0 else ''}{abs(v):.1f}%"

# 判定短線燈號邏輯
def get_short_term_status(last, ma5, ma10):
    if pd.isna(last) or pd.isna(ma5) or pd.isna(ma10): return "—"
    if last > ma5 and last > ma10: return "🔥雙線站上"
    if last > ma5 and last < ma10: return "⚡站回五日"
    return "⚠️雙線跌破"

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
            if title is not None and link is not None:
                news.append({"title": title.text, "link": link.text})
    except Exception as e:
        print(f"[錯誤] 抓取新聞失敗: {e}", file=sys.stderr)
        
    if not news:
        news.append({"title": "暫時無法取得最新新聞，請稍後再試。", "link": "#"})
    return news

# ---------------------------------------------------------------------------
# 匯出與報表生成
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
        if s is None or len(s) < 1:
            continue
        
        last = float(s.iloc[-1])
        pct_1d = pct(last, s.iloc[-2]) if len(s) >= 2 else None
        net_1d = (last - s.iloc[-2]) if len(s) >= 2 else None
        pct_1w = pct(last, s.iloc[-6]) if len(s) >= 6 else None
        pct_1m = pct(last, s.iloc[-22]) if len(s) >= 22 else None
        pct_1y = pct(last, s.iloc[-253]) if len(s) >= 253 else None
        
        ma = compute_ma(s)
        ma20w = weekly_ma(s, 20)
        kd = kd_data.get(sym, {})
        kd_d = kd.get("日", (None, None))
        kd_w = kd.get("週", (None, None))
        kd_m = kd.get("月", (None, None))
        
        audit_flag = "⚠️日期延遲" if sym in stale_set else "✅正常"
        short_status = get_short_term_status(last, ma.get("MA5"), ma.get("MA10"))
        
        row = {
            "序號": i, "長線燈號": t["group"], "短線燈號": short_status, "市場": t["market"], "代碼": t["code"],
            "名稱": t["name"], "yf_symbol": sym, "最新收盤": round(last, 2),
            "資料日期": s.index[-1].strftime("%Y-%m-%d"), "稽核狀態": audit_flag,
            "當日漲跌": round(net_1d, 2) if net_1d is not None else None,
            "當日%": round(pct_1d, 2) if pct_1d is not None else None,
            "週%": round(pct_1w, 2) if pct_1w is not None else None,
            "月%": round(pct_1m, 2) if pct_1m is not None else None,
            "年%": round(pct_1y, 2) if pct_1y is not None else None,
            "日K": kd_d[0], "日D": kd_d[1], "週K": kd_w[0], "週D": kd_w[1], "月K": kd_m[0], "月D": kd_m[1],
            "MA5": ma.get("MA5"), "MA10": ma.get("MA10"), 
            "MA20": ma.get("MA20"), "MA60": ma.get("MA60"), "MA240": ma.get("MA240"), "MA20週": ma20w
        }
        rows.append(row)
        
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    
    try:
        with pd.ExcelWriter(REPORT_EXCEL_PATH, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='股市監控報表', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['股市監控報表']
            
            white_bg = workbook.add_format({'bg_color': '#FFFFFF'})
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'center', 'align': 'center',
                'fg_color': '#D7E4BC', 'border': 1
            })
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            worksheet.set_column('A:E', 10, white_bg)
            worksheet.set_column('F:G', 12, white_bg)  
            worksheet.set_column('H:Z', 10, white_bg)  
            
            worksheet.freeze_panes(1, 6)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            tw_format = workbook.add_format({'bg_color': '#FFF9DB'})
            worksheet.conditional_format(1, 0, len(df), len(df.columns) - 1, {
                'type': 'formula',
                'criteria': '=$D2="台股"', 
                'format': tw_format
            })
            
    except Exception as e:
        print(f"[錯誤] 無法產生 Excel 報表: {e}", file=sys.stderr)
        
    return df

def generate_html_report(df, vix_hist, news_data, audit_summary):
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_str = dt.datetime.now(tw_tz).strftime("%Y/%m/%d %H:%M")
    
    idx_df = df[df['長線燈號'] == '大盤']
    us_stocks = df[(df['市場'] == '美股') & (df['長線燈號'] != '大盤')]
    tw_stocks = df[(df['市場'] == '台股') & (df['長線燈號'] != '大盤')]
    
    corr_n = audit_summary.get("corrected", 0)
    stale_n = len(audit_summary.get("stale_syms", set()))
    tw_d = audit_summary.get("tw_date", "—")
    us_d = audit_summary.get("us_date", "—")
    if stale_n == 0:
        audit_banner = f"<div style='background:#e8f5e9; color:#1b8e3a; padding:8px 12px; border-radius:6px; font-size:12px; margin:8px 0 14px 0; font-weight:bold;'>✅ 自動稽核通過｜美股基準日：{us_d}｜台股基準日：{tw_d}（雙重資料源已即時驗證，自動校正 {corr_n} 檔延遲收盤價）</div>"
    else:
        logs_str = "；".join(audit_summary.get("logs", [])[:3])
        audit_banner = f"<div style='background:#fff3e0; color:#ef6c00; padding:8px 12px; border-radius:6px; font-size:12px; margin:8px 0 14px 0; font-weight:bold;'>⚠️ 自動稽核提醒｜美股基準日：{us_d}｜台股基準日：{tw_d}｜已校正 {corr_n} 檔，仍有 {stale_n} 檔日期異常：{logs_str}</div>"

    def fmt_net(val, is_idx=False):
        if pd.isna(val): return '<td class="num mute">—</td>'
        cls = 'up' if val > 0 else 'down' if val < 0 else 'mute'
        mark = '+' if val > 0 else ''
        decimals = 0 if is_idx else 2
        return f'<td class="num {cls}">{mark}{val:,.{decimals}f}</td>'

    def fmt_pct_td(val):
        if pd.isna(val): return '<td class="num mute">—</td>'
        cls = 'up' if val > 0 else 'down' if val < 0 else 'mute'
        mark = '▲' if val > 0 else '▼' if val < 0 else ''
        return f'<td class="num {cls}">{mark}{abs(val):.1f}%</td>'

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
        if last < ma20w:
            return '<span class="badge b-warn">⚠️ 破線空手</span>'
        kd_list = [v for v in [kd_d, kd_w, kd_m] if pd.notna(v)]
        if any(k < 20 for k in kd_list):
            return '<span class="badge b-down">🟢 日K超跌</span>'
        if any(k > 80 for k in kd_list):
            return '<span class="badge b-up">🔴 熱K>80</span>'
        return '<span class="badge b-safe">✓ 線上安全</span>'
    
    def get_short_badge(status_text):
        if status_text == "🔥雙線站上": return '<span class="badge b-up">🔥雙線站上</span>'
        if status_text == "⚡站回五日": return '<span class="badge" style="background:#e3f2fd; color:#1976d2;">⚡站回五日</span>'
        if status_text == "⚠️雙線跌破": return '<span class="badge b-down">⚠️雙線跌破</span>'
        return '<span class="mute">—</span>'
        
    def get_dist_pct(last, ma20w):
        if pd.isna(last) or pd.isna(ma20w) or ma20w == 0: return "—"
        dist = (last - ma20w) / ma20w * 100
        mark = '+' if dist > 0 else ''
        return f"{mark}{dist:.1f}%"

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
a {{ color:#1456b8; text-decoration:none; font-weight:bold; }} a:hover {{ text-decoration:underline; }}
.badge {{ padding: 3px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block; }}
.b-up {{ background: #ffebee; color: #d32f2f; }}
.b-down {{ background: #e8f5e9; color: #1b8e3a; }}
.b-warn {{ background: #fff3e0; color: #ef6c00; }}
.b-safe {{ background: #f0f3f7; color: #333; }}
.foot {{ color:var(--mute); font-size:11px; margin-top:14px; text-align:center; }}
.action-btn {{ display:inline-block; padding:10px 20px; border-radius:4px; text-decoration:none; font-size:14px; font-weight:bold; text-align:center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); transition: opacity 0.2s; margin-top: 10px; }}
.action-btn:hover {{ opacity: 0.8; text-decoration:none; }}
.btn-download {{ background:#1b8e3a; color:#fff; width: 100%; max-width: 300px; }}

.freeze-single {{
    position: -webkit-sticky; position: sticky; left: 0; z-index: 2;
    background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2);
}}
.freeze-idx {{
    position: -webkit-sticky; position: sticky; left: 0; z-index: 2;
    background: #fff; min-width: 25px; max-width: 25px; 
    padding-left: 4px; padding-right: 4px;
}}
.freeze-name {{
    position: -webkit-sticky; position: sticky; left: 25px; z-index: 2; 
    background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2);
    min-width: 70px; max-width: 90px; 
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    padding-left: 6px; padding-right: 6px;
}}
thead th.freeze-single, thead th.freeze-idx, thead th.freeze-name {{
    background: #eef1f5; z-index: 3;
}}

tr.tw-row td {{ background-color: #fff9db !important; }}
tr.tw-row td.freeze-single, tr.tw-row td.freeze-idx, tr.tw-row td.freeze-name {{ background-color: #fff9db !important; }}
</style></head><body><div class='wrap'>
<h1>📊 股市核心監控報表</h1><div class='sub'>產生時間：{now_str} (台灣時間)</div>
{audit_banner}

<h2>大盤與 VIX 狀況 (+/- 2%)</h2>
<div class='table-container'><table>
<thead><tr>
    <th class='freeze-single'>指數</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>當日燈號</th><th class='num'>股價位階</th><th class='num'>長線燈號</th><th class='num'>資料日期</th>
</tr></thead><tbody>
"""
    
    for _, row in idx_df.iterrows():
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma20, ma60, ma240, ma20w = row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        pos = get_position(last, ma20, ma60, ma240)
        weather = get_weather_icon(pct_val, 2.0)
        status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        d_tag = row['資料日期'][5:] + ("⚠️" if "⚠️" in row['稽核狀態'] else "")
        
        html += f"<tr class='{bg_cls}'><td class='freeze-single'><a href='https://finance.yahoo.com/quote/{row['yf_symbol']}' target='_blank'>{row['名稱']}</a></td>"
        html += f"<td class='num'>{last:,.0f}</td>"
        html += fmt_net(net, is_idx=True)
        html += fmt_pct_td(pct_val)
        html += f"<td class='num'>{weather}</td><td class='num'>{pos}</td><td class='num'>{status}</td><td class='num mute'>{d_tag}</td></tr>"

    for v_name, v_sym in [("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")]:
        h = vix_hist.get(v_sym, [0])
        val = h[0]
        if val is None or val <= 0:
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>暫無報價</td><td class='num mute'>—</td></tr>"
        else:
            status = "<span class='badge b-down'>🟢 安全</span>" if val < 35 else "<span class='badge b-up'>🔴 恐慌</span>"
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num'>{val:.2f}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num'>{status}</td><td class='num mute'>—</td></tr>"
            
    html += "</tbody></table></div>"
    
    html += "<h2>個股追蹤清單 (支援短線/長線雙燈號)</h2><div class='table-container'><table><thead><tr>"
    html += "<th class='num freeze-idx'>#</th><th class='freeze-name'>名稱</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>短線燈號</th><th class='num'>5MA/10MA</th><th class='num'>長線燈號</th><th class='num'>日/週/月 K值</th><th class='num'>MA20週均</th><th class='num'>距20週均(%)</th><th class='num'>資料日期</th></tr></thead><tbody>"

    def render_stock_row(row, idx):
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma5, ma10, ma20, ma60, ma240, ma20w = row['MA5'], row['MA10'], row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        long_status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        short_status_text = get_short_term_status(last, ma5, ma10)
        short_badge = get_short_badge(short_status_text)
        
        dist_str = get_dist_pct(last, ma20w)
        
        kd_str = f"{int(kd_d) if pd.notna(kd_d) else '-'}/{int(kd_w) if pd.notna(kd_w) else '-'}/{int(kd_m) if pd.notna(kd_m) else '-'}"
        ma20w_str = f"{ma20w:,.2f}" if pd.notna(ma20w) else "—"
        ma5_10_str = f"{ma5:,.2f} / {ma10:,.2f}" if pd.notna(ma5) and pd.notna(ma10) else "—"
        
        url = f"https://finance.yahoo.com/quote/{row['yf_symbol']}"
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        d_tag = row['資料日期'][5:] + ("⚠️" if "⚠️" in row['稽核狀態'] else "")
        
        res = f"<tr class='{bg_cls}'><td class='num idx freeze-idx'>{idx}</td><td class='name freeze-name'><a href='{url}' target='_blank'>{row['名稱']}</a></td><td class='num'>{last:,.2f}</td>"
        res += fmt_net(net, is_idx=False) + fmt_pct_td(pct_val)
        res += f"<td class='num'>{short_badge}</td><td class='num'>{ma5_10_str}</td><td class='num'>{long_status}</td><td class='num'>{kd_str}</td><td class='num'>{ma20w_str}</td><td class='num { 'up' if '+' in dist_str else 'down' }'>{dist_str}</td><td class='num mute'>{d_tag}</td></tr>"
        return res

    html += "<tr class='grp'><td colspan='12'>美股觀察區</td></tr>"
    idx = 1
    for _, row in us_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "<tr class='grp'><td colspan='12'>台股觀察區</td></tr>"
    for _, row in tw_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "</tbody></table></div>"

    # ================= 專屬短線與位階警示區塊 (已移除粗體) =================
    html += "<h2>四、其他位階警示</h2>"
    html += "<ol style='line-height:1.8; font-size:14px; padding-left:24px; margin-bottom:24px;'>"
    
    alert_df = df[df['長線燈號'].isin(['大盤', '核心'])]
    
    alert_count_html = 0
    for _, r in alert_df.iterrows():
        last = r['最新收盤']
        ma60 = r['MA60']
        ma240 = r['MA240']
        
        drop = []
        # 跌破季線標綠色 🟢，跌破年線標紅色 🔴 (一般字體，不加粗)
        if pd.notna(ma60) and last < ma60: drop.append("<span style='color:#1b8e3a;'>🟢[跌]破季線</span>")
        if pd.notna(ma240) and last < ma240: drop.append("<span style='color:#d32f2f;'>🔴[跌]破年線</span>")
        
        w_ret, m_ret = r['週%'], r['月%']
        if (pd.notna(w_ret) and w_ret <= -20) or (pd.notna(m_ret) and m_ret <= -20): 
            drop.append("<span style='color:#1b8e3a;'>🟢週/月跌>20%</span>")
            
        ko80 = []
        if pd.notna(r['日K']) and r['日K'] > 80: ko80.append("日")
        if pd.notna(r['週K']) and r['週K'] > 80: ko80.append("週")
        if pd.notna(r['月K']) and r['月K'] > 80: ko80.append("月")
        
        if drop or ko80:
            li_html = f"<li>{r['名稱']}: "
            if drop: li_html += " ".join(drop) + " "
            if ko80: li_html += f"<span style='color:#d32f2f;'>🔴[熱]{'/'.join(ko80)}K>80</span>"
            li_html += "</li>"
            html += li_html
            alert_count_html += 1
            
    if alert_count_html == 0:
        html += "<li>*(今日無觸發任何位階警示)*</li>"
        
    html += "</ol>"

    # ================= 市場新聞區塊 =================
    html += "<h2>市場新聞 ( Yahoo 財經 )</h2>"
    html += "<ol style='line-height:1.8; font-size:14px; padding-left:24px; margin-bottom:24px;'>"
    for item in news_data:
        html += f"<li><a href='{item['link']}' target='_blank' style='font-weight:normal;'>{item['title']}</a></li>"
    html += "</ol>"

    # ================= 沉底備註區塊 =================
    html += """
    <div class="notes" style="margin-top: 24px; font-size: 12px; color: #666; background: #fcfcfc; padding: 16px; border-radius: 6px; border: 1px solid var(--line);">
        <b style="font-size: 13px; color: #333;">📝 燈號與位階判斷標準</b>
        <div style="display: flex; flex-wrap: wrap; gap: 16px; margin-top: 10px; line-height: 1.7;">
            <div style="flex: 1 1 280px;">
                <b style="color:#1f3b66;">1. 短線燈號 (5MA/10MA)</b><br>
                🔥 <b>雙線站上</b>：收盤 ＞ 5MA 且 ＞ 10MA (短線強勢)<br>
                ⚡ <b>站回五日</b>：收盤 ＞ 5MA 但 ＜ 10MA (止跌轉強)<br>
                ⚠️ <b>雙線跌破</b>：收盤 ＜ 5MA 且 ＜ 10MA (短線弱勢)
            </div>
            <div style="flex: 1 1 280px;">
                <b style="color:#1f3b66;">2. 長線燈號 (20週均線)</b><br>
                ⚠️ <b>破線空手</b>：跌破 20 週均線 (長線趨勢轉弱)<br>
                🟢 <b>K值超跌</b>：20 週線上，且任一 K值 ＜ 20<br>
                🔴 <b>熱K>80</b>：20 週線上，且任一 K值 ＞ 80<br>
                ✓ <b>線上安全</b>：20 週線上，且 KD 無極端狀況
            </div>
            <div style="flex: 1 1 220px;">
                <b style="color:#1f3b66;">3. 其他位階與當日燈號</b><br>
                <span style="color:#1b8e3a;">🟢[跌]破季線</span>：跌破 60MA (中線轉弱)<br>
                <span style="color:#d32f2f;">🔴[跌]破年線</span>：跌破 240MA (長線警戒)<br>
                ☀️ 大漲(大盤≥2%/個股≥5%) ｜ ⛈️ 大跌
            </div>
        </div>
    </div>
    """

    # ================= 專屬 Excel 按鈕區塊 =================
    html += """
    <div class='foot' style='margin-top:20px; font-size:12px;'>
        資料來源：台灣奇摩股市 / 證交所 / Yahoo Finance ｜ 報表產生器：V8.2 清爽字體版<br><br>
        <div style="display:flex; justify-content:center; align-items:center;">
            <a href='report_data.xlsx' download='stock_monitor.xlsx' class='action-btn btn-download'>📥 下載完整 Excel (.xlsx) 報表</a>
        </div>
    </div>
    </div></body></html>
    """
    
    with open(REPORT_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)

# ---------------------------------------------------------------------------
# 組裝 LINE Flex Message 卡片
# ---------------------------------------------------------------------------
def build_flex_message(closes, watch, kd_data, vix_hist, audit_summary, df_report):
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    today = dt.datetime.now(tw_tz)
    wd = "一二三四五六日"[today.weekday()]
    
    C_UP = "#d32f2f"
    C_DOWN = "#1b8e3a"
    C_FLAT = "#8a9099"
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
            fmt_p = f"{last:,.0f}" if is_idx else f"{last:,.1f}"
            fmt_d = f"{diff:+,.0f}" if is_idx else f"{diff:+,.1f}"
            
            sec1.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": f"{item_idx}. {name} {fmt_p} ("},
                    {"type": "span", "text": f"{fmt_d} ; {fmt_pct(dchg)}", "color": val_color, "weight": "bold"},
                    {"type": "span", "text": f") {weather}{stale_mark}"} 
                ]
            })
            item_idx += 1
            return True
        return False

    for sym, name in INDICES:
        add_sec1_item(name, sym, closes.get(sym), True)
        
    sec1.append({"type": "text", "text": "-" * 15, "size": "sm", "color": "#8a9099"})
    
    core_items = []
    seen_core = set()
    s_twii = closes.get("^TWII")
    
    if add_sec1_item("台大盤", "^TWII", s_twii, True):
        core_items.append({"name": "台大盤", "yf": "^TWII", "last": float(s_twii.iloc[-1]), "ma20w": weekly_ma(s_twii, 20)})
        
    for w in watch:
        sym = w["yf_symbol"]
        if w["market"] == "台股" and (w["group"] == "核心" or "2330" in sym or "0050" in sym):
            if "正2" in w["name"] or "反1" in w["name"] or sym in seen_core:
                continue
            seen_core.add(sym)
            s = closes.get(sym)
            if add_sec1_item(w["name"], sym, s, False):
                core_items.append({"name": w["name"], "yf": sym, "last": float(s.iloc[-1]), "ma20w": weekly_ma(s, 20)})
                
    contents.extend(sec1)
    contents.append({"type": "separator", "margin": "md"})
    
    sec2 = [{"type": "text", "text": "▌二、VIX 恐慌指數 (<35)", "size": "sm", "color": "#333333"}]
    for i, (v_name, v_sym) in enumerate([("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")], 1):
        h = vix_hist.get(v_sym, [0, 0, 0, 0])
        val_today, val_yest, val_prev_w = h[0], h[1], h[3]
        if val_today is None or val_today <= 0:
            sec2.append({"type": "text", "text": f"{i}. {v_name} — (暫無報價)", "size": "sm", "color": "#8a9099"})
        else:
            status = "🟢安全" if val_today < 35 else "🔴恐慌"
            c_color = C_DOWN if val_today < 35 else C_UP
            yest_str = f"昨日{val_yest:.1f}" if (val_yest and val_yest > 0) else "昨日—"
            prev_w_str = f"前一週{val_prev_w:.1f}" if (val_prev_w and val_prev_w > 0) else "前一週—"
            sec2.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": f"{i}. {v_name} {val_today:.1f}("},
                    {"type": "span", "text": status, "color": c_color, "weight": "bold"},
                    {"type": "span", "text": f"), {yest_str}, {prev_w_str}"}
                ]
            })
    contents.extend(sec2)
    contents.append({"type": "separator", "margin": "md"})
    
    sec3 = [{"type": "text", "text": "▌三、台股核心指標 (長/短線燈號)", "size": "sm", "color": "#333333"}]
    for i, item in enumerate(core_items, 1):
        last = item["last"]
        ma20w = item["ma20w"]
        
        df_row = df_report[df_report['yf_symbol'] == item["yf"]]
        if df_row.empty: continue
        r = df_row.iloc[0]
        
        kd_d, kd_w, kd_m = r['日K'], r['週K'], r['月K']
        short_status = r['短線燈號']
        short_color = C_UP if "站上" in short_status else C_DOWN if "跌破" in short_status else "#1976d2"
        
        price_fmt = f"({last:,.0f})" if item["yf"] == "^TWII" else f"({last:,.1f})"
        sec3.append({"type": "text", "text": f"{i}. {item['name']} {price_fmt}", "size": "sm", "color": "#333333", "margin": "sm"})
        sec3.append({
            "type": "text", "size": "sm", "wrap": True,
            "contents": [
                {"type": "span", "text": f"   K值: 日{int(kd_d) if pd.notna(kd_d) else '-'}/週{int(kd_w) if pd.notna(kd_w) else '-'}/月{int(kd_m) if pd.notna(kd_m) else '-'}｜短線: ", "color": "#666666"},
                {"type": "span", "text": short_status, "color": short_color}
            ]
        })
        
        if last and ma20w:
            dist = (last - ma20w) / ma20w * 100
            if dist < 0: status = "⚠️破線空手"
            else:
                if pd.notna(kd_d) and kd_d < 20: status = "🟢日K超跌"
                elif pd.notna(kd_w) and kd_w < 20: status = "🟢週K超跌"
                elif pd.notna(kd_m) and kd_m < 20: status = "🟢月K超跌"
                else: status = "✓線上安全"
            
            ma_fmt = f"{ma20w:,.0f}" if item["yf"] == "^TWII" else f"{ma20w:,.1f}"
            dist_color = C_UP if dist > 0 else C_DOWN if dist < 0 else C_FLAT
            
            sec3.append({
                "type": "text", "size": "sm", "wrap": True,
                "contents": [
                    {"type": "span", "text": f"   長線距20週({ma_fmt}): "},
                    {"type": "span", "text": f"{dist:+.1f}%", "color": dist_color, "weight": "bold"},
                    {"type": "span", "text": f" ({status})"}
                ]
            })
    contents.extend(sec3)
    contents.append({"type": "separator", "margin": "md"})
    
    # ================= 四、其他位階警示 (已移除粗體) =================
    sec4 = [{"type": "text", "text": "▌四、其他位階警示", "size": "sm", "color": "#333333"}]
    alerts = []
    alert_df = df_report[df_report['長線燈號'].isin(['大盤', '核心'])]
    alert_count = 1
    for _, r in alert_df.iterrows():
        last = r['最新收盤']
        ma60 = r['MA60']
        ma240 = r['MA240']
        
        drop_q = (pd.notna(ma60) and last < ma60)
        drop_y = (pd.notna(ma240) and last < ma240)
        w_ret, m_ret = r['週%'], r['月%']
        drop_20 = (pd.notna(w_ret) and w_ret <= -20) or (pd.notna(m_ret) and m_ret <= -20)
        
        ko80 = []
        if pd.notna(r['日K']) and r['日K'] > 80: ko80.append("日")
        if pd.notna(r['週K']) and r['週K'] > 80: ko80.append("週")
        if pd.notna(r['月K']) and r['月K'] > 80: ko80.append("月")
        
        if drop_q or drop_y or drop_20 or ko80:
            spans = [{"type": "span", "text": f"{alert_count}. {r['名稱']}: "}]
            # 跌破季線標綠色 🟢，跌破年線標紅色 🔴 (一般字體，不加粗)
            if drop_q:
                spans.append({"type": "span", "text": "🟢[跌]破季線 ", "color": C_DOWN})
            if drop_y:
                spans.append({"type": "span", "text": "🔴[跌]破年線 ", "color": C_UP})
            if drop_20:
                spans.append({"type": "span", "text": "🟢週/月跌>20% ", "color": C_DOWN})
            if ko80:
                spans.append({"type": "span", "text": f"🔴[熱]{'/'.join(ko80)}K>80", "color": C_UP})
            alerts.append({"type": "text", "size": "sm", "wrap": True, "contents": spans})
            alert_count += 1
            
    if alerts: sec4.extend(alerts)
    else: sec4.append({"type": "text", "text": "*(今日無觸發)*", "size": "sm", "color": "#8a9099"})
    contents.extend(sec4)
    
    # 底部極簡名詞備註
    contents.append({"type": "separator", "margin": "md"})
    contents.append({
        "type": "text",
        "text": "💡 短線: 🔥雙線站上(>5/10MA)｜⚡站回五日(>5MA)｜⚠️雙線跌破\n💡 長線: 20週均上為✓安全，破線為⚠️空手｜🟢破季線 🔴破年線",
        "size": "xxs", "color": "#8a9099", "wrap": True
    })
    
    contents.append({
        "type": "button", "style": "primary", "color": "#1456b8", "height": "sm", "margin": "sm",
        "action": { "type": "uri", "label": "👉 開啟詳細網頁報表", "uri": os.environ.get("GITHUB_PAGES_URL", "").strip() or "https://ckliu60.github.io/stock-monitor/" }
    })

    return {
        "type": "flex",
        "altText": f"📊 股市快報 {today.month:02d}/{today.day:02d} 已送達",
        "contents": {
            "type": "bubble",
            "size": "mega",
            "body": {
                "type": "box",
                "layout": "vertical",
                "spacing": "md",
                "contents": contents
            }
        }
    }

def main():
    force_dry_run = "--dry-run" in sys.argv
    watch = load_watchlist()
    
    idx_syms = [s for s, _ in INDICES]
    stk_syms = [w["yf_symbol"] for w in watch]
    all_syms = list(dict.fromkeys(idx_syms + ["^TWII"] + stk_syms))
    tw_symbols_set = set(["^TWII"] + [w["yf_symbol"] for w in watch if w["market"] == "台股"])
    
    closes, kd_data, audit_summary = fetch_and_audit_all(all_syms, tw_symbols_set)

    vix_hist = {}
    v_closes, _, _ = fetch_and_audit_all(["^VIX", "^VIXTWN"], set(), period="1mo")
    for v_sym in ["^VIX", "^VIXTWN"]:
        s = v_closes.get(v_sym)
        if s is not None and len(s) >= 6:
            vix_hist[v_sym] = [float(s.iloc[-1]), float(s.iloc[-2]), float(s.iloc[-3]), float(s.iloc[-6])]
        elif s is not None and len(s) > 0:
            vix_hist[v_sym] = [float(s.iloc[-1]), 0, 0, 0]
        else:
            vix_hist[v_sym] = [0, 0, 0, 0]

    news_data = fetch_yahoo_news(limit=5)
    df_report = export_all_data(closes, watch, kd_data, audit_summary)
    generate_html_report(df_report, vix_hist, news_data, audit_summary)
    flex_message = build_flex_message(closes, watch, kd_data, vix_hist, audit_summary, df_report)

    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
    uid = os.environ.get("LINE_USER_ID", "").strip()
    
    if force_dry_run or not token or not uid:
        print("[DRY-RUN] 訊息預覽 JSON：\n" + json.dumps(flex_message, indent=2, ensure_ascii=False) + "\n")
        return 0

    resp = requests.post(LINE_PUSH_URL, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                         json={"to": uid, "messages": [flex_message]}, timeout=20)
    
    try:
        resp.raise_for_status()
        print("[OK] 已成功推播卡片訊息至 LINE。")
    except Exception as e:
        print(f"[錯誤] LINE 推播失敗: {e}\n{resp.text}", file=sys.stderr)
        
    return 0

if __name__ == "__main__":
    sys.exit(main())
