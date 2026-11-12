#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (V6.2 雲端全自動版 - 名稱縮寫微調版)
---------------------------------------------------------------------------
更新日誌 (V6.2): 
- 取消「富邦科技」的縮寫，保留原名以利辨識。
- 將「台積電ADR」縮寫改為「台積電R」。
- 保留 V6.1 的手機版凍結窗格欄寬最佳化與台股高亮功能。
"""

import os
import sys
import io
import csv
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

ALERT_TARGETS = [
    "道瓊", "那指", "標普", "費半",
    "台積電ADR", "輝達", "特斯拉", "SpaceX",
    "台積電", "聯發科", "元大台灣50"
]

MA_WINDOWS = [("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]
LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRa7RG_nKx8DTBSsOsBPz4MQZPw_mYByDYy7hux2MwuhqRn1VIvqr5XWiZttrW24A/pub?gid=1903088963&single=true&output=csv"

HERE = Path(__file__).resolve().parent
PUBLIC_DIR = HERE / "public"
REPORT_HTML_PATH = PUBLIC_DIR / "index.html"
REPORT_EXCEL_PATH = PUBLIC_DIR / "report_data.xlsx"

# ---------------------------------------------------------------------------
# 工具函式
# ---------------------------------------------------------------------------

def shorten_name(name):
    """將過長的股票名稱縮短，以節省手機顯示空間"""
    n = name.strip()
    if "群益台灣精選高息" in n: return "00919"
    if "元大台灣50" in n: return "0050"
    if n == "台積電ADR": return "台積電R"
    return n

def load_watchlist():
    try:
        resp = requests.get(SHEET_CSV_URL, timeout=20)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        text = resp.text
    except Exception as e:
        print(f"[錯誤] 讀取雲端試算表失敗: {e}", file=sys.stderr)
        return []

    rows = []
    for raw in csv.DictReader(io.StringIO(text)):
        code = (raw.get("code") or "").strip()
        yf_symbol = (raw.get("yf_symbol") or "").strip()
        market = (raw.get("market") or "").strip()
        
        if not yf_symbol:
            if market == "台股" or code.isdigit():
                yf_symbol = f"{code}.TW"
            else:
                yf_symbol = code.replace(".", "-")
        else:
            if market == "美股":
                yf_symbol = yf_symbol.replace(".", "-")

        original_name = (raw.get("name") or yf_symbol).strip()
        rows.append({
            "market": market, "code": code, "yf_symbol": yf_symbol,
            "name": shorten_name(original_name), # 載入時直接縮短名稱
            "group": (raw.get("group") or "觀察").strip()
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

def fetch_yahoo_news(limit=5):
    url = "https://tw.news.yahoo.com/rss/finance"
    news = []
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(url, headers=headers, timeout=10)
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
def export_all_data(closes, watch, kd_data):
    all_targets = [{"market": "美股", "code": s, "name": n, "yf": s, "group": "核心"} for s, n in INDICES] + \
                  [{"market": "台股", "code": "^TWII", "name": "台股大盤", "yf": "^TWII", "group": "核心"}] + \
                  [{"market": w["market"], "code": w["code"], "name": w["name"], "yf": w["yf_symbol"], "group": w["group"]} for w in watch]
    
    rows = []
    for i, t in enumerate(all_targets, 1):
        sym = t["yf"]
        s = closes.get(sym)
        if s is None or len(s) < 2: continue
        
        last = s.iloc[-1]
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
        
        row = {
            "序號": i, "群組": t["group"], "市場": t["market"], "代碼": t["code"],
            "名稱": t["name"], "yf_symbol": sym, "最新收盤": round(last, 2), "資料日期": s.index[-1].strftime("%Y-%m-%d"),
            "當日漲跌": round(net_1d, 2) if net_1d is not None else None,
            "當日%": round(pct_1d, 2) if pct_1d is not None else None,
            "週%": round(pct_1w, 2) if pct_1w is not None else None,
            "月%": round(pct_1m, 2) if pct_1m is not None else None,
            "年%": round(pct_1y, 2) if pct_1y is not None else None,
            "日K": kd_d[0], "日D": kd_d[1], "週K": kd_w[0], "週D": kd_w[1], "月K": kd_m[0], "月D": kd_m[1],
            "MA10": ma.get("MA10"), "MA20": ma.get("MA20"), "MA60": ma.get("MA60"), "MA240": ma.get("MA240"), "MA20週": ma20w
        }
        rows.append(row)
        
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    
    try:
        with pd.ExcelWriter(REPORT_EXCEL_PATH, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='股市監控報表', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['股市監控報表']
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'center', 'align': 'center',
                'fg_color': '#D7E4BC', 'border': 1
            })
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Excel 欄寬配合名稱縮短，稍微調小
            worksheet.set_column('E:E', 12)  
            worksheet.set_column('F:F', 12)  
            worksheet.set_column('G:Y', 10)  
            
            worksheet.freeze_panes(1, 5)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            tw_format = workbook.add_format({'bg_color': '#FFF9DB'})
            worksheet.conditional_format(1, 0, len(df), len(df.columns) - 1, {
                'type': 'formula',
                'criteria': '=$C2="台股"', 
                'format': tw_format
            })
            
    except Exception as e:
        print(f"[錯誤] 無法產生 Excel 報表: {e}", file=sys.stderr)
        
    return df

def generate_html_report(df, vix_hist, news_data):
    tw_tz = dt.timezone(dt.timedelta(hours=8))
    now_str = dt.datetime.now(tw_tz).strftime("%Y/%m/%d %H:%M")
    
    idx_df = df[df['名稱'].isin([n for _, n in INDICES])]
    tw_idx = df[df['名稱'] == '台股大盤']
    us_stocks = df[(df['市場'] == '美股') & (~df['名稱'].isin([n for _, n in INDICES]))]
    tw_stocks = df[(df['市場'] == '台股') & (df['名稱'] != '台股大盤')]
    
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

/* === 縮減欄寬的凍結窗格 (Sticky Column) === */
.freeze-single {{
    position: -webkit-sticky; position: sticky; left: 0; z-index: 2;
    background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2);
}}
.freeze-idx {{
    position: -webkit-sticky; position: sticky; left: 0; z-index: 2;
    background: #fff; min-width: 25px; max-width: 25px; /* 縮減序號欄寬 */
    padding-left: 4px; padding-right: 4px;
}}
.freeze-name {{
    position: -webkit-sticky; position: sticky; left: 25px; z-index: 2; /* 緊貼序號欄 */
    background: #fff; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.2);
    min-width: 70px; max-width: 90px; /* 縮小名稱欄寬，配合名稱縮短 */
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    padding-left: 6px; padding-right: 6px;
}}
thead th.freeze-single, thead th.freeze-idx, thead th.freeze-name {{
    background: #eef1f5; z-index: 3;
}}

/* === 台股高亮淺黃底色 (雙大括號) === */
tr.tw-row td {{ background-color: #fff9db !important; }}
tr.tw-row td.freeze-single, tr.tw-row td.freeze-idx, tr.tw-row td.freeze-name {{ background-color: #fff9db !important; }}
</style></head><body><div class='wrap'>
<h1>📊 股市核心監控報表</h1><div class='sub'>產生時間：{now_str} (台灣時間)</div>

<h2>大盤與 VIX 狀況 (+/- 2%)</h2>
<div class='table-container'><table>
<thead><tr>
    <th class='freeze-single'>指數</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>當日燈號</th><th class='num'>股價位階</th><th class='num'>位階燈號</th>
</tr></thead><tbody>
"""
    
    for _, row in pd.concat([idx_df, tw_idx]).iterrows():
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma20, ma60, ma240, ma20w = row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        pos = get_position(last, ma20, ma60, ma240)
        weather = get_weather_icon(pct_val, 2.0)
        status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        
        html += f"<tr class='{bg_cls}'><td class='freeze-single'><a href='https://finance.yahoo.com/quote/{row['yf_symbol']}' target='_blank'>{row['名稱']}</a></td>"
        html += f"<td class='num'>{last:,.0f}</td>"
        html += fmt_net(net, is_idx=True)
        html += fmt_pct_td(pct_val)
        html += f"<td class='num'>{weather}</td><td class='num'>{pos}</td><td class='num'>{status}</td></tr>"

    for v_name, v_sym in [("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")]:
        h = vix_hist.get(v_sym, [0])
        val = h[0]
        if val is None or val <= 0:
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>暫無報價</td></tr>"
        else:
            status = "<span class='badge b-down'>🟢 安全</span>" if val < 35 else "<span class='badge b-up'>🔴 恐慌</span>"
            html += f"<tr><td class='freeze-single'>{v_name}</td><td class='num'>{val:.2f}</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num mute'>—</td><td class='num'>{status}</td></tr>"
            
    html += "</tbody></table></div>"
    
    html += "<h2>個股追蹤清單</h2><div class='table-container'><table><thead><tr>"
    html += "<th class='num freeze-idx'>#</th><th class='freeze-name'>名稱</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>漲跌幅(%)</th><th class='num'>當日燈號</th><th class='num'>股價位階</th><th class='num'>位階燈號</th><th class='num'>日/週/月 K值</th><th class='num'>MA20週均</th><th class='num'>距20週均(%)</th></tr></thead><tbody>"

    def render_stock_row(row, idx):
        last, net, pct_val = row['最新收盤'], row['當日漲跌'], row['當日%']
        ma20, ma60, ma240, ma20w = row['MA20'], row['MA60'], row['MA240'], row['MA20週']
        kd_d, kd_w, kd_m = row['日K'], row['週K'], row['月K']
        
        pos = get_position(last, ma20, ma60, ma240)
        weather = get_weather_icon(pct_val, 5.0) 
        status = get_status_badge(last, ma20w, kd_d, kd_w, kd_m)
        dist_str = get_dist_pct(last, ma20w)
        
        kd_str = f"{int(kd_d) if pd.notna(kd_d) else '-'}/{int(kd_w) if pd.notna(kd_w) else '-'}/{int(kd_m) if pd.notna(kd_m) else '-'}"
        ma20w_str = f"{ma20w:,.2f}" if pd.notna(ma20w) else "—"
        url = f"https://finance.yahoo.com/quote/{row['yf_symbol']}"
        
        bg_cls = "tw-row" if row['市場'] == '台股' else ""
        
        res = f"<tr class='{bg_cls}'><td class='num idx freeze-idx'>{idx}</td><td class='name freeze-name'><a href='{url}' target='_blank'>{row['名稱']}</a></td><td class='num'>{last:,.2f}</td>"
        res += fmt_net(net, is_idx=False) + fmt_pct_td(pct_val)
        res += f"<td class='num'>{weather}</td><td class='num'>{pos}</td><td class='num'>{status}</td><td class='num'>{kd_str}</td><td class='num'>{ma20w_str}</td><td class='num { 'up' if '+' in dist_str else 'down' }'>{dist_str}</td></tr>"
        return res

    html += "<tr class='grp'><td colspan='11'>美股觀察區</td></tr>"
    idx = 1
    for _, row in us_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "<tr class='grp'><td colspan='11'>台股觀察區</td></tr>"
    for _, row in tw_stocks.iterrows():
        html += render_stock_row(row, idx)
        idx += 1
        
    html += "</tbody></table></div>"

    # ================= 市場新聞區塊 =================
    html += "<h2>市場新聞 ( Yahoo 財經 )</h2>"
    html += "<ol style='line-height:1.8; font-size:14px; padding-left:24px; margin-bottom:24px;'>"
    for item in news_data:
        html += f"<li><a href='{item['link']}' target='_blank' style='font-weight:normal;'>{item['title']}</a></li>"
    html += "</ol>"

    # ================= 沉底備註區塊 (並排縮減高度版) =================
    html += """
    <div class="notes" style="margin-top: 24px; font-size: 12px; color: #666; background: #fcfcfc; padding: 16px; border-radius: 6px; border: 1px solid var(--line);">
        <b style="font-size: 13px; color: #333;">📝 燈號與位階判斷標準</b>
        <div style="display: flex; flex-wrap: wrap; gap: 16px; margin-top: 10px; line-height: 1.7;">
            <div style="flex: 1 1 300px;">
                <b style="color:#1f3b66;">1. 當日燈號 (大盤 ±2% / 個股 ±5%)</b><br>
                ☀️ <b>大晴天</b>：大盤漲 >= 2.0% / 個股漲 >= 5.0%<br>
                🌤️ <b>晴天</b>：漲幅 > 0% &nbsp;&nbsp;&nbsp; ☁️ <b>平盤</b>：漲幅 = 0%<br>
                🌧️ <b>雨天</b>：跌幅 < 0% &nbsp;&nbsp;&nbsp; ⛈️ <b>雷雨</b>：大盤 <= -2% / 個股 <= -5%
            </div>
            <div style="flex: 1 1 300px;">
                <b style="color:#1f3b66;">2. 位階燈號 (大盤與個股通用)</b><br>
                ⚠️ <b>破線空手</b>：跌破 20 週均線（長線趨勢轉弱）<br>
                🟢 <b>K值超跌</b>：20 週線上, 且任一 K值 < 20<br>
                🔴 <b>熱K>80</b>：20 週線上, 且任一 K值 > 80<br>
                ✓ <b>線上安全</b>：20 週線上, 且 KD 數值無極端狀況
            </div>
            <div style="flex: 1 1 180px;">
                <b style="color:#1f3b66;">3. 股價位階</b><br>
                依據收盤價落點分為：<br>
                <span style="display:inline-block; background:#eef1f5; padding:2px 6px; border-radius:4px; margin-top:4px;">月線上</span>
                <span style="display:inline-block; background:#eef1f5; padding:2px 6px; border-radius:4px; margin-top:4px;">月~季</span><br>
                <span style="display:inline-block; background:#eef1f5; padding:2px 6px; border-radius:4px; margin-top:4px;">季~年</span>
                <span style="display:inline-block; background:#eef1f5; padding:2px 6px; border-radius:4px; margin-top:4px;">年線下</span>
            </div>
        </div>
    </div>
    """

    # ================= 專屬 Excel 按鈕區塊 =================
    html += """
    <div class='foot' style='margin-top:20px; font-size:12px;'>
        資料來源：Yahoo Finance ｜ 報表產生器：V6.2 名稱縮寫微調版<br><br>
        <div style="display:flex; justify-content:center; align-items:center;">
            <a href='report_data.xlsx' target='_blank' class='action-btn btn-download'>📥 下載完整 Excel (.xlsx) 報表</a>
        </div>
    </div>
    </div></body></html>
    """
    
    with open(REPORT_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------------------------------------------------------------------
# 組裝 LINE Flex Message 卡片
# ---------------------------------------------------------------------------
def build_flex_message(closes, watch, kd_data, vix_hist):
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    
    sec1 = ["▌一、大盤與核心狀況 (+/- 2%)"]
    for i, (sym, name) in enumerate(INDICES, 1):
        s = closes.get(sym)
        if s is not None and len(s) >= 2:
            last, prev = s.iloc[-1], s.iloc[-2]
            diff = last - prev
            dchg = pct(last, prev)
            sec1.append(f"{i}. {name} {last:,.0f} ({diff:+,.0f} ; {fmt_pct(dchg)}) {get_weather_line(dchg)}")
            
    sec1.append("-" * 15)
    core_items = []
    seen_core = set()
    s_twii = closes.get("^TWII")
    item_idx = len(INDICES) + 1
    
    if s_twii is not None and len(s_twii) >= 2:
        last, prev = s_twii.iloc[-1], s_twii.iloc[-2]
        diff = last - prev
        dchg = pct(last, prev)
        sec1.append(f"{item_idx}. 台股大盤 {last:,.0f} ({diff:+,.0f} ; {fmt_pct(dchg)}) {get_weather_line(dchg)}")
        core_items.append({"name": "台股大盤", "yf": "^TWII", "last": last, "ma20w": weekly_ma(s_twii, 20)})
        item_idx += 1
    
    for w in watch:
        sym = w["yf_symbol"]
        if ("2330" in sym or "0050" in sym) and "正2" not in w["name"] and "反1" not in w["name"]:
            if sym in seen_core: continue
            seen_core.add(sym)
            s = closes.get(sym)
            if s is not None and len(s) >= 2:
                last, prev = s.iloc[-1], s.iloc[-2]
                diff = last - prev
                dchg = pct(last, prev)
                display_name = w["name"] if "0050" not in sym or "元大" in w["name"] else "元大台灣50"
                sec1.append(f"{item_idx}. {display_name} {last:,.1f} ({diff:+,.1f} ; {fmt_pct(dchg)}) {get_weather_line(dchg)}")
                core_items.append({"name": display_name, "yf": sym, "last": last, "ma20w": weekly_ma(s, 20)})
                item_idx += 1

    sec2 = ["▌二、VIX 恐慌指數 (<35)"]
    for i, (v_name, v_sym) in enumerate([("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")], 1):
        h = vix_hist.get(v_sym, [0, 0, 0, 0])
        val_today, val_yest, val_prev_w = h[0], h[1], h[3]
        if val_today is None or val_today <= 0:
            sec2.append(f"{i}. {v_name} — (暫無報價)")
        else:
            status = "🟢安全" if val_today < 35 else "🔴恐慌"
            yest_str = f"昨日{val_yest:.1f}" if (val_yest and val_yest > 0) else "昨日—"
            prev_w_str = f"前一週{val_prev_w:.1f}" if (val_prev_w and val_prev_w > 0) else "前一週—"
            sec2.append(f"{i}. {v_name} {val_today:.1f}({status}), {yest_str}, {prev_w_str}")

    sec3 = ["▌三、台股核心指標 (K<20)"]
    for i, item in enumerate(core_items, 1):
        kd = kd_data.get(item["yf"], {})
        kd_d = kd.get("日", (None,))[0]
        kd_w = kd.get("週", (None,))[0]
        kd_m = kd.get("月", (None,))[0]
        
        last = item["last"]
        ma20w = item["ma20w"]
        
        price_fmt = f"({last:,.0f})" if item["yf"] == "^TWII" else f"({last:,.1f})"
        sec3.append(f"{i}. {item['name']} {price_fmt}")
        sec3.append(f"   K值: 日{int(kd_d) if kd_d else '-'} / 週{int(kd_w) if kd_w else '-'} / 月{int(kd_m) if kd_m else '-'}")
        
        if last and ma20w:
            dist = (last - ma20w) / ma20w * 100
            if dist < 0: status = "⚠️ 破線空手"
            else:
                if kd_d and kd_d < 20: status = "🟢 日K超跌"
                elif kd_w and kd_w < 20: status = "🟢 週K超跌"
                elif kd_m and kd_m < 20: status = "🟢 月K超跌"
                else: status = "✓ 線上安全"
            ma_fmt = f"{ma20w:,.0f}" if item["yf"] == "^TWII" else f"{ma20w:,.1f}"
            sec3.append(f"   距20週均({ma_fmt}): {dist:+.1f}% ({status})")

    sec4 = ["▌四、其他位階警示"]
    alerts = []
    all_targets = [{"name": n, "yf": s} for s, n in INDICES] + [{"name": w["name"], "yf": w["yf_symbol"]} for w in watch]
    alert_count = 1
    for t in all_targets:
        if t["name"] not in ALERT_TARGETS: continue
        s = closes.get(t["yf"])
        if s is None or len(s) < 2: continue
        last = s.iloc[-1]
        ma = compute_ma(s)
        drop = []
        if ma.get("MA60") and last < ma.get("MA60"): drop.append("破季線")
        if ma.get("MA240") and last < ma.get("MA240"): drop.append("破年線")
        w_ret = pct(last, s.iloc[-6]) if len(s)>=6 else None
        m_ret = pct(last, s.iloc[-22]) if len(s)>=22 else None
        if (w_ret and w_ret <= -20) or (m_ret and m_ret <= -20): drop.append("週/月跌>20%")
        kd = kd_data.get(t["yf"], {})
        ko80 = [tf for tf in ["日", "週", "月"] if kd.get(tf) and kd[tf][0] is not None and kd[tf][0] > 80]
        res = ""
        if drop: res += f"🟢[跌]{','.join(drop)} "
        if ko80: res += f"🔴[熱]{'/'.join(ko80)}K>80"
        if res: 
            alerts.append(f"{alert_count}. {t['name']}: {res.strip()}")
            alert_count += 1
            
    if alerts: sec4.extend(alerts)
    else: sec4.append("*(今日無觸發)*")

    def create_box(lines, color="#333333"):
        return { "type": "text", "text": "\n".join(lines), "size": "sm", "wrap": True, "color": color, "lineSpacing": "4px" }

    report_url = os.environ.get("GITHUB_PAGES_URL", "").strip() or "https://ckliu60.github.io/stock-monitor/"

    flex_payload = {
        "type": "flex",
        "altText": f"📊 股市快報 {today.month:02d}/{today.day:02d} 已送達",
        "contents": {
            "type": "bubble",
            "size": "mega",
            "body": {
                "type": "box",
                "layout": "vertical",
                "spacing": "md",
                "contents": [
                    { "type": "text", "text": f"📊 股市核心快報 {today.month:02d}/{today.day:02d} ({wd})", "weight": "bold", "size": "md", "color": "#1f3b66" },
                    {"type": "separator", "margin": "md"}, create_box(sec1),
                    {"type": "separator", "margin": "md"}, create_box(sec2),
                    {"type": "separator", "margin": "md"}, create_box(sec3),
                    {"type": "separator", "margin": "md"}, create_box(sec4),
                    {"type": "separator", "margin": "md"},
                    {
                        "type": "button", "style": "primary", "color": "#1456b8", "height": "sm",
                        "action": { "type": "uri", "label": "👉 開啟詳細網頁報表", "uri": report_url }
                    }
                ]
            }
        }
    }
    return flex_payload

def main():
    force_dry_run = "--dry-run" in sys.argv
    watch = load_watchlist()
    
    idx_syms = [s for s, _ in INDICES]
    stk_syms = [w["yf_symbol"] for w in watch]
    all_syms = list(set(idx_syms + stk_syms + ["^TWII"]))
    closes = fetch_data(all_syms)

    v_data = fetch_data(["^VIX", "^VIXTWN"], period="1mo")
    vix_hist = {}
    for v_sym in ["^VIX", "^VIXTWN"]:
        s = v_data.get(v_sym)
        if s is not None and len(s) >= 6:
            vix_hist[v_sym] = [s.iloc[-1], s.iloc[-2], s.iloc[-3], s.iloc[-6]]
        elif s is not None and len(s) > 0:
            vix_hist[v_sym] = [s.iloc[-1], 0, 0, 0]
        else:
            vix_hist[v_sym] = [0, 0, 0, 0]

    kd_data = {}
    try:
        raw_kd = yf.download(all_syms, period="2y", interval="1d", auto_adjust=False, progress=False, group_by="ticker")
        for sym in all_syms:
            df = raw_kd if len(all_syms) == 1 else raw_kd[sym]
            if "Close" in df:
                df = df[["High", "Low", "Close"]].dropna()
                kd_data[sym] = kd_timeframes(df)
    except: pass

    # 抓取 5 則最新市場新聞
    news_data = fetch_yahoo_news(limit=5)

    # 生成網頁與 Excel 檔案
    df_report = export_all_data(closes, watch, kd_data)
    generate_html_report(df_report, vix_hist, news_data)

    # 建立卡片格式訊息
    flex_message = build_flex_message(closes, watch, kd_data, vix_hist)

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
