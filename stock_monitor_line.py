#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股／台股監控推播系統 (V4.1 雲端全自動版 - 手機排版優化)
---------------------------------------------------------------------------
更新日誌 (V4.1): 
- 移除全形空白對齊，解決 LINE 手機版文字擠壓與不正常折行問題。
- 「台股核心指標」改為上下分行結構，提升手機閱讀舒適度。
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

PERIODS = [("週", 5), ("月", 21), ("年", 252)]
MA_WINDOWS = [("MA10", 10), ("MA20", 20), ("MA60", 60), ("MA240", 240)]

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"
LINE_MAX_LEN = 4800

# 雲端 CSV 網址 (V4.0 正確版)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRa7RG_nKx8DTBSsOsBPz4MQZPw_mYByDYy7hux2MwuhqRn1VIvqr5XWiZttrW24A/pub?gid=1903088963&single=true&output=csv"

# 檔案路徑 (供 GitHub Actions 雲端網頁使用)
HERE = Path(__file__).resolve().parent
PUBLIC_DIR = HERE / "public"
REPORT_HTML_PATH = PUBLIC_DIR / "index.html"
REPORT_CSV_PATH = PUBLIC_DIR / "report_data.csv"

# ---------------------------------------------------------------------------
# 工具函式
# ---------------------------------------------------------------------------

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
            yf_symbol = f"{code}.TW" if (market=="台股" or code.isdigit()) else code
        rows.append({
            "market": market, "code": code, "yf_symbol": yf_symbol,
            "name": (raw.get("name") or yf_symbol).strip(),
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

def get_weather(pct_val):
    if pct_val is None: return ""
    if pct_val >= 2.0: return "☀️"
    elif pct_val > 0: return "🌤️"
    elif pct_val == 0: return "☁️"
    elif pct_val > -2.0: return "🌧️"
    else: return "⛈️"

def fmt_pct(v):
    if v is None: return "—"
    return f"{'▲' if v>0 else '▼' if v<0 else ''}{abs(v):.1f}%"

# ---------------------------------------------------------------------------
# 匯出所有數據 (報表生成保持不變)
# ---------------------------------------------------------------------------
def export_all_data(closes, watch, kd_data):
    all_targets = [{"market": "美股", "code": s, "name": n, "yf": s, "group": "核心"} for s, n in INDICES] + \
                  [{"market": "台股", "code": "^TWII", "name": "台股大盤", "yf": "^TWII", "group": "核心"}] + \
                  [{"market": w["market"], "code": w["code"], "name": w["name"], "yf": w["yf_symbol"], "group": w["group"]} for w in watch]
    
    rows = []
    for i, t in enumerate(all_targets, 1):
        sym = t["yf"]
        s = closes.get(sym)
        if s is None or len(s) < 252: continue
        
        last = s.iloc[-1]
        dt_last = s.index[-1].strftime("%Y-%m-%d")
        
        pct_1d = pct(last, s.iloc[-2]) if len(s) >= 2 else None
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
            "序號": i,
            "群組": t["group"],
            "市場": t["market"],
            "代碼": t["code"],
            "名稱": t["name"],
            "yf_symbol": sym,
            "最新收盤": round(last, 2),
            "資料日期": dt_last,
            "當日%": round(pct_1d, 2) if pct_1d is not None else None,
            "週%": round(pct_1w, 2) if pct_1w is not None else None,
            "月%": round(pct_1m, 2) if pct_1m is not None else None,
            "年%": round(pct_1y, 2) if pct_1y is not None else None,
            "日K": kd_d[0], "日D": kd_d[1],
            "週K": kd_w[0], "週D": kd_w[1],
            "月K": kd_m[0], "月D": kd_m[1],
            "MA10": ma.get("MA10"),
            "MA20": ma.get("MA20"),
            "MA60": ma.get("MA60"),
            "MA240": ma.get("MA240"),
            "MA20週": ma20w
        }
        rows.append(row)
        
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    df.to_csv(REPORT_CSV_PATH, index=False, encoding="utf-8-sig")
    return df

def generate_html_report(df, vix_hist):
    now_str = dt.datetime.now().strftime("%Y/%m/%d %H:%M")
    
    idx_df = df[df['名稱'].isin([n for _, n in INDICES])]
    tw_idx = df[df['名稱'] == '台股大盤']
    us_stocks = df[(df['市場'] == '美股') & (~df['名稱'].isin([n for _, n in INDICES]))]
    tw_stocks = df[(df['市場'] == '台股') & (df['名稱'] != '台股大盤')]
    
    def render_pct_td(val):
        if pd.isna(val): return '<td class="num mute">—</td>'
        cls = 'up' if val > 0 else 'down' if val < 0 else 'mute'
        mark = '▲' if val > 0 else '▼' if val < 0 else ''
        return f'<td class="num {cls}">{mark}{abs(val):.1f}%</td>'
        
    def render_row(row, idx):
        url = f"https://finance.yahoo.com/quote/{row['yf_symbol']}"
        return f"<tr><td class='num idx'>{idx}</td><td class='name'><a href='{url}' target='_blank'>{row['名稱']}</a> <span class='code'>{row['代碼']}</span></td><td class='num'>{row['最新收盤']:,.2f}</td>{render_pct_td(row['當日%'])}{render_pct_td(row['週%'])}{render_pct_td(row['月%'])}{render_pct_td(row['年%'])}</tr>"

    html = f"<!doctype html><html lang='zh-Hant'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>股市監控報表</title><style>:root {{ --up:#d32f2f; --down:#1b8e3a; --line:#e3e6ea; --mute:#8a9099; --head:#1f3b66; }} * {{ box-sizing:border-box; }} body {{ margin:0; background:#f6f7f9; color:#222; font-family:'Microsoft JhengHei',system-ui,sans-serif; font-size:13px; }} .wrap {{ max-width:760px; margin:0 auto; padding:16px; }} h1 {{ font-size:18px; margin:4px 0; }} h2 {{ font-size:14px; margin:18px 0 6px; color:var(--head); border-left:4px solid var(--head); padding-left:8px; }} .sub {{ color:var(--mute); font-size:12px; margin-bottom:6px; }} table {{ width:100%; border-collapse:collapse; background:#fff; border:1px solid var(--line); border-radius:8px; overflow:hidden; }} th,td {{ padding:6px 8px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }} thead th {{ background:#eef1f5; color:#333; font-weight:600; font-size:12px; }} thead th.num {{ text-align:center; }} thead th small {{ color:var(--mute); font-weight:400; font-size:10px; display:block; }} tbody tr:last-child td {{ border-bottom:none; }} td.num {{ text-align:center; font-variant-numeric:tabular-nums; }} td.idx {{ color:var(--mute); width:30px; text-align:center; }} .code {{ color:var(--mute); font-size:11px; }} .up {{ color:var(--up); }} .down {{ color:var(--down); }} .mute {{ color:var(--mute); }} tr.grp td {{ background:#f0f3f7; font-weight:700; color:var(--head); }} .foot {{ color:var(--mute); font-size:11px; margin-top:14px; line-height:1.6; }} a {{ color:#1456b8; text-decoration:none; }} a:hover {{ text-decoration:underline; }}</style></head><body><div class='wrap'><h1>📊 股市核心監控報表</h1><div class='sub'>產生時間：{now_str}</div><h2>大盤與 VIX 狀況</h2><div style='overflow-x:auto;'><table><thead><tr><th>指數</th><th class='num'>收盤</th><th class='num'>當日漲跌</th><th class='num'>燈號</th></tr></thead><tbody>"
    
    for _, row in idx_df.iterrows():
        dchg = row['當日%']
        html += f"<tr><td><a href='https://finance.yahoo.com/quote/{row['yf_symbol']}' target='_blank'>{row['名稱']}</a></td><td class='num'>{row['最新收盤']:,.0f}</td>{render_pct_td(dchg)}<td class='num'>{get_weather(dchg)}</td></tr>"
        
    for _, row in tw_idx.iterrows():
        dchg = row['當日%']
        html += f"<tr><td><a href='https://finance.yahoo.com/quote/{row['yf_symbol']}' target='_blank'>{row['名稱']}</a></td><td class='num'>{row['最新收盤']:,.0f}</td>{render_pct_td(dchg)}<td class='num'>{get_weather(dchg)}</td></tr>"
        
    for v_name, v_sym in [("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")]:
        h = vix_hist.get(v_sym, [0])
        status = "🟢安全" if h[0] < 35 else "🔴恐慌"
        html += f"<tr><td>{v_name}</td><td class='num'>{h[0]:.2f}</td><td class='num mute'>—</td><td class='num'>{status}</td></tr>"
        
    html += "</tbody></table></div><h2>個股追蹤清單</h2><div style='overflow-x:auto;'><table><thead><tr><th class='num'>#</th><th>名稱</th><th class='num'>收盤</th><th class='num'>當日</th><th class='num'>週<small>(約5日)</small></th><th class='num'>月<small>(約21日)</small></th><th class='num'>年<small>(約252日)</small></th></tr></thead><tbody><tr class='grp'><td colspan='7'>美股觀察區</td></tr>"
      
    idx = 1
    for _, row in us_stocks.iterrows():
        html += render_row(row, idx)
        idx += 1
        
    html += "<tr class='grp'><td colspan='7'>台股觀察區</td></tr>"
    for _, row in tw_stocks.iterrows():
        html += render_row(row, idx)
        idx += 1
        
    html += "</tbody></table></div><h2>超跌警示</h2><ul style='list-style:none; padding:0; margin:0;'>"
  
    over_sold = False
    for _, row in df.iterrows():
        if row['名稱'] in [n for _, n in INDICES] or row['名稱'] == '台股大盤': continue
        ma20w = row['MA20週']
        last = row['最新收盤']
        pct_1m = row['月%']
        alerts = []
        if pd.notna(ma20w) and last < ma20w:
            alerts.append(f"跌破 20 週均線 ({(last-ma20w)/ma20w*100:.1f}%)")
        if pd.notna(pct_1m) and pct_1m <= -20:
            alerts.append(f"月跌幅達 {pct_1m:.1f}%")
            
        if alerts:
            over_sold = True
            html += f"<li style='background:#fff; border:1px solid var(--line); border-radius:6px; padding:6px 10px; margin-bottom:4px;'>⚡ <b>{row['名稱']}</b> <span class='code'>{row['代碼']}</span>：{', '.join(alerts)}</li>"
            
    if not over_sold:
        html += "<li style='background:#fff; border:1px solid var(--line); border-radius:6px; padding:6px 10px; margin-bottom:4px;'>✅ 今日無標的觸發超跌警示。</li>"

    html += "</ul><div class='foot'>資料來源：<a href='https://finance.yahoo.com/' target='_blank'>Yahoo Finance</a><br>報表產生器：V4.1 全自動雲端版<br><a href='report_data.csv' target='_blank'>📥 下載完整 CSV 原始數據</a></div></div></body></html>"
    
    with open(REPORT_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)

# ---------------------------------------------------------------------------
# 組裝 LINE 訊息 (V4.1 手機版優化：捨棄全形空白對齊，改為多行縮排結構)
# ---------------------------------------------------------------------------
def build_message(closes, watch, kd_data, vix_hist):
    today = dt.datetime.now()
    wd = "一二三四五六日"[today.weekday()]
    lines = [f"📊 股市核心快報 {today.month:02d}/{today.day:02d} ({wd})\n"]

    lines.append("▌一、大盤與核心狀況")
    for i, (sym, name) in enumerate(INDICES, 1):
        s = closes.get(sym)
        if s is not None and len(s) >= 2:
            last, prev = s.iloc[-1], s.iloc[-2]
            dchg = pct(last, prev)
            lines.append(f"{i}. {name} {last:,.0f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
            
    lines.append("-" * 20)
    
    core_items = []
    seen_core = set()
    s_twii = closes.get("^TWII")
    item_idx = len(INDICES) + 1
    
    if s_twii is not None and len(s_twii) >= 2:
        last, prev = s_twii.iloc[-1], s_twii.iloc[-2]
        dchg = pct(last, prev)
        lines.append(f"{item_idx}. 台股大盤 {last:,.0f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
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
                dchg = pct(last, prev)
                display_name = w["name"] if "0050" not in sym or "元大" in w["name"] else "元大台灣50"
                lines.append(f"{item_idx}. {display_name} {last:,.1f} ({fmt_pct(dchg)}) {get_weather(dchg)}")
                core_items.append({"name": display_name, "yf": sym, "last": last, "ma20w": weekly_ma(s, 20)})
                item_idx += 1

    lines.extend(["\n▌二、VIX 恐慌指數 (<35)"])
    for i, (v_name, v_sym) in enumerate([("美 VIX", "^VIX"), ("台 VIX", "^VIXTWN")], 1):
        h = vix_hist.get(v_sym, [0,0,0,0])
        status = "🟢安全" if h[0] < 35 else "🔴恐慌"
        lines.append(f"{i}. {v_name} {h[0]:.1f} ({status})")

    lines.extend(["\n▌三、台股核心指標 (K<20)"])
    for i, item in enumerate(core_items, 1):
        kd = kd_data.get(item["yf"], {})
        kd_d = kd.get("日", (None,))[0]
        kd_w = kd.get("週", (None,))[0]
        kd_m = kd.get("月", (None,))[0]
        
        # 【V4.1 修改重點】上下分行結構，手機閱讀更舒適
        lines.append(f"{i}. {item['name']}")
        lines.append(f"   K值: 日{int(kd_d) if kd_d else '-'} / 週{int(kd_w) if kd_w else '-'} / 月{int(kd_m) if kd_m else '-'}")
        
        last = item["last"]
        ma20w = item["ma20w"]
        if last and ma20w:
            dist = (last - ma20w) / ma20w * 100
            if dist < 0: status = "⚠️ 破線空手"
            else:
                if kd_d and kd_d < 20: status = "🟢 日K超跌"
                elif kd_w and kd_w < 20: status = "🟢 週K超跌"
                elif kd_m and kd_m < 20: status = "🟢 月K超跌"
                else: status = "✓ 線上安全"
            lines.append(f"   距20週均: {dist:+.1f}% ({status})")

    lines.extend(["\n▌四、其他位階警示"])
    alerts = []
    
    all_targets = [{"name": n, "yf": s} for s, n in INDICES] + \
                  [{"name": w["name"], "yf": w["yf_symbol"]} for w in watch]
    
    alert_count = 1
    for t in all_targets:
        if t["name"] not in ALERT_TARGETS: continue
            
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
        if drop: res += f"🟢[跌]{','.join(drop)} "
        if ko80: res += f"🔴[熱]{'/'.join(ko80)}K>80"
        
        if res: 
            alerts.append(f"{alert_count}. {t['name']}: {res.strip()}")
            alert_count += 1

    if alerts: lines.extend(alerts)
    else: lines.append("*(今日無觸發)*")

    report_url = os.environ.get("GITHUB_PAGES_URL", "https://ckliu60.github.io/stock-monitor/").strip()
    lines.extend([
        f"\n👉 五、詳細個股網頁報表：\n{report_url}"
    ])

    text = "\n".join(lines).rstrip()
    return text[:LINE_MAX_LEN] + "\n…（訊息截斷）" if len(text) > LINE_MAX_LEN else text


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
        else:
            vix_hist[v_sym] = [0,0,0,0]

    kd_data = {}
    try:
        raw_kd = yf.download(all_syms, period="2y", interval="1d", auto_adjust=False, progress=False, group_by="ticker")
        for sym in all_syms:
            df = raw_kd if len(all_syms) == 1 else raw_kd[sym]
            if "Close" in df:
                df = df[["High", "Low", "Close"]].dropna()
                kd_data[sym] = kd_timeframes(df)
    except: pass

    # 1. 匯出完整數據 CSV 與 HTML 報表
    df_report = export_all_data(closes, watch, kd_data)
    generate_html_report(df_report, vix_hist)
    print(f"[OK] 成功產生 report.html 與 report_data.csv")

    # 2. 推播 LINE 訊息
    text = build_message(closes, watch, kd_data, vix_hist)

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