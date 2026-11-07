# 美股／台股監控推播系統

每週二、四、六 台灣時間 06:30，自動抓美股與台股真實收盤價，計算跨期漲跌、標示階梯式警示，整理成一則精簡訊息推到你的 LINE。並可在 LINE 直接打字增刪股票。

- 資料源：Yahoo Finance（`yfinance`），只用真實收盤價。
- 推播：LINE Messaging API。
- 排程：GitHub Actions（電腦免開機）；另附本機排程選項。
- 清單：Google 試算表（可在 LINE 增刪）；未設定時用本機 `watchlist.csv`。

---

## 檔案說明

| 檔案 | 用途 |
| --- | --- |
| `stock_monitor_line.py` | 主程式：抓資料、算漲跌、組訊息、推 LINE |
| `watchlist.csv` | 初始股票清單（本機後援／試算表來源） |
| `requirements.txt` | Python 相依套件 |
| `.env.example` | 金鑰範本（複製成 `.env` 填入，勿上傳） |
| `.github/workflows/stock_monitor.yml` | GitHub Actions 雲端排程 |
| `run_job.bat` | 本機 Windows 排程用批次檔 |
| `apps_script/` | 方案 A：LINE 增刪股票（Google 試算表 + Apps Script） |

---

## 步驟一：本機跑跑看（不需金鑰）

```bash
pip install -r requirements.txt
python stock_monitor_line.py --dry-run
```

會直接把訊息印在畫面上（不推播）。確認數字與 Yahoo Finance 一致即可。

> 數據防呆：若某檔的歷史價格出現「被漏記的股票分割」（例如某台股 ETF），程式會自動偵測並還原，避免跨期報酬出現假訊號。

---

## 步驟二：取得 LINE 金鑰

到 [LINE Developers](https://developers.line.biz/)（用你既有的 channel）：

1. **Messaging API** 分頁 → 產生 / 複製 **Channel access token（long-lived）** → 這是 `LINE_CHANNEL_ACCESS_TOKEN`。
2. **Basic settings** 分頁 → 最下方 **Your user ID**（`U` 開頭約 33 碼）→ 這是 `LINE_USER_ID`。
   - ⚠️ 注意：這**不是**加好友用的 LINE ID（例如 `ckliu60`），那個不能用於推播。
3. 用手機 LINE **把這個官方帳號加為好友**（否則推播會失敗）。

---

## 步驟三：本機實推一次

把 `.env.example` 複製成 `.env`，填入 token 與 User ID，然後：

```bash
pip install python-dotenv
python -c "from dotenv import load_dotenv; load_dotenv(); import runpy; runpy.run_path('stock_monitor_line.py', run_name='__main__')"
```

或直接在這次終端機設環境變數後執行（PowerShell）：

```powershell
$env:LINE_CHANNEL_ACCESS_TOKEN="你的token"; $env:LINE_USER_ID="你的UserID"; python stock_monitor_line.py
```

手機收到訊息即成功。

---

## 步驟四：雲端自動排程（GitHub Actions）

1. 在 GitHub 新建一個**私有** repo，把本資料夾所有檔案上傳（`.env` 不會上傳，已被 `.gitignore` 忽略）。
2. repo → **Settings → Secrets and variables → Actions → New repository secret**，新增：
   - `LINE_CHANNEL_ACCESS_TOKEN`
   - `LINE_USER_ID`
   - （可選）`SHEET_CSV_URL`
3. repo → **Actions** 分頁 → 選「美股台股監控推播」→ **Run workflow** 手動觸發一次，確認手機收到。
4. 之後它會依排程自動執行：cron `30 22 * * 1,3,5`（UTC）＝台灣 **週二、四、六 06:30**。

> GitHub Actions 的排程在尖峰可能延遲數分鐘，屬正常。

---

## 步驟五：在 LINE 增刪股票（方案 A，選做）

詳見 [`apps_script/README.md`](apps_script/README.md)。完成後可在 LINE 直接打：

- `清單`
- `加 台股 2330`、`加 NVDA`
- `刪 2330`、`刪 NVDA`

---

## （可選）本機 Windows 排程

若不想用雲端，可用 Windows 工作排程器：

1. 先在本機設好 `LINE_CHANNEL_ACCESS_TOKEN`、`LINE_USER_ID`（系統環境變數）。
2. 新增工作 → 觸發程序：每週，勾選 **二、四、六**，時間 **06:30**。
3. 動作：啟動程式，選 `run_job.bat`。
4. 條件：可勾「只有在電腦處於連線狀態時才啟動」。

> 本機排程需電腦在 06:30 為開機／連線狀態；雲端（步驟四）則不需要。

---

## 指令速查（LINE）

| 指令 | 作用 |
| --- | --- |
| `清單` | 列出目前追蹤清單 |
| `加 台股 2330` 或 `加 2330` | 新增（純數字自動判為台股、加 .TW） |
| `加 美股 NVDA` 或 `加 NVDA` | 新增美股 |
| `刪 2330` / `刪 NVDA` | 移除 |

---

## 警示符號

依報酬率「絕對幅度」標示（漲跌皆適用，下跌額外標「（逢低）」）：

| 符號 | 幅度 |
| --- | --- |
| ⚠️ | > 10% |
| 🚨 | > 20% |
| ⚡ | > 30% |
| 💥 | > 50% |

漲用 ▲、跌用 ▼。
