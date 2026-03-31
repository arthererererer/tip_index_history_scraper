# TIP 臺灣指數公司 — 歷史指數 CSV 下載與合併

從 [TIP 各項指數（一次載入全部）](https://taiwanindex.com.tw/indexes?count=-1&page=1) 取得所有指數代碼與名稱後，可下載單一指數歷史 CSV，或以 `--all` 將**全部指數**在指定日期區間內的歷史資料**合併為同一個 CSV**（每列含指數代碼、指數名稱）。

## 操作指南

 - 全部指數合併成單一csv
python scrape_tip_history.py --all --start 2026/01/01 --end 2026/03/30 -o output\all_history.csv

 - 只跑前三支指數合併成單一csv
python scrape_tip_history.py --all --limit 3 --start 2026/03/20 --end 2026/03/30

 - 只抓單一指數
python scrape_tip_history.py --code IX0232 --start 2026/03/20 --end 2026/03/30

## 環境需求

- Python 3.10+（建議）
- 依賴套件見 `requirements.txt`

### 安裝

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## 程式說明

### `scrape_tip_history.py`

結合 **requests**（列表頁、自訂 User-Agent）與 **Playwright**（歷史頁日期與下載）。

| 常數 / 設定 | 說明 |
|-------------|------|
| `BASE` | 官網網址，預設 `https://taiwanindex.com.tw` |
| `INDEX_LIST_URL` | 指數列表頁，預設 `/indexes?count=-1&page=1`（與官網「全部」筆數一致） |
| `USER_AGENT` | 模擬 Chrome 146 on Windows，用於 **requests** 與 **Playwright** |
| `EXCLUDED_INDEX_SLUGS` | 非指數詳情頁的 `/indexes/<slug>`（如 `comparison`、`board`）略過 |
| `DEFAULT_INDEX_NAME` | 單筆模式預設要比對的指數名稱 |
| `BTN_SEARCH` / `BTN_DOWNLOAD` | 「搜尋」「下載」按鈕文字（Unicode 轉義，避免編碼問題） |
| `MERGED_COLUMNS` | `--all` 合併檔的欄位順序：指數代碼、指數名稱、日期、價格指數值、報酬指數值、漲跌點數、漲跌百分比 |

| 函式 | 說明 |
|------|------|
| `session_with_ua()` | 建立帶 `User-Agent`、`Accept`、`Accept-Language` 的 `requests.Session` |
| `_register_index(seen, code, name)` | 將代碼與名稱寫入去重字典（略過排除清單） |
| `list_all_indexes()` | GET `INDEX_LIST_URL`，解析 `div.index-box`（長列表）及 `<tr>` 表格列，回傳排序後的 `[(代碼, 名稱), ...]` |
| `resolve_index_code(index_name)` | 在完整列表中以子字串比對指數名稱，回傳代碼 |
| `_download_history_on_page(page, index_code, start_date, end_date, dest)` | 在既有 Playwright `Page` 上開啟歷史頁；以 `fill(..., force=True)` 填日期（避免「查無資料」提示擋住點擊）；搜尋、下載至 `dest` |
| `scrape_history_csv(...)` | **單一指數**：啟動瀏覽器並下載一個官方格式 CSV |
| `_read_official_history_rows(csv_path)` | 讀取官網下載的 CSV（UTF-8 BOM），回傳資料列字典 |
| `scrape_all_merged_csv(...)` | **全部指數**：同一瀏覽器連續處理；**每支成功後即寫入並 flush** 合併 CSV（中斷也可保留已完成列）；終端機顯示 `[目前/總數] 代碼 名稱` 與成功/失敗摘要；**每兩支指數之間**隨機等待 1～5 秒；失敗寫入 `.errors.txt` |
| `main()` | 命令列進入點 |

### 命令列參數

| 參數 | 說明 |
|------|------|
| `--all` | 抓取列表中的**全部**指數並合併為單一 CSV |
| `--index-name` | 單筆模式：列表上的指數名稱（子字串比對） |
| `--code` | 單筆模式：直接指定代碼（如 `IX0232`） |
| `--start` / `--end` | 起訖日期，格式 `YYYY/MM/DD` |
| `-o` / `--output` | 輸出路徑。單筆預設 `output/<CODE>_history.csv`；`--all` 預設 `output/all_indexes_history.csv` |
| `--headed` | 顯示瀏覽器視窗（除錯） |
| `--limit` | 僅處理前 N 支指數（`0` 表示不限制；可用於測試） |

### 使用範例

**全部指數合併**（程式會在**每兩支指數之間**自動隨機等待 **1～5 秒**）：

```bash
python scrape_tip_history.py --all --start 2026/01/01 --end 2026/03/30 -o output/all_history.csv
```

**只測試前 3 支：**

```bash
python scrape_tip_history.py --all --limit 3 --start 2026/03/20 --end 2026/03/30
```

**單一指數（與先前相同）：**

```bash
python scrape_tip_history.py --code IX0232 --start 2026/03/20 --end 2026/03/30
```

## 輸出 CSV 欄位

- **單筆下載**：與官網相同（`日期`、`價格指數值`、`報酬指數值`、`漲跌點數`、`漲跌百分比`）。
- **`--all` 合併**：上述五欄前加上 **`指數代碼`**、**`指數名稱`**。

若部分指數下載失敗，已成功列仍會留在合併檔中，並另存 `<輸出檔名>.errors.txt` 記錄錯誤。

### 常見失敗原因（`*.errors.txt`）

- **官網提示「查無指數相關歷史資料」**：所選區間內沒有資料（例如部分聯名指數、尚未發布歷史），浮層曾擋住日期欄導致逾時；程式已改以 `fill(..., force=True)` 填日期以降低此情況。
- **沒有兩個日期輸入欄**：多為 **產業類股指數**（代碼常為 `t01`、`t02`…）等頁面與「一般指數」歷史下載介面不同，無法用同一流程自動下載。
- **逾時**：網路慢、伺服器回應久，可改 `--headed` 觀察或稍後重跑失敗代碼。

## 注意事項

- `--all` 指數數量多時執行時間長（已內建每支之間隨機 1～5 秒間隔），請遵守網站使用條款。
- 列表頁版面若改版，`list_all_indexes` 的 `index-box` / 表格解析可能需調整。
- 歷史頁 UI 變更時請以 `--headed` 除錯並更新選擇器。
