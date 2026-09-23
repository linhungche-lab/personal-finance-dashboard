# Personal Finance Dashboard — Phase 3.2

純 HTML / CSS / JavaScript / Chart.js 的 V1 Dashboard。保留 Calibri 優先、深海軍藍、暖白與低飽和金色的原有介面及財務公式。新增本機 Excel 匯入，不使用框架、後端或雲端資料庫。

## 使用方式

1. 直接開啟 `index.html`，或將下列網站檔案放到 GitHub Pages。
2. 初始顯示「尚未載入財務資料」，不會自動使用示範數字。
3. 按「選擇 Excel」，選取你在本機填好的 V1.1 格式 `.xlsx` 或 `.xls`（不支援加密檔，最大 5 MB）。
4. 驗證成功後顯示「✓ 財務資料已載入」、檔名與載入時間。載入時間是匯入時間，不是資產估值日。
5. 更新本機 Excel 後，按「重新選擇 Excel」。重新整理或關閉分頁會清除資料，必須再次選檔。

可用靜態伺服器預覽（僅供開發，網站不需要後端）：

```sh
python3 -m http.server 8000
```

在瀏覽器開啟 `http://localhost:8000`。HTTP(S) / GitHub Pages 使用 Web Worker 解析，15 秒未完成會終止並提示；`file://` 或不支援 Worker 的瀏覽器使用相同 parser 在主執行緒本機解析，大檔可能短暫影響操作。

## 網站檔案

部署只需：`index.html`、`style.css`、`app.js`、`excel-import.js`、`excel-worker.js` 與整個 `vendor/`。

不要部署 `sample/`、`work/`、`tests/` 或任何含真實資料的 Excel。`.gitignore` 排除 Excel 與暫存檔，但不會移除已追蹤的檔案；本次沒有進行 commit、push 或部署。若採用上面的 Python 靜態伺服器，僅供自己本機測試，不應將整個專案目錄作為公開檔案伺服器。

## Excel 格式：依 V1.1 模板

本 repository 不包含 Excel 模板或任何財務資料。請使用你另外保存的 V1.1 Excel Template 填入資料；未填必要金額不是 0，匯入器會提示。沒有某項資產、負債、收入或支出時，明確填 `0`。

### 財務資料輸入

表頭在前 30 列內，必須包含 `程式欄位`、`單位`、`輸入值 / 自動值`；依表頭定位，不固定欄位字母。模板位於第 4 列，程式欄位在 C 欄、單位 D 欄、原始金額 E 欄。其餘分類、中文名稱、說明不參與計算。

| app.js 物件 | 必要程式欄位 | 單位 |
| --- | --- | --- |
| assets | realEstate, cash, foreignCash, otherInvestments, otherAssets | TWD |
| liabilities | mortgage, personalLoan, investmentLoan, otherLoans, otherLiabilities | TWD |
| income | salaryIncome, dividendIncome, rentalIncome, otherIncome | TWD / 年 |
| expenses | livingExpenses, mortgageInterest, otherInterest, fixedExpenses | TWD / 年 |

- 收入與支出是全年金額；外幣現金也必須先以 TWD 填入，沿用現有模型。
- `stocks` 列整列數值忽略，無論它是公式、快取或手填值，都不會成為網站的股票市值來源。
- 重複、缺少、未知程式欄位及錯誤單位會列出問題，避免靜默漏計。
- 必要原始欄位必須是有限的非負數值；可接受 `1,234.50` 形式的文字數字，不接受貨幣符號、百分比、日期、布林或 Excel 錯誤值。

### 持股明細

表頭在前 30 列內，模板位於第 6 列。接受表頭中的空格、換行及全形符號差異，欄位順序可以變動。

| Excel 表頭 | app.js 欄位 | 規則 |
| --- | --- | --- |
| 市場 | market | 必填文字 |
| 股票代號 | ticker | 必填；建議使用文字儲存，例如 0050 |
| 名稱 | name | 可空白 |
| 類型 | type | 可空白，不影響 V1 財務公式 |
| 幣別 | currency | 三碼代碼，例如 TWD、USD |
| 持有股數 | shares | 大於 0，支援零股及小數股數 |
| 平均成本/股 | avgCost | 大於或等於 0，必填，以免現有成本公式把缺漏當 0 |
| 目前價格/股 | currentPrice | 大於 0，使用者維護的原幣價格 |
| 匯率 (TWD/1原幣) | fxToTwd | 大於 0；TWD 必須填 1 |

完全空白的原始輸入列略過，即使計算欄預填公式也不算持股。模板的合併說明列略過。部分填寫的持股列會拒絕整份匯入，不會悄悄排除後產生較低的總市值。允許同代號多筆紀錄，沿用原有逐筆計算方式，不自行合併。

網站不讀取持股的市值、成本總額、損益、報酬率或權重計算欄。持股可延伸到原模板第 106 列之外；兩張資料表各限制在 10,000 列、64 欄內，超限會報錯而非截斷。

## 本機資料流程與原有公式

```text
使用者選取 Excel
→ File.arrayBuffer()
→ SheetJS（只解析財務資料輸入、持股明細）
→ ExcelImport 驗證並映射原始輸入
→ { assets, liabilities, income, expenses, holdings }
→ 原有 buildFinancialModel / normalizeHoldings
→ Dashboard
```

`V1財務模型規格` 不納入網站解析或計算。原始輸入欄位若包含公式（包括有 cached value 的公式），會拒絕並提供 Sheet、儲存格位置及欄位名稱；不執行 Excel 公式。財務指標由原有 JavaScript 重新計算。

`stocks = Σ(shares × currentPrice × fxToTwd)`。原有資產、負債、淨資產、比例、緊急預備金、成本與損益公式均保留。

匯入採整份驗證：錯誤會集中列在上方，保留上一份成功資料並明示「目前仍顯示上一份成功載入的資料」；第一份匯入失敗則保持未載入狀態。檔名與持股文字均安全呈現，不能被解讀成 HTML。

## 隱私與依賴

- Excel 不上傳、不回寫，不使用 localStorage、sessionStorage、IndexedDB、cookie 或服務工作者保存財務資料。
- 僅保存目前分頁記憶體中的原始映射資料與計算結果。不自動開啟或抓取 `sample/`。
- SheetJS CE 0.20.3：官方 standalone build，Apache-2.0，固定在 `vendor/xlsx-0.20.3.full.min.js`。
- Chart.js 4.4.7：保留原版本，MIT，改放 `vendor/chart-4.4.7.umd.min.js`。
- 相對路徑適合 GitHub Pages 專案子路徑及直接開啟 HTML。兩套函式庫與授權文字均在 `vendor/`，不需要執行時 CDN。
- 頁面 CSP 禁止連線請求與表單提交（`connect-src 'none'`、`form-action 'none'`）。頁面首次載入網站程式檔案與 Excel 的本機解析是兩回事。

官方參考：
- https://docs.sheetjs.com/docs/getting-started/installation/standalone/
- https://docs.sheetjs.com/docs/api/parse-options/
- https://docs.sheetjs.com/docs/csf/features/formulae/

## 測試

無需安裝網站執行依賴。Node.js 可執行 parser / 原有模型測試：

```sh
node --test tests/excel-import.test.cjs
```

有 Playwright 與 Chromium 測試環境時，可執行：

```sh
node tests/browser.test.cjs
```

DOM 狀態與 CSS 語法測試使用開發工具 `linkedom` / `postcss`（不部署到網站）。本次暫存於 `work/node_modules`，可重跑：

```sh
NODE_PATH=./work/node_modules node --test tests/excel-import.test.cjs tests/dom.test.cjs
```

本次驗證：上述 22 項測試全部通過，含在 VM 模擬 Worker 環境執行實際 worker/parser 程式；`buildFinancialModel` 與 `normalizeHoldings` 與修改前逐字一致，來源範例 Excel 雜湊未變。原生 Chromium 因 macOS 沙盒的 MachPort 權限限制無法啟動，因此 `browser.test.cjs` 尚未通過實際執行，桌機／手機視覺、原生 Worker 與 CSP 的瀏覽器實測仍待在一般瀏覽器環境確認。DOM 模擬不等於視覺驗證。

測試資料全部是獨立虛構資料，不使用真實財務數字。瀏覽器測試包含背景解析、直接 file:// 開啟、錯誤後復原、保留舊資料、無持股重載、重新整理清空、HTML 跳脫、手機版面及沒有資料儲存／對外資料請求。Playwright 僅為開發測試工具，不會加入網站執行依賴。
