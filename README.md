## 快速開始

以下步驟可在本機啟動與使用此專案：

1. 安裝相依套件

```bash
npm install
```

2. 啟動開發伺服器（Vite）

```bash
npm run dev
```

瀏覽器開啟 http://localhost:5173（或以 Vite 顯示的位址）即可看到應用程式。

3. 建置（Production）

```bash
npm run build
```

4. 預覽建置結果

```bash
npm run preview
```

使用說明：

- 在網頁中點選「選取 Excel 檔案」上傳 `.xlsx` 或 `.xls` 檔案。
- 應用程式會解析第一個工作表，嘗試擷取 `InvoiceDate`、`CustomerID`、`Country` 欄位（有欄位才會使用）。
- 解析時畫面會顯示進度條；完成後會顯示兩個圖表：
  - 客戶數量按國家（圓餅圖）
  - 每日發票數（長條圖），可切換月份檢視每日分佈

相關程式碼位置：

- `src/pages/UploadPage.tsx`：上傳、解析與圖表顯示邏輯。
- `src/workers/parseWorker.ts`：Web Worker 版本的解析程式（若有可用會優先使用）。

常見問題：

- 若瀏覽器或環境無法建立 Web Worker，程式會退回主執行緒解析（`xlsx` 套件）。
- 若資料量非常大，圖表會自動以月份彙總以維持效能。

```

```
