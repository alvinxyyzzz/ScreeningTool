# Screening Atlas — Factory Screening Web App

**繁體中文｜個人版 1.0｜2026-11-07**

這是一個可部署到 GitHub Pages、可加到 iPhone 主畫面的靜態 PWA。設計原則是 **Code public, data private**：GitHub 放程式碼，你公司的 Excel 或 JSON **永遠不要放進 GitHub 專案**。

## 1. 部署到 GitHub Pages

1. 下載 `Screening_Atlas_Web_App_v1.0.zip` 並解壓縮。
2. 在 GitHub 建立一個 Repository（例如 `screening-atlas`）。用 `Add file` → `Upload files`，上傳**解壓縮後的全部內容**，包括 `index.html`、`styles.css`、`js/` 和 `assets/`。**不要上傳 ZIP 本身，也不要上傳任何 Excel 或私有 JSON。**
3. 在 GitHub Repository → **Settings → Pages**，選擇 `Deploy from a branch` → `main` → `/ (root)`，儲存。
4. 你的網頁一般會是 `https://你的帳號.github.io/screening-atlas/`，實際網址以 GitHub Pages 畫面顯示為準。
5. 用 iPhone 的 **Safari** 開啟網址，點「分享」→「加入主畫面」，就會顯示 Screening Atlas icon。第一次開啟 App 時，點「資料更新」匯入 `Screening Tool_v1.9.xlsm` 或提供的私人 JSON。

⚠️ 有些方案的 GitHub Pages / 私人儲存庫規則可能不同。**私人 Repository 不代表網站內容只有你看得到**；此 App 原則上不在網站上部署任何公司資料。

## 2. 每月更新：不用再請 AI 轉 JSON

1. 繼續在原始 Excel 更新 CLP、CTA、Aii、CTM、Strategic List 等表。完成後，在 Excel **完整重新計算公式並儲存**。
2. 將最新版 `.xlsm` 或 `.xlsx` 存到你的 iPhone「檔案」（例如 iCloud Drive）或 Mac 本機。
3. 打開 Screening Atlas → **資料更新** → **選擇檔案** → 選新版 Excel。
4. App 直接在當前裝置解析 Excel，不執行巨集、不上傳到任何 API；顯示 Master、Engagement、Strategic、Assessment 筆數與舊版本差異。
5. 確認無誤後按 **確認更新本機資料**，資料庫便在這台裝置更新。保留上一版可回復。
6. 建議每月下載一次 **私人 JSON 備份**，放到你有權使用的加密檔案空間。**這份 JSON 含公司資料，不要 commit 到公開 Repo。**

**網站程式碼不必每月重新部署。** 只有介面功能修改、欄位結構有大變動或匯入規則需要更新時，才需重新上傳程式碼。

## 3. 現有功能

- 總覽：Master 工廠、合作歷史、戰略來源與資料健康度指標。
- 單筆工廠篩查：Higg ID、工廠名稱、國家、品牌。
- 批次工廠篩查：貼上最多 200 行資料，匯出 CSV。
- 合作資料庫：工廠、國家、品牌、CTA/CTS/BPAP/CTM/Implementation 篩選。
- 跨來源查詢：**包括不在 Master List 裡的 Strategic / Assessment / Manufacturer / Implementation 工廠**，但不假裝那些工廠已經合作。
- 工廠詳情：Master 別名、Higg ID、品牌、各筆合作來源、Assessment、Strategic、Implementation、Manufacturer。
- 資料健康度：同一 Master 多 Higg ID、跨國合併、同 ID 跨多個 Master 等提示。
- 本機版本管理：匯入前比較筆數、保留前一版、匯出／回復 JSON。
- 手機介面及 iOS PWA 圖示；裝好後 App shell 支援離線開啟（私人資料必須已存於該裝置）。

## 4. 資料結構與對應工作表

| 用途 | 原始 Excel 工作表 | 注意事項 |
| --- | --- | --- |
| 工廠主檔 | `Master list` | 1666 Master 紀錄；多行文字可能代表合併別名與多 Higg ID |
| 合作歷史明細 | `Consolidated list` | 2913 來源紀錄；`Master code` 是追蹤連結鍵 |
| 戰略名單 | `Input - Strategic list` | 5238 原始名單列；不等同已合作工廠 |
| Assessment | `Input - Assessment Records` | 1482 條；同工廠可出現多次工程師／時間紀錄 |
| 母集團 | `Input - Manufacturer info` | 521 條；母集團不能被當成單一廠址 |
| Implementation | `Input - Implementation records` | 35 條；各階段與責任人按原始紀錄顯示 |
| 更新履歷 | `Update Log`、來源 sheet 第一列 | 顯示各來源不同更新時間，不只看檔案名稱 |

**安全配對順序**：精確 Higg ID → 同國家、正規化後名稱完全一致 → 名稱相似僅列候選。身份有衝突或多對多時不強行綁定；保留為「Source only」待人工確認。

⚠️ Excel 版使用 `XLOOKUP`、`FILTER`、`LET`、`TEXTJOIN`、`SEARCH` 等運算與 cached formula results；網頁版不是 Excel 的公式運算引擎，也不會重新執行 VBA。匯入新版前務必在 Excel 計算、儲存。若日後新增工作表、變更欄位排列／合併邏輯，請同步更新 `js/excel.js`。

## 5. 安全與跨裝置限制

- 使用者匯入的資料以 **IndexedDB（瀏覽器本機儲存）** 儲存在該裝置；程式碼不內嵌公司原始資料。
- 程式沒有第三方 JS/CDN、分析追蹤、API 上傳；Content Security Policy 限制資源與對外連線。
- **IndexedDB 並非端對端加密的機密資料保險箱**。遺失裝置、瀏覽器清除資料、iOS 儲存空間回收都有風險。請設裝置密碼，並另存私人備份。
- iPhone 與 Mac **不會自動同步**，要各自在同一 App 網址匯入 Excel 或私人備份。
- PWA shell 離線使用不等於安全分享／多人權限管理。未來若要同事共享與即時同步，應以有認證與權限控管的私有後端取代本機資料快照。
- 使用與部署公司的資料前，先確認你有權將其放到個人裝置及使用未經公司批准的雲端／網站。最保守的做法是維持公司授權裝置與批准的部署環境。

## 6. 檔案說明

```
index.html             UI 與 icon 系統
styles.css             桌面／手機響應式樣式
manifest.webmanifest   iOS／PWA metadata
sw.js                  僅快取 App shell（不快取私人匯入資料）
assets/icon.svg        原創 icon 向量主檔
assets/icon-180.png    iPhone 主畫面圖示
assets/icon-192.png    PWA icon
assets/icon-512.png    PWA icon
assets/icon-1024.png   高清 icon
js/excel.js            前端解析 Excel xlsm / xlsx ZIP XML，免依賴
js/model.js            配對、來源關聯、品質檢查、CSV
js/app.js              介面、資料更新、IndexedDB
```

## 7. 設計參考

- **Linear**：一致性導航、低干擾結構、清楚資訊階層。
- **Airtable Interfaces**：以 source record 為中心的列表、篩選、Dashboard 與深入細節。
- **Attio** 類型的資料工作台概念：快速找到同一實體不同來源的紀錄。

這不是照抄現成 UI，而是為碳減排顧問的實務工廠篩查特別設計。
