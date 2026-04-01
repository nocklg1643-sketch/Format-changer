# 格式轉換

「格式轉換」是一個適合私下分享的輕量網址工具。

這一版的目標不是自建高保真 Office 轉檔引擎，而是：

- 保留目前適合在瀏覽器內直接完成的本地輕量轉換
- 把介面整理成適合手機與電腦打開的網址版
- 將高保真 `Office -> PDF` 明確獨立成入口區塊
- 避免誤導使用者以為目前前端已經能完美轉 `DOC / DOCX / PPT / PPTX -> PDF`

## 私下分享建議

- 這個網站適合私下分享給朋友、家人或熟人使用
- 已在 HTML 中加入 `noindex, nofollow`，降低被搜尋引擎正常收錄的機率
- 建議不要主動提交到搜尋引擎
- 如果之後需要更高隱私，可以再加上密碼保護或部署平台的訪問限制

## 目前哪些功能現在就能用

以下功能會直接在瀏覽器端處理，不需要自建後端：

- `TXT -> MD / HTML / PDF`
- `MD -> TXT / HTML / PDF`
- `HTML -> TXT / MD / PDF`
- `JSON -> TXT / MD / HTML / PDF`
- `XML -> TXT / HTML / PDF`
- `CSV -> XLSX / JSON / HTML / PDF`
- `XLSX -> CSV / JSON / HTML / PDF`
- `JPG / JPEG -> PNG / WEBP / PDF`
- `PNG -> JPG / WEBP / PDF`
- `WEBP -> JPG / PNG / PDF`
- `PDF -> TXT`（限可擷取文字型 PDF）
- `PDF -> DOCX`（輕量文字擷取路線）
- `DOCX -> TXT / HTML / PDF`（輕量轉換）
- `PPTX -> TXT / HTML / PDF`（以可擷取文字為主）

## 哪些功能目前只是入口 / 預留

以下高保真功能目前不在瀏覽器內硬做：

- `DOC -> PDF`
- `DOCX -> PDF`
- `PPT -> PDF`
- `PPTX -> PDF`

目前做法是：

- 介面上清楚列出這些高保真轉換需求
- 先導向外部專業服務首頁
- 程式結構中保留未來改接 API / 後端的位置

補充：

- 上方主工具區仍然保留 `PDF / DOCX / PPTX` 的輕量互轉功能，適合快速整理文字內容
- 如果你在意版面、圖片、字型、投影片樣式，請改用下方高保真入口

## 為什麼這樣設計

高保真 Office 轉 PDF 若要真正保留：

- 原始版面
- 字型
- 圖片
- 投影片配置
- 頁面分頁

通常需要專業文件引擎，不適合在這一版純前端工具裡硬做低品質模擬。

## 目前仍存在的外部依賴

這一版仍使用 CDN 載入部分函式庫，所以目前不是完全離線版本：

- `pdf.js`
- `jszip`
- `mammoth`
- `xlsx`
- `pdf-lib`
- `docx`
- Google Fonts

如果之後要改成更穩定的靜態網站版本，可以把這些依賴改成本地檔案並一起部署。

## 如何部署成網址

這個專案目前維持純 `HTML / CSS / JS`，可直接部署到靜態網站平台：

### GitHub Pages

1. 建立一個 GitHub repository
2. 上傳 `index.html`、`styles.css`、`app.js`、`README.md`
3. 到 repository 的 `Settings -> Pages`
4. 選擇從主分支發布
5. 發布後會得到一個網址

### Netlify

1. 建立新站點
2. 直接拖曳整個資料夾，或連接 GitHub repository
3. Build command 留空
4. Publish directory 設為根目錄

### Vercel

1. 匯入這個專案
2. 不需要額外框架設定
3. 直接用靜態網站方式部署即可

## 未來如果要接真正高保真 API，應該從哪裡開始

建議從 [app.js](/C:/格式轉換器/app.js) 裡的高保真入口資料開始：

- `HIGH_FIDELITY_ENTRIES`

未來可把目前外部連結改成：

- 呼叫第三方轉檔 API
- 呼叫你自己的後端服務
- 先上傳檔案再回傳高保真 PDF

最簡單的升級方向是：

1. 保持現在這個前端 UI 不變
2. 只替高保真入口加上 API 串接
3. 本地輕量轉換繼續保留在瀏覽器完成

這樣可以兼顧：

- 前端簡單
- 手機可用
- 分享方便
- 未來可逐步升級
