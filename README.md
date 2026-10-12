# Internal Tools — Python 小工具集

彼此獨立的內部小程式，依用途分資料夾。每支程式可單獨執行，之間無程式碼相依。

## 工具總覽

| 類別 | 工具 | 路徑 | 版本 | 用途 | 外部相依 |
|---|---|---|---|---|---|
| 訂單處理 | PHH25 (BTG) PO Extractor | `order-processing/PHH25(BTG)_PO_Extractor_V1.0.py` | V1.0 | 擷取 PHH25 訂單（PDF），輸出成 Excel 匯總表 | pdfplumber, openpyxl |
| 訂單處理 | BHP Order Extractor | `order-processing/BHPOrderExtractor_v1.5.py` | v1.5 | 擷取 BHP 客戶訂單（程式內標示 Kingfisher / Castorama PO，PDF），輸出成 Excel；可一次選多張 PDF 合併成一份，並自動核對每張 PO 的總數量與總金額 | PyMuPDF, openpyxl |
| 訂單處理 | DIB00 PO Extractor | `order-processing/dib-po-extractor/PO_PDF_Extractor_DIB00_V1.1.py` | V1.1 | 擷取 DIB00 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl |
| 訂單處理 | DIB01 PO Extractor | `order-processing/dib-po-extractor/PO_PDF_Extractor_DIB01_V1.1.py` | V1.1 | 擷取 DIB01 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl, Tesseract OCR |
| 訂單處理 | DIB02 PO Extractor | `order-processing/dib-po-extractor/PO_PDF_Extractor_DIB02_V1.2.py` | V1.2 | 擷取 DIB02 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl |
| 出貨文件 | TG Modify Commercial Invoice | `shipping-documents/TG_Modify_Commercial_Invoice_V4.py` | V4 | 修正 TG 抬頭為 AS 抬頭，供 AS 出口文件使用 | PyMuPDF |
| 溫室氣體盤查 | Export Declaration A 版 | `ghg-inventory/Extract_Declaration_A版 V1.0 (較慢，是用定位法來抓資料).py` | V1.0 | 擷取出口報單資料；定位法，速度較慢 | PyMuPDF, pdfplumber, openpyxl |
| 溫室氣體盤查 | Export Declaration B 版 | `ghg-inventory/Extract_Declaration_B版 V1.0 (較快，是用轉換成文字檔後來抓資料).py` | V1.0 | 擷取出口報單資料；先轉文字檔再抓取，速度較快 | PyMuPDF, openpyxl |
| 溫室氣體盤查 | Discharge Port（ChatGPT 版） | `ghg-inventory/取得DischargePort_ChatGPT V1.0.py` | V1.0 | 取得提單中的 Discharge Port 資料 | pdfplumber, pdf2image, Tesseract OCR |
| 溫室氣體盤查 | Discharge Port（Claude 版） | `ghg-inventory/取得DischargePort_Claude V1.0.py` | V1.0 | 取得提單中的 Discharge Port 資料 | pdfplumber, pdf2image, Tesseract OCR |
| 快遞 | FedEx Tracking Info | `express-carriers/Extract_FedEx_Tracking_Info_V2.1(修正bug及新增兼容AS reference的格式).py` | V2.1 | 擷取 FedEx Tracking／Reference 等追蹤資料，LF 及 AS 使用；V2.1 修正 bug 並兼容 AS reference 格式 | pypdf, openpyxl |
| 快遞 | UPS Invoice Extractor | `express-carriers/UPS發票內容擷取V1.0.py` | V1.0 | 擷取 UPS 發票內容，適用大量發票批次核對 | pdfplumber, openpyxl |
| 關稅監控 | CF US Tariff Monitor | `cf-us-tariff-monitor/` | — | 官方來源每日檢查，經 GitHub Actions 部署 GitHub Pages 儀表板 | 僅標準函式庫 |

## 使用方式

| 步驟 | 指令／動作 | 備註 |
|---|---|---|
| 1. 安裝套件 | `pip install -r requirements.txt` | 未指定版本 |
| 2. 需 OCR 的工具 | 依 [docs/tesseract-setup.md](docs/tesseract-setup.md) 安裝 | DIB01、DischargePort 兩支 |
| 3. 執行 | `python <路徑>` | 各工具為 tkinter 檔案選取介面（關稅監控除外）；路徑含空格或括號時需加雙引號 |
| 3a. BHP 命令列模式 | `python order-processing/BHPOrderExtractor_v1.5.py 訂單1.pdf 訂單2.pdf ... 輸出.xlsx` | 不帶參數為 GUI；帶參數時最後一個為輸出 Excel，供 VBA 等呼叫。回傳碼：0 成功、1 錯誤、2 參數不足、3 總數量／總金額檢核不符 |

## BHP Order Extractor 輸出格式

| 項目 | 內容 |
|---|---|
| 工作表 | `Order Data` |
| A–H 欄 | PO#、Banner Article No.、Article Description、Quantity、Unit Cost、Amount、POD、LRD |
| LRD | 原始 DD/MM/YYYY 轉為 Excel 真日期（`yyyy/m/d`），可排序、篩選 |
| L 欄 `Validation Result` | 每張 PO 一列，顯示總數量與總金額是否與 PDF 相符；不符時列出 PDF 與輸出的數字 |

## 目錄結構

```
internal-tools/
├── order-processing/          訂單處理
│   ├── BHPOrderExtractor_v1.5.py
│   ├── PHH25(BTG)_PO_Extractor_V1.0.py
│   └── dib-po-extractor/      DIB00 / DIB01 / DIB02
├── shipping-documents/        出貨文件
├── ghg-inventory/             溫室氣體盤查（報單、提單 Discharge Port）
├── express-carriers/          快遞（FedEx、UPS）
├── cf-us-tariff-monitor/      關稅監控（與 GitHub Actions 綁定，路徑固定）
├── docs/                      環境安裝說明
├── .github/workflows/         自動化排程與部署
└── requirements.txt
```

## 維護規則

| 規則 | 說明 |
|---|---|
| 檔名沿用原命名 | 檔名含版本號（如 `_V1.0`），升版時改檔名並同步更新本表「路徑」「版本」兩欄 |
| 不留舊版檔案 | 舊版以 Git 歷史查閱，不在資料夾中並存 |
| 新增工具 | 放入對應類別資料夾，並在「工具總覽」補一列 |
| `cf-us-tariff-monitor/` 不可任意搬移 | 兩個 workflow 以固定路徑引用此資料夾 |
| exe 不進 repo | 打包後的 exe 請放 GitHub Releases |
