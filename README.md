# Internal Tools — Python 小工具集

彼此獨立的內部小程式，依用途分資料夾。每支程式可單獨執行，之間無程式碼相依。

## 工具總覽

| 類別 | 工具 | 路徑 | 版本 | 用途 | 外部相依 |
|---|---|---|---|---|---|
| 訂單處理 | PHH25 (BTG) PO Extractor | `order-processing/phh25_btg_po_extractor.py` | V1.0 | 擷取 PHH25 訂單（PDF），輸出成 Excel 匯總表 | pdfplumber, openpyxl |
| 訂單處理 | DIB00 PO Extractor | `order-processing/dib-po-extractor/dib00_po_extractor.py` | V1.1 | 擷取 DIB00 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl |
| 訂單處理 | DIB01 PO Extractor | `order-processing/dib-po-extractor/dib01_po_extractor.py` | V1.1 | 擷取 DIB01 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl, Tesseract OCR |
| 訂單處理 | DIB02 PO Extractor | `order-processing/dib-po-extractor/dib02_po_extractor.py` | V1.2 | 擷取 DIB02 訂單（PDF），輸出成 Excel 匯總表 | PyMuPDF, openpyxl |
| 出貨文件 | TG Modify Commercial Invoice | `shipping-documents/tg_modify_commercial_invoice.py` | V4 | 修正 TG 抬頭為 AS 抬頭，供 AS 出口文件使用 | PyMuPDF |
| 溫室氣體盤查 | Export Declaration A 版 | `ghg-inventory/export_declaration_a_locator.py` | V1.0 | 擷取出口報單資料；定位法，速度較慢 | PyMuPDF, pdfplumber, openpyxl |
| 溫室氣體盤查 | Export Declaration B 版 | `ghg-inventory/export_declaration_b_text.py` | V1.0 | 擷取出口報單資料；先轉文字檔再抓取，速度較快 | PyMuPDF, openpyxl |
| 溫室氣體盤查 | Discharge Port（ChatGPT 版） | `ghg-inventory/discharge_port_chatgpt.py` | V1.0 | 取得提單中的 Discharge Port 資料 | pdfplumber, pdf2image, Tesseract OCR |
| 溫室氣體盤查 | Discharge Port（Claude 版） | `ghg-inventory/discharge_port_claude.py` | V1.0 | 取得提單中的 Discharge Port 資料 | pdfplumber, pdf2image, Tesseract OCR |
| 快遞 | FedEx Tracking Info | `express-carriers/fedex_tracking_info.py` | V2.1 | 擷取 FedEx Tracking／Reference 等追蹤資料，LF 及 AS 使用；V2.1 修正 bug 並兼容 AS reference 格式 | pypdf, openpyxl |
| 快遞 | UPS Invoice Extractor | `express-carriers/ups_invoice_extractor.py` | V1.0 | 擷取 UPS 發票內容，適用大量發票批次核對 | pdfplumber, openpyxl |
| 關稅監控 | CF US Tariff Monitor | `cf-us-tariff-monitor/` | — | 官方來源每日檢查，經 GitHub Actions 部署 GitHub Pages 儀表板 | 僅標準函式庫 |

## 使用方式

| 步驟 | 指令／動作 | 備註 |
|---|---|---|
| 1. 安裝套件 | `pip install -r requirements.txt` | 未指定版本 |
| 2. 需 OCR 的工具 | 依 [docs/tesseract-setup.md](docs/tesseract-setup.md) 安裝 | DIB01、DischargePort 兩支 |
| 3. 執行 | `python <路徑>` | 各工具為 tkinter 檔案選取介面（關稅監控除外） |

## 目錄結構

```
internal-tools/
├── order-processing/          訂單處理
│   ├── phh25_btg_po_extractor.py
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
| 檔名不含版本號 | 版本以本表「版本」欄與 Git 歷史為準 |
| 新增工具 | 放入對應類別資料夾，並在「工具總覽」補一列 |
| `cf-us-tariff-monitor/` 不可任意搬移 | 兩個 workflow 以固定路徑引用此資料夾 |
| exe 不進 repo | 打包後的 exe 請放 GitHub Releases |
