# Tesseract OCR 安裝說明

| 項目 | 內容 |
|---|---|
| 版本（原記錄） | 5.5.3 |
| 安裝路徑 | `C:\Program Files\Tesseract-OCR` |
| 可執行檔 | `C:\Program Files\Tesseract-OCR\tesseract.exe` |

## 需要 Tesseract 的程式

| 程式 | 路徑 | 路徑處理方式 |
|---|---|---|
| DIB01 訂單擷取 | `order-processing/dib-po-extractor/PO_PDF_Extractor_DIB01_V1.1.py` | 寫死 `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| DischargePort（ChatGPT 版） | `ghg-inventory/取得DischargePort_ChatGPT V1.0.py` | 寫死 `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| DischargePort（Claude 版） | `ghg-inventory/取得DischargePort_Claude V1.0.py` | 先查 PATH，再依序嘗試 `Program Files`、`Program Files (x86)`、`%LOCALAPPDATA%\Programs` |

## 注意事項

| 項目 | 說明 |
|---|---|
| 安裝位置不同 | 寫死路徑的兩支程式，Tesseract 裝在其他位置時需改程式碼中的 `tesseract_cmd` |
| pdf2image 相依 | DischargePort 兩支使用 `pdf2image`，另需安裝 Poppler 並加入 PATH [需查證：實際環境的 Poppler 安裝位置] |
