"""
出口報單資料批次抓取工具
功能：批次讀取出口報單PDF，提取 Invoice#、總毛重、總淨重、裝貨港名稱，輸出Excel
作者：AI Assistant
日期：2026/07/28
"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

try:
    import pytesseract
    from PIL import Image
    import io
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


# ===== 1. 設定與常數 =====
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
DATA_FONT = Font(name="Arial", size=11)
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

# 已知港口清單（可擴充）。value為該港口在PDF中可能出現的拼法變體
# （報單PDF常見OCR/人工鍵入誤拼，如「高雄」正確拼法KAOHSIUNG，
# 但部分報單誤植為KAOHSTUNG，此處一併比對，回傳統一的標準拼法）
KNOWN_PORTS = {
    'TAICHUNG': ['TAICHUNG'],
    'KEELUNG': ['KEELUNG'],
    'KAOHSIUNG': ['KAOHSIUNG', 'KAOHSTUNG'],
    'TAIPEI': ['TAIPEI'],
    'TAOYUAN': ['TAOYUAN'],
}


def extract_data_from_pdf(pdf_path):
    """
    從單一PDF中提取4個欄位資料
    回傳: dict 或 None（失敗時）
    """
    try:
        filename = os.path.basename(pdf_path)
        doc = fitz.open(pdf_path)
        all_text = ""
        for page in doc:
            all_text += page.get_text() + "\n"

        # 若PDF為掃描影像（無文字層），改用OCR辨識
        # 掃描影像的表格版面複雜，單一OCR設定無法兼顧所有欄位的行序準確度：
        # - 預設設定（無psm參數）：對零散單行文字（如頁尾總毛重列）辨識順序較穩定
        # - --psm 6（視為單一文字塊）：對密集表格內的多欄數字行序較準確（如Total列淨重）
        # 因此兩種都跑，供後續個別欄位各自挑選較準確的文字來源。
        used_ocr = False
        ocr_text_default = ""
        ocr_text_psm6 = ""
        if not all_text.strip():
            if not OCR_AVAILABLE:
                doc.close()
                print(f"{filename}：偵測為掃描影像，但未安裝OCR套件(pytesseract)，無法處理。")
                return None
            used_ocr = True
            for page in doc:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes('png')))
                ocr_text_default += pytesseract.image_to_string(img, lang='eng') + "\n"
                ocr_text_psm6 += pytesseract.image_to_string(img, lang='eng', config='--psm 6') + "\n"
            all_text = ocr_text_default
        doc.close()

        if not all_text.strip():
            return None

        data = {
            'invoice': '',
            'gross_weight': '',
            'net_weight': '',
            'port': ''
        }

        # ----- 1. 裝貨港名稱 -----
        # OCR辨識掃描版PDF時，港口拼字中間常被誤插空格（如 "KAOHS TUNG"），
        # 故先產生一份移除空白的比對用文字，再逐一比對港口變體。
        # 合併兩種OCR結果一併比對，避免其中一種模式漏字。
        combined_for_port = all_text + (ocr_text_psm6 if used_ocr else "")
        text_no_space = re.sub(r'\s+', '', combined_for_port)
        port_found = False
        for standard_name, variants in KNOWN_PORTS.items():
            for variant in variants:
                if variant in text_no_space:
                    data['port'] = standard_name
                    port_found = True
                    break
            if port_found:
                break
        if not port_found:
            m = re.search(r'TWTXG\s+\w+\s+\w+\s+\w+\s+([A-Z]{3,})', all_text)
            if m:
                data['port'] = m.group(1)

        # ----- 2. Invoice# (PDF內文字) -----
        m = re.search(r'\b([A-Z0-9]{2,6}-\d{6}[A-Z])\b', all_text)
        if m:
            data['invoice'] = m.group(1)

        # 2b. 備援：檔名第一個 "_" 前的字串
        if not data['invoice']:
            data['invoice'] = filename.split('_')[0]

        # ----- 3. 總淨重 (Total: 之後數行內尋找) -----
        # 掃描 Total: 之後最多10行，逐行找符合金額格式的數字。
        # 若某行明確帶有 KGM 標記（如 "1,771.00 KGM"），優先採用該行；
        # 否則沿用舊邏輯，採用第一個符合格式的數字（PHH-240223A 這類單頁報單，
        # 淨重就在 Total: 的下一行）。
        # OCR掃描版優先使用--psm 6辨識結果，表格行序較準確。
        net_weight_source = ocr_text_psm6 if used_ocr and ocr_text_psm6.strip() else all_text
        lines_for_net = net_weight_source.split('\n')
        lines = all_text.split('\n')  # 供後續總毛重等其他欄位使用
        for i, line in enumerate(lines_for_net):
            # 冒號設為可選：一般PDF文字層為 "Total:"，但OCR辨識掃描影像時
            # 常會漏掉緊鄰的冒號（僅剩 "Total"），故此處放寬比對。
            if re.search(r'\bTotal\s*[:：]?\s*$', line, re.IGNORECASE) or re.search(r'Total\s*[:：]', line, re.IGNORECASE):
                candidates = []
                kgm_value = None
                for j in range(i + 1, min(len(lines_for_net), i + 10)):
                    next_line = lines_for_net[j].strip()
                    if not next_line:
                        continue
                    # 修復被空格拆開的小數，如 "33,812.8 6" -> "33,812.86"
                    next_line_fixed = re.sub(r'(\d)\.(\d)\s+(\d)', r'\1.\2\3', next_line)
                    for m in re.finditer(r'([\d,]+\.\d+)', next_line_fixed):
                        val = m.group(1)
                        # 排除稅則號格式的誤判：稅則號如 "6303.92" 不含千分位逗號，
                        # 且整數部分恰為4碼（HTS code慣例），淨重數字通常帶千分位逗號
                        # 或整數部分不是剛好4碼，藉此過濾。
                        int_part = val.split(',')[0].split('.')[0] if ',' not in val else val.split('.')[0].replace(',', '')
                        has_comma = ',' in val
                        if not has_comma and len(int_part) == 4:
                            continue  # 疑似稅則號，跳過
                        candidates.append(val)
                        if 'KGM' in next_line_fixed.upper():
                            kgm_value = val
                if kgm_value:
                    data['net_weight'] = kgm_value
                elif candidates:
                    data['net_weight'] = candidates[0]
                break

        # ----- 4. 總毛重 -----
        # 方法A：標籤定位法。找含「總毛重」的行，數字可能與標籤同行
        # （如 PyMuPDF 把同一儲存格列黏在一起："12 PLT (365TN) 4,833.72"，
        # 該行同時含「總毛重」標籤在前段表頭），也可能在標籤下一行。
        # 因此對含「總毛重」的那一行本身、以及其後2行都嘗試抓取金額，
        # 取第一個成功命中者，並排除件數欄常見的整數格式干擾（如 "12 PLT"）。
        for i, line in enumerate(lines):
            if '總毛重' in line:
                for j in range(i, min(len(lines), i + 3)):
                    candidate = lines[j].strip()
                    candidate_fixed = re.sub(r'(\d)\.(\d)\s+(\d)', r'\1.\2\3', candidate)
                    m = re.search(r'([\d,]+\.\d+)', candidate_fixed)
                    if m:
                        data['gross_weight'] = m.group(1)
                        break
                if data['gross_weight']:
                    break

        # 方法B：找 CTN/PKG/PLT 行（總件數/單位欄）後、下方帶前導空格的獨立數字行
        # （原有備援邏輯；加入 PLT 因部分報單以「PLT」作為包裝單位，如 PHH-240223A）
        if not data['gross_weight']:
            for i, line in enumerate(lines):
                if ('CTN' in line or 'PKG' in line or 'PLT' in line) and 'PAPER' not in line:
                    for j in range(i + 1, min(len(lines), i + 5)):
                        candidate = lines[j].strip()
                        # 跳過包裝說明行如 (1259CTN+50PAPER PLT)
                        if candidate.startswith('(') and 'CTN' in candidate:
                            continue
                        m = re.match(r'^[\s]*([\d,]+\.\d+)$', candidate)
                        if m:
                            data['gross_weight'] = m.group(1)
                            break
                    if data['gross_weight']:
                        break

        # 方法D：OCR掃描版備援。中文標籤「總毛重」在OCR結果中常變成亂碼，
        # 改以總件數行（如 "1,401 CTN"）本身作為定位點。OCR常見誤讀
        # CTN→CIN/CTM等，故以較寬鬆的 "數字+3字英文碼" 型態比對；
        # 又件數與毛重可能同一行黏在一起（如 "1,401 CTN ... 14,864.73"），
        # 故先在同一行找該行最後一個金額數字，找不到才往下數行找獨立數字。
        if not data['gross_weight']:
            for i, line in enumerate(lines):
                if re.search(r'\d[\d,]*\s*[A-Z]{3}\b', line) and re.search(r'(?i)\d\s*(CTN|CIN|CTM|PKG|PLT)', line):
                    same_line_matches = re.findall(r'([\d,]+\.\d+)', line)
                    if same_line_matches:
                        data['gross_weight'] = same_line_matches[-1]
                        break
                    for j in range(i + 1, min(len(lines), i + 8)):
                        candidate = lines[j].strip()
                        m = re.match(r'^([\d,]+\.\d+)$', candidate)
                        if m:
                            data['gross_weight'] = m.group(1)
                            break
                    if data['gross_weight']:
                        break

        # 方法C：傳統regex（總毛重標籤為可提取文字時的最後備援）
        if not data['gross_weight']:
            gross_patterns = [
                r'總毛重\s*\(?公斤\)?\s*\(?44\)?\s*[:：]?\s*([\d,]+\.?\d*)',
                r'總毛重.*?\(?44\)?.*?([\d,]+\.?\d+)',
                r'\(?44\)?.*?總毛重.*?([\d,]+\.?\d+)',
                r'GROSS\s*WEIGHT.*?([\d,]+\.?\d+)',
            ]
            for pat in gross_patterns:
                m = re.search(pat, all_text, re.IGNORECASE)
                if m:
                    val = m.group(1).strip()
                    if val:
                        data['gross_weight'] = val
                        break

        return data

    except Exception as e:
        print(f"讀取 {os.path.basename(pdf_path)} 時發生錯誤: {e}")
        return None


def create_excel(output_path, records):
    """
    建立Excel檔案
    records: list of dict，每個dict含 invoice, gross_weight, net_weight, port, status
             status為 'ok' 表示正常取得資料；'failed' 表示該檔案處理失敗（無法取得資料）
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "出口報單資料"

    # 欄位順序：invoice#, 總毛重, 總淨重, 裝貨港名稱, 備註
    headers = ['Invoice#', '總毛重', '總淨重', '裝貨港名稱', '備註']
    keys = ['invoice', 'gross_weight', 'net_weight', 'port']

    NUMBER_FORMAT = '#,##0.00'

    # 寫入標題列
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # 寫入資料列（成功與失敗皆列出，不反黃）
    for row_idx, record in enumerate(records, 2):
        is_failed = record.get('status') == 'failed'

        for col, key in enumerate(keys, 1):
            value = record.get(key, '')
            cell = ws.cell(row=row_idx, column=col)
            if key in ('gross_weight', 'net_weight'):
                # 毛重/淨重：轉為數字格式（千分位+小數點後兩位），文字轉數值失敗則留空
                if value not in ('', None):
                    try:
                        num_value = float(str(value).replace(',', ''))
                        cell.value = num_value
                        cell.number_format = NUMBER_FORMAT
                    except ValueError:
                        cell.value = value
                else:
                    cell.value = ''
            else:
                cell.value = value
            cell.font = DATA_FONT
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER

        # 備註欄：處理失敗則註明「未取得資料」
        note_cell = ws.cell(row=row_idx, column=5, value='未取得資料' if is_failed else '')
        note_cell.font = DATA_FONT
        note_cell.alignment = CENTER_ALIGN
        note_cell.border = THIN_BORDER

    # 自動調整欄寬
    col_widths = [18, 15, 15, 15, 15]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width

    # 凍結首列
    ws.freeze_panes = 'A2'

    wb.save(output_path)
    return True


def main():
    # 隱藏主視窗
    root = tk.Tk()
    root.withdraw()

    # 選擇PDF檔案（多選）
    pdf_files = filedialog.askopenfilenames(
        title="選擇出口報單 PDF 檔案（可複選）",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )

    if not pdf_files:
        messagebox.showinfo("提示", "未選擇任何檔案，程式結束。")
        return

    # 決定輸出路徑（取第一個PDF的資料夾）
    first_dir = os.path.dirname(pdf_files[0])
    output_path = os.path.join(first_dir, "出口報單彙總.xlsx")

    # 處理每個PDF
    records = []
    failed_count = 0

    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        print(f"處理中: {filename} ...")
        data = extract_data_from_pdf(pdf_path)
        if data:
            data['status'] = 'ok'
            records.append(data)
        else:
            # 處理失敗（例如純圖檔PDF且無OCR支援）：仍列入Excel。
            # Invoice#欄取檔名第一個"_"前的字串（與正常解析時的備援規則一致），
            # 其餘欄位留空，備註欄由create_excel標註「未取得資料」
            records.append({
                'invoice': filename.split('_')[0],
                'gross_weight': '',
                'net_weight': '',
                'port': '',
                'status': 'failed'
            })
            failed_count += 1

    if not records:
        messagebox.showerror("錯誤", "無法從選取的PDF中提取任何資料。")
        return

    # 建立Excel
    try:
        create_excel(output_path, records)
        ok_count = len(records) - failed_count
        msg = f"已成功輸出！\n\n輸出路徑:\n{output_path}\n\n共處理 {len(records)} 筆，成功 {ok_count} 筆。"
        if failed_count:
            msg += f"\n\n{failed_count} 筆處理失敗，已於Excel備註欄標註「未取得資料」。"
        messagebox.showinfo("完成", msg)
        print(f"輸出完成: {output_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"輸出Excel時發生錯誤:\n{e}")


if __name__ == "__main__":
    main()
