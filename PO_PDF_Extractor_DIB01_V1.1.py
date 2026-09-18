import re
import fitz
import pytesseract
from PIL import Image
from io import BytesIO
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# DIB01 PO PDF Extractor
# 2026/08/18
#
# 功能簡述：
# 1. 專門處理 DIB01 / Maintenance Supply Solutions PO PDF 格式
# 2. 可一次選取多份 PDF
# 3. 支援掃描圖片型 PDF，使用 Tesseract OCR
# 4. 擷取 PO#、Required Date、Item、Description、Part Number、
#    Quantity、UOM、Unit Price、Pricing UOM、Pricing Unit Size、Extended Price
# 5. 輸出成 Excel 總表
# 6. Item / Quantity / Unit Price / Extended Price 為必要資料；其餘欄位可空白
# ============================================================


# ===== 0. Tesseract OCR 路徑設定 =====
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def extract_pdf_text(pdf_path: str) -> str:
    """
    功能：
    1. 讀取 PDF 文字內容。
    2. 若 PDF 本身可直接讀文字，就直接讀。
    3. 若讀不到文字，表示大概率是掃描圖片型 PDF，改用 OCR。
    """

    text_list = []

    with fitz.open(pdf_path) as doc:
        for page in doc:
            text = page.get_text("text")

            if text.strip():
                text_list.append(text)
            else:
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                img = Image.open(BytesIO(pix.tobytes("png")))

                ocr_text = pytesseract.image_to_string(
                    img,
                    lang="eng",
                    config="--psm 6"
                )

                text_list.append(ocr_text)

    return "\n".join(text_list)


def clean_text(text: str) -> str:
    """
    功能：
    清理 PDF / OCR 取得的文字。
    """

    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()


def convert_mmddyyyy_to_yyyymmdd(date_text: str) -> str:
    """
    功能：
    將 MM/DD/YYYY 轉成 YYYY/MM/DD。

    範例：
    03/17/2026 -> 2026/03/17
    """

    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", date_text)

    if not m:
        return ""

    mm, dd, yyyy = m.groups()
    return f"{yyyy}/{mm}/{dd}"


def extract_po_no(text: str) -> str:
    """
    功能：
    擷取 PO#。

    DIB01 格式位置：
    Purchase Order Number
    4008091

    邏輯：
    1. 找到 Purchase Order Number。
    2. 往下找 6~10 碼數字。
    """

    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    lines = [x for x in lines if x]

    for i, line in enumerate(lines):
        if "Purchase Order Number" in line:
            for j in range(i + 1, min(i + 8, len(lines))):
                m = re.search(r"\b(\d{6,10})\b", lines[j])
                if m:
                    return m.group(1)

    m = re.search(r"Purchase\s+Order\s+Number\s+(\d{6,10})", text, re.I)

    return m.group(1) if m else ""


def clean_description(desc: str) -> str:
    """
    功能：
    清理產品描述。
    """

    desc = re.sub(r"\s+", " ", desc).strip()
    desc = re.sub(r"\s+1\.0000$", "", desc).strip()
    desc = re.sub(r"\s+\.+$", "", desc).strip()

    return desc


def is_noise_line(line: str) -> bool:
    """
    功能：
    判斷是否為頁首、頁尾、表頭、地址等雜訊。
    """

    keywords = [
        "PURCHASE ORDER",
        "Maintenance Supply Solutions",
        "Carrier Pkwy",
        "Grand Prairie",
        "Phone:",
        "Fax:",
        "Purchase Order Number",
        "External PO Number",
        "Quantity",
        "Required Date",
        "Item ID",
        "Item Description",
        "Net Unit Price",
        "Pricing UOM",
        "Extended Price",
        "Date Page",
        "Page",
        "TOTAL:",
        "4008091"
    ]

    return any(k.upper() in line.upper() for k in keywords)


def extract_product_lines(text: str) -> list[dict]:
    """
    功能：
    擷取 DIB01 表身資料。

    修正原則：
    1. Item / Quantity / Unit Price / Extended Price 為必要資料。
    2. Required Date / UOM / Pricing UOM / Pricing Unit Size /
       Description / Part Number 若 PDF 無資料，則輸出空白。
    3. 以 Quantity × Unit Price = Extended Price 驗證價格欄位，
       避免 optional 欄位缺值造成整筆資料解析失敗。
    """

    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    lines = [x for x in lines if x]

    products = []
    main_rows = []

    item_pattern = re.compile(r"\b(?P<item>\d{6})\b")
    number_pattern = re.compile(r"\$?\d[\d,]*(?:\.\d+)?")

    for idx, line in enumerate(lines):
        if is_noise_line(line):
            continue

        item_match = item_pattern.search(line)
        if not item_match:
            continue

        item = item_match.group("item")
        before_item = line[:item_match.start()]
        after_item = line[item_match.end():]

        # Item 前方找 Quantity；若同時有日期，先移除日期避免誤判
        before_for_qty = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", " ", before_item)
        qty_candidates = [
            int(x.replace(",", ""))
            for x in re.findall(r"\b[\d,]+\b", before_for_qty)
            if re.fullmatch(r"\d+", x.replace(",", ""))
        ]

        if not qty_candidates:
            continue

        qty = qty_candidates[0]

        # Item 後方的數值中找 Unit Price / Extended Price
        numeric_tokens = [
            m.group(0).replace("$", "")
            for m in number_pattern.finditer(after_item)
        ]

        unit_price = None
        extended_price = None

        for i in range(len(numeric_tokens) - 1):
            try:
                candidate_price = float(numeric_tokens[i].replace(",", ""))
                candidate_ext = float(numeric_tokens[i + 1].replace(",", ""))
            except ValueError:
                continue

            if abs(qty * candidate_price - candidate_ext) < 0.011:
                unit_price = candidate_price
                extended_price = candidate_ext
                break

        if unit_price is None or extended_price is None:
            continue

        req_date = ""
        date_match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", before_item)
        if date_match:
            req_date = convert_mmddyyyy_to_yyyymmdd(date_match.group(1))

        uom = ""
        uom_match = re.search(r"\b([A-Z]{2})\b", before_item)
        if uom_match:
            uom = uom_match.group(1)

        pricing_uom = ""
        pricing_uom_match = re.search(
            r"\b([A-Z]{2})\b",
            after_item,
            re.I
        )
        if pricing_uom_match:
            pricing_uom = pricing_uom_match.group(1)

        main_rows.append({
            "idx": idx,
            "item": item,
            "qty": qty,
            "uom": uom,
            "req_date": req_date,
            "unit_price": unit_price,
            "pricing_uom": pricing_uom,
            "extended_price": extended_price
        })

    for n, row_info in enumerate(main_rows):
        start_idx = row_info["idx"]
        end_idx = main_rows[n + 1]["idx"] if n + 1 < len(main_rows) else len(lines)
        block_lines = lines[start_idx + 1:end_idx]

        item = row_info["item"]
        unit_size = ""
        description = ""
        part_number = ""

        desc_started = False
        desc_parts = []

        for block_line in block_lines:
            if is_noise_line(block_line):
                continue

            part_match = re.search(r"Part\s*No\.?\s*:?\s*(\d{6})", block_line, re.I)
            if part_match:
                part_number = part_match.group(1)
                desc_started = False
                continue

            desc_match = re.match(
                r"^(?P<unit_size>1(?:\.0{1,4})?)\s+(?P<desc>.+)$",
                block_line,
                re.I
            )

            if desc_match:
                possible_desc = clean_description(desc_match.group("desc"))
                if "BLIND" in possible_desc.upper():
                    unit_size = desc_match.group("unit_size")
                    desc_parts.append(possible_desc)
                    desc_started = True
                continue

            if desc_started:
                extra_desc = clean_description(block_line)

                if not extra_desc:
                    continue
                if re.search(r"Part\s*No\.?", extra_desc, re.I):
                    desc_started = False
                    continue
                if re.fullmatch(r"\d+", extra_desc):
                    continue
                if re.fullmatch(r"1(?:\.0{1,4})?", extra_desc):
                    continue
                if re.search(r"\d{2}/\d{2}/\d{4}", extra_desc):
                    continue
                if re.search(r"\b\d{6}\b", extra_desc):
                    continue
                if re.search(r"^[\d,]+\.\d{2,4}$", extra_desc):
                    continue
                if is_noise_line(extra_desc):
                    continue

                desc_parts.append(extra_desc)

        if desc_parts:
            description = clean_description(" ".join(desc_parts))

        if not part_number:
            part_number = item

        products.append({
            "Item": item,
            "Description": description,
            "Part Number": part_number,
            "Quantity": row_info["qty"],
            "Unit Of Measure": row_info["uom"],
            "Required Date": row_info["req_date"],
            "Unit Price": row_info["unit_price"],
            "Pricing UOM": row_info["pricing_uom"],
            "Pricing Unit Size": unit_size,
            "Extended Price": row_info["extended_price"],
        })

    return products


def parse_po(pdf_path: str) -> list[dict]:
    """
    功能：
    單一 DIB01 PDF 的主解析流程。
    """

    raw_text = extract_pdf_text(pdf_path)
    text = clean_text(raw_text)

    po_no = extract_po_no(text)
    products = extract_product_lines(text)

    if not po_no:
        raise ValueError("找不到 PO#")

    if not products:
        raise ValueError("找不到產品明細")

    file_name = Path(pdf_path).name

    for p in products:
        p["File Name"] = file_name
        p["PO#"] = po_no

    return products


def export_to_excel(all_rows: list[dict], output_path: str):
    """
    功能：
    將解析結果輸出 Excel。
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "PO_Data"

    headers = [
        "File Name",
        "PO#",
        "Required Date",
        "Item",
        "Description",
        "Part Number",
        "Quantity",
        "Unit Of Measure",
        "Unit Price",
        "Pricing UOM",
        "Pricing Unit Size",
        "Extended Price",
    ]

    ws.append(headers)

    for row in all_rows:
        ws.append([
            row.get("File Name", ""),
            row.get("PO#", ""),
            row.get("Required Date", ""),
            row.get("Item", ""),
            row.get("Description", ""),
            row.get("Part Number", ""),
            row.get("Quantity", ""),
            row.get("Unit Of Measure", ""),
            row.get("Unit Price", ""),
            row.get("Pricing UOM", ""),
            row.get("Pricing Unit Size", ""),
            row.get("Extended Price", ""),
        ])

    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial")
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 9).number_format = "#,##0.000"
        ws.cell(row, 12).number_format = "#,##0.00"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = 0

        for cell in ws[col_letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = min(max_len + 2, 55)

    wb.save(output_path)

    if not Path(output_path).exists():
        raise FileNotFoundError(f"Excel 檔案未成功產生：{output_path}")


def main():
    """
    功能：
    程式進入點。
    """

    root = Tk()
    root.withdraw()

    # ===== 檢查 Tesseract 是否存在 =====
    tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    if not Path(tesseract_path).exists():
        messagebox.showerror(
            "錯誤",
            "未安裝Tesseract，請check Joe安裝！"
        )
        return

    # 若存在，才設定路徑
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

    pdf_paths = filedialog.askopenfilenames(
        title="請選擇一份或多份 DIB01 PO PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_paths:
        return

    output_path = filedialog.asksaveasfilename(
        title="請選擇 Excel 輸出位置",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile="DIB01_PO_Extract_All.xlsx"
    )

    if not output_path:
        return

    all_rows = []
    error_list = []

    for pdf_path in pdf_paths:
        try:
            rows = parse_po(pdf_path)
            all_rows.extend(rows)

        except Exception as e:
            error_list.append(f"{Path(pdf_path).name}：{e}")

    if not all_rows:
        messagebox.showerror(
            "錯誤",
            "所有 PDF 都讀取失敗，未產生 Excel。\n\n" + "\n".join(error_list)
        )
        return

    try:
        export_to_excel(all_rows, output_path)

    except Exception as e:
        messagebox.showerror(
            "輸出失敗",
            f"Excel 檔案輸出失敗：\n{e}\n\n請確認檔案沒有被開啟，且資料夾有寫入權限。"
        )
        return

    msg = (
        f"DIB01 PO資料擷取完成！\n\n"
        f"選取檔案數：{len(pdf_paths)}\n"
        f"成功產品筆數：{len(all_rows)}\n\n"
        f"輸出檔案：\n{output_path}"
    )

    if error_list:
        msg += "\n\n以下檔案讀取失敗：\n" + "\n".join(error_list)

    messagebox.showinfo("完成", msg)


if __name__ == "__main__":
    main()