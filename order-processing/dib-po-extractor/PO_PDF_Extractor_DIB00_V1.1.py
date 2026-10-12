import re
import fitz  # PyMuPDF
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# DIB00 PO PDF Extractor
# 2026/08/17
#
# 功能簡述：
# 1. 專門處理最一開始的 DIB00 PO PDF 格式
# 2. 可一次選取多份 PDF
# 3. 從 PDF 擷取：
#    - File Name
#    - PO#
#    - Requested Ship Date
#    - Item
#    - Description
#    - Model Number
#    - Quantity
#    - Unit Of Measure
#    - UPC Code
#    - Case Pack
#    - Unit Price
#    - Extended Price
#    - QTY Per Carton
#    - Unit Cube
#    - Unit Weight
# 4. 輸出成 Excel 總表
# 5. 輸出 Excel 位置可由使用者自行選擇
# 6. 預設輸出檔名前綴為 DIB00_
# 7. 產品明細僅要求 Item / Quantity / Unit Price / Extended Price 必須有資料
#    其餘欄位若 PDF 無資料則輸出空白
# ============================================================


def extract_pdf_text(pdf_path: str) -> str:
    """
    功能：
    1. 讀取 PDF 文字內容。
    2. 此 DIB00 格式屬於文字型 PDF，所以直接使用 PyMuPDF 擷取文字。
    """

    text_list = []

    with fitz.open(pdf_path) as doc:
        for page in doc:
            text_list.append(page.get_text("text"))

    return "\n".join(text_list)


def clean_text(text: str) -> str:
    """
    功能：
    1. 清理 PDF 取得的文字。
    2. 統一換行與空白，方便後續正規表示式比對。
    """

    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)

    return text.strip()


def extract_po_no(text: str) -> str:
    """
    功能：
    擷取 PO#。

    範例：
    Order Number DIB243803-ARCHIVE

    輸出：
    DIB243803
    """

    m = re.search(r"Order Number\s+([A-Z0-9]+)", text, re.I)

    return m.group(1).strip() if m else ""


def extract_requested_ship_date(text: str) -> str:
    """
    功能：
    擷取 Requested Ship Date，並轉成 yyyy/mm/dd。

    範例：
    Requested Ship Date 05/28/2026

    輸出：
    2026/05/28
    """

    m = re.search(r"Requested Ship Date\s+(\d{2})/(\d{2})/(\d{4})", text, re.I)

    if not m:
        return ""

    mm, dd, yyyy = m.groups()

    return f"{yyyy}/{mm}/{dd}"


def extract_product_lines(text: str) -> list[dict]:
    """
    功能：
    擷取產品明細。

    修正原則：
    1. Item / Quantity / Unit Price / Extended Price 為必要資料。
    2. Description / Model Number / UOM / UPC / Case Pack /
       QTY Per Carton / Unit Cube / Unit Weight 若 PDF 沒有資料，
       則輸出空白，不因缺值造成整筆產品明細解析失敗。
    3. Unit Price / Extended Price 以 Quantity × Unit Price = Extended Price
       的關係判斷，避免 Case Pack、Unit Cube、Unit Weight 缺值時欄位錯位。
    4. UPC Code 若有資料仍保留前導 0。
    """

    # 先抓到每筆產品的穩定核心欄位：
    # Line Number / Item / Description / Model Number / Quantity / UOM / UPC
    # UPC 後方的欄位可能有空白，因此不在此階段強制要求。
    core_pattern = re.compile(
        r"""
        (?P<line_no>\d+)\s+
        (?P<item>\d{6})\s+
        (?P<description>.+?)\s+
        (?P<model>\d{6})\s+
        (?P<qty>\d+)\s+
        (?P<uom>EA)\s+
        (?P<upc>\d{14})
        """,
        re.I | re.S | re.X
    )

    matches = list(core_pattern.finditer(text))
    products = []

    # 擷取 UPC 後方可能存在的數值欄位
    number_pattern = re.compile(
        r"(?<![\d.])\d[\d,]*(?:\.\d+)?(?![\d.])"
    )

    for index, m in enumerate(matches):
        # 每筆產品的尾端，以「下一筆產品起點」作為界線。
        # 最後一筆則讀到全文結尾。
        block_end = (
            matches[index + 1].start()
            if index + 1 < len(matches)
            else len(text)
        )

        tail_text = text[m.end():block_end]
        number_tokens = [
            token.group(0)
            for token in number_pattern.finditer(tail_text)
        ]

        qty = int(m.group("qty"))
        unit_price = None
        extended_price = None
        price_index = None

        # PDF 可能缺 Case Pack / Unit Cube / Unit Weight。
        # 以「Quantity × Unit Price = Extended Price」找出真正的價格欄位，
        # 不依賴前後欄位是否存在。
        for i in range(len(number_tokens) - 1):
            try:
                candidate_unit_price = float(
                    number_tokens[i].replace(",", "")
                )
                candidate_extended_price = float(
                    number_tokens[i + 1].replace(",", "")
                )
            except ValueError:
                continue

            if abs(
                qty * candidate_unit_price - candidate_extended_price
            ) < 0.011:
                unit_price = candidate_unit_price
                extended_price = candidate_extended_price
                price_index = i
                break

        # Item / Quantity / Unit Price / Extended Price 為必要資料。
        # 找不到價格關係時，不把該筆視為有效產品明細。
        if unit_price is None or extended_price is None:
            continue

        desc = re.sub(
            r"\s+",
            " ",
            m.group("description")
        ).strip()

        case_pack = ""
        unit_cube = ""
        unit_weight = ""

        # 若 Unit Price 前面剛好有一個整數，視為 Case Pack。
        # 若沒有資料則維持空白。
        if price_index is not None and price_index > 0:
            previous_token = number_tokens[price_index - 1]

            if re.fullmatch(r"\d+", previous_token):
                case_pack = int(previous_token)

        # Extended Price 後方若緊接兩個數值，依原格式視為
        # Unit Cube / Unit Weight；沒有資料則維持空白。
        if price_index is not None:
            remaining_tokens = number_tokens[price_index + 2:]

            if len(remaining_tokens) >= 2:
                try:
                    unit_cube = float(
                        remaining_tokens[0].replace(",", "")
                    )
                    unit_weight = float(
                        remaining_tokens[1].replace(",", "")
                    )
                except ValueError:
                    unit_cube = ""
                    unit_weight = ""

        products.append({
            "Item": m.group("item"),
            "Description": desc,
            "Model Number": m.group("model"),
            "Quantity": qty,
            "Unit Of Measure": m.group("uom"),
            "UPC Code": m.group("upc"),
            "Case Pack": case_pack,
            "Unit Price": unit_price,
            "Extended Price": extended_price,
            "QTY Per Carton": "",
            "Unit Cube": unit_cube,
            "Unit Weight": unit_weight,
        })

    return products


def export_to_excel(all_rows: list[dict], output_path: str):
    """
    功能：
    將多份 PO PDF 的解析結果輸出成同一份 Excel 總表。
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "PO_Data"

    # ===== 1. 建立 Excel 抬頭 =====
    headers = [
        "File Name",
        "PO#",
        "Requested Ship Date",
        "Item",
        "Description",
        "Model Number",
        "Quantity",
        "Unit Of Measure",
        "UPC Code",
        "Case Pack",
        "Unit Price",
        "Extended Price",
        "QTY Per Carton",
        "Unit Cube",
        "Unit Weight"
    ]

    ws.append(headers)

    # ===== 2. 寫入資料 =====
    for row in all_rows:
        ws.append([
            row["File Name"],
            row["PO#"],
            row["Requested Ship Date"],
            row["Item"],
            row["Description"],
            row["Model Number"],
            row["Quantity"],
            row["Unit Of Measure"],
            row["UPC Code"],
            row["Case Pack"],
            row["Unit Price"],
            row["Extended Price"],
            row["QTY Per Carton"],
            row["Unit Cube"],
            row["Unit Weight"],
        ])

    # ===== 3. 設定 Excel 樣式 =====
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

    # ===== 4. 設定數字格式 =====
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 9).number_format = "@"          # UPC Code 保留前導 0
        ws.cell(row, 11).number_format = "#,##0.00"  # Unit Price
        ws.cell(row, 12).number_format = "#,##0.00"  # Extended Price
        ws.cell(row, 14).number_format = "0.0000"    # Unit Cube
        ws.cell(row, 15).number_format = "0.000"     # Unit Weight

    # ===== 5. 固定抬頭與篩選 =====
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ===== 6. 自動調整欄寬 =====
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = 0

        for cell in ws[col_letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    # ===== 7. 儲存 Excel =====
    wb.save(output_path)

    if not Path(output_path).exists():
        raise FileNotFoundError(f"Excel 檔案未成功產生：{output_path}")


def parse_po(pdf_path: str) -> list[dict]:
    """
    功能：
    單一 PDF 的解析流程。

    流程：
    1. 讀取 PDF 文字
    2. 清理文字
    3. 擷取 PO#
    4. 擷取 Requested Ship Date
    5. 擷取產品明細
    6. 補上 File Name / PO# / Requested Ship Date
    """

    raw_text = extract_pdf_text(pdf_path)
    text = clean_text(raw_text)

    po_no = extract_po_no(text)
    ship_date = extract_requested_ship_date(text)
    products = extract_product_lines(text)

    if not po_no:
        raise ValueError("找不到 PO#")

    if not ship_date:
        raise ValueError("找不到 Requested Ship Date")

    if not products:
        raise ValueError("找不到產品明細")

    file_name = Path(pdf_path).name

    for product in products:
        product["File Name"] = file_name
        product["PO#"] = po_no
        product["Requested Ship Date"] = ship_date

    return products


def main():
    """
    功能：
    程式主流程。

    流程：
    1. 選取一份或多份 PDF
    2. 選擇 Excel 輸出位置
    3. 逐份解析 PDF
    4. 合併輸出成 Excel
    5. 顯示成功或失敗訊息
    """

    root = Tk()
    root.withdraw()

    # ===== 1. 選取 PDF 檔案，可多選 =====
    pdf_paths = filedialog.askopenfilenames(
        title="請選擇一份或多份 DIB00 PO PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_paths:
        return

    # ===== 2. 選擇 Excel 輸出位置 =====
    output_path = filedialog.asksaveasfilename(
        title="請選擇 Excel 輸出位置",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile="DIB00_PO_Extract_All.xlsx"
    )

    if not output_path:
        return

    all_rows = []
    error_list = []

    # ===== 3. 逐份 PDF 解析 =====
    for pdf_path in pdf_paths:
        try:
            rows = parse_po(pdf_path)
            all_rows.extend(rows)

        except Exception as e:
            error_list.append(f"{Path(pdf_path).name}：{e}")

    # ===== 4. 若全部失敗，不產生 Excel =====
    if not all_rows:
        messagebox.showerror(
            "錯誤",
            "所有 PDF 都讀取失敗，未產生 Excel。\n\n" + "\n".join(error_list)
        )
        return

    # ===== 5. 輸出 Excel =====
    try:
        export_to_excel(all_rows, output_path)

    except Exception as e:
        messagebox.showerror(
            "輸出失敗",
            f"Excel 檔案輸出失敗：\n{e}\n\n請確認檔案沒有被開啟，且資料夾有寫入權限。"
        )
        return

    # ===== 6. 顯示完成訊息 =====
    msg = (
        f"DIB00 PO資料擷取完成！\n\n"
        f"選取檔案數：{len(pdf_paths)}\n"
        f"成功產品筆數：{len(all_rows)}\n\n"
        f"輸出檔案：\n{output_path}"
    )

    if error_list:
        msg += "\n\n以下檔案讀取失敗：\n" + "\n".join(error_list)

    messagebox.showinfo("完成", msg)


if __name__ == "__main__":
    main()