"""出口報單批次擷取工具
'2026/07/28 16:10
- 多選PDF出口報單，擷取：Invoice#, 總毛重, 總淨重, 裝貨港名稱
- 輸出Excel至PDF所在資料夾
- 將擷取到的4個欄位在原PDF上反黃標記，覆蓋原檔
"""
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pdfplumber
import fitz  # PyMuPDF
import openpyxl


def extract_gross_weight(page1_words):
    """總毛重(44)：top≈590.7, x0>500 的數字。回傳 (文字, bbox) 或 (None, None)"""
    candidates = [w for w in page1_words if 585 < w['top'] < 596 and w['x0'] > 500]
    for w in candidates:
        if re.match(r'^[\d,]+\.\d+$', w['text']):
            return w['text'], (w['x0'], w['top'], w['x1'], w['bottom'])
    return None, None


def extract_port(page1_words):
    """裝貨港名稱(8)：top≈137.7 那一行，x0介於70~165的英文字（右側為目的國家欄，需排除）
    回傳 (文字, 合併後bbox) 或 (None, None)"""
    candidates = [w for w in page1_words if 135 < w['top'] < 141 and 70 < w['x0'] < 165]
    candidates.sort(key=lambda w: w['x0'])
    if candidates:
        text = " ".join(w['text'] for w in candidates)
        bbox = (candidates[0]['x0'], candidates[0]['top'],
                candidates[-1]['x1'], candidates[0]['bottom'])
        return text, bbox
    return None, None


def extract_net_weight(pdf):
    """總淨重：逐頁找 'Total' word，取其下方3~15px範圍內、x0<480的數字。
    回傳 (文字, 頁碼index, 合併後bbox) 或 (None, None, None)"""
    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words()
        for w in words:
            if w['text'] == 'Total':
                total_top = w['top']
                below = [ww for ww in words
                         if total_top + 3 < ww['top'] < total_top + 15 and ww['x0'] < 480]
                below.sort(key=lambda w: w['x0'])
                merged = "".join(w['text'] for w in below)
                m = re.search(r'[\d,]+\.\d+', merged)
                if m:
                    bbox = (below[0]['x0'], below[0]['top'],
                            below[-1]['x1'], below[0]['bottom'])
                    return m.group(), page_idx, bbox
    return None, None, None


def extract_invoice_no(filepath):
    """Invoice#：取檔名（去除副檔名與常見後綴，含「出口報單」及其後綴符號）"""
    name = os.path.splitext(os.path.basename(filepath))[0]
    name = re.sub(r'_?出口報單[-_]*$', '', name)
    return name


def extract_one_pdf(filepath):
    with pdfplumber.open(filepath) as pdf:
        page1_words = pdf.pages[0].extract_words()

        if not page1_words:
            # 純掃描圖片PDF，無文字層可擷取，跳過並標記
            return {
                'invoice_no': extract_invoice_no(filepath),
                'gross_weight': None,
                'net_weight': None,
                'port': None,
                'note': '需人工處理（純圖片PDF，無文字層）',
            }

        gross_text, gross_bbox = extract_gross_weight(page1_words)
        port_text, port_bbox = extract_port(page1_words)
        net_text, net_page_idx, net_bbox = extract_net_weight(pdf)

    highlight_pdf(filepath, gross_bbox, port_bbox, net_page_idx, net_bbox)

    missing = [label for label, val in
               [('總毛重', gross_text), ('總淨重', net_text), ('裝貨港名稱', port_text)]
               if val is None]
    note = f"未抓到：{'、'.join(missing)}" if missing else ''

    return {
        'invoice_no': extract_invoice_no(filepath),
        'gross_weight': gross_text,
        'net_weight': net_text,
        'port': port_text,
        'note': note,
    }


def highlight_pdf(filepath, gross_bbox, port_bbox, net_page_idx, net_bbox):
    """在原PDF上對已擷取到的欄位加上黃色highlight，覆蓋原檔。抓不到的欄位略過不標記。"""
    doc = fitz.open(filepath)
    page1 = doc[0]
    if gross_bbox:
        page1.add_highlight_annot(fitz.Rect(*gross_bbox))
    if port_bbox:
        page1.add_highlight_annot(fitz.Rect(*port_bbox))
    if net_bbox is not None and net_page_idx is not None:
        doc[net_page_idx].add_highlight_annot(fitz.Rect(*net_bbox))
    doc.save(filepath, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    doc.close()


def main():
    root = tk.Tk()
    root.withdraw()

    filepaths = filedialog.askopenfilenames(
        title="選擇出口報單PDF（可多選）",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not filepaths:
        return

    total = len(filepaths)
    results = []
    for idx, fp in enumerate(filepaths, start=1):
        print(f"處理中 {idx}/{total}：{os.path.basename(fp)}")
        try:
            row = extract_one_pdf(fp)
        except Exception as e:
            row = {'invoice_no': extract_invoice_no(fp), 'gross_weight': None,
                   'net_weight': None, 'port': None, 'note': f'處理失敗：{e}'}
            print(f"[錯誤] {fp}: {e}")
        results.append(row)
    print(f"全部處理完成，共 {total} 個檔案")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出口報單擷取"
    ws.append(["Invoice#", "總毛重", "總淨重", "裝貨港名稱", "備註"])
    for r in results:
        gross = float(r['gross_weight'].replace(',', '')) if r['gross_weight'] else None
        net = float(r['net_weight'].replace(',', '')) if r['net_weight'] else None
        ws.append([r['invoice_no'], gross, net, r['port'], r.get('note', '')])

    for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            if cell.value is not None:
                cell.number_format = '#,##0.00'

    out_dir = os.path.dirname(filepaths[0])
    out_path = os.path.join(out_dir, "出口報單擷取結果.xlsx")
    try:
        wb.save(out_path)
    except PermissionError:
        import time
        out_path = os.path.join(out_dir, f"出口報單擷取結果_{int(time.time())}.xlsx")
        wb.save(out_path)
        messagebox.showwarning(
            "檔案被佔用",
            f"原檔名無法寫入（可能已開啟於Excel中），已改存為：\n{out_path}"
        )

    messagebox.showinfo("完成", f"已處理 {len(results)} 份報單\n輸出至：{out_path}")


if __name__ == "__main__":
    main()
