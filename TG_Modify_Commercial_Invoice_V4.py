# -*- coding: utf-8 -*-
import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import fitz  # PyMuPDF


OLD_TITLE = "LG SOURCING INC COMMERCIAL INVOICE"
NEW_TITLE = "ALL STRONG INDUSTRY USA INC COMMERCIAL INVOICE"

COMMERCIAL_LABEL = "COMMERCIAL INVOICE NUMBER:"
VENDOR_LABEL = "VENDOR INVOICE NUMBER:"


def get_spans(page):
    """取得頁面內所有文字 span。"""
    spans = []

    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                spans.append(span)

    return spans


def find_value_below_label(page, label_text):
    """依欄位標題，取得其正下方的欄位值。"""
    spans = get_spans(page)

    label_span = None
    for span in spans:
        if span["text"].strip() == label_text:
            label_span = span
            break

    if label_span is None:
        return None

    label_rect = fitz.Rect(label_span["bbox"])
    candidates = []

    for span in spans:
        text = span["text"].strip()
        if not text or span is label_span:
            continue

        rect = fitz.Rect(span["bbox"])

        # 僅找標題正下方、且左側位置接近的文字
        if (
            rect.y0 >= label_rect.y1 - 0.5
            and rect.y0 <= label_rect.y1 + 30
            and abs(rect.x0 - label_rect.x0) <= 10
        ):
            candidates.append(span)

    if not candidates:
        return None

    candidates.sort(
        key=lambda s: (
            fitz.Rect(s["bbox"]).y0,
            abs(fitz.Rect(s["bbox"]).x0 - label_rect.x0),
        )
    )

    return candidates[0]


def modify_pdf(input_path, output_dir):
    """修正 Commercial Invoice，並將處理後檔案存至指定資料夾。"""
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    output_path = output_dir / input_path.name
    doc = fitz.open(str(input_path))

    temp_path = None

    try:
        invoice_page_index = None
        vendor_span = None
        commercial_span = None

        # ── 1. 找出 Invoice 第一頁及 Vendor Invoice Number ─────────────
        for page_index, page in enumerate(doc):
            page_text = page.get_text()

            if OLD_TITLE not in page_text:
                continue

            current_vendor_span = find_value_below_label(page, VENDOR_LABEL)
            current_commercial_span = find_value_below_label(
                page, COMMERCIAL_LABEL
            )

            if current_vendor_span and current_commercial_span:
                invoice_page_index = page_index
                vendor_span = current_vendor_span
                commercial_span = current_commercial_span
                break

        if invoice_page_index is None:
            raise ValueError(
                "找不到同時包含 Commercial Invoice Number "
                "及 Vendor Invoice Number 的 Invoice 頁面。"
            )

        vendor_invoice_number = vendor_span["text"].strip()

        # ── 2. 找出所有 Commercial Invoice 頁面的原標題 ───────────────
        title_items = []

        for page_index, page in enumerate(doc):
            spans = get_spans(page)

            for span in spans:
                if span["text"].strip() == OLD_TITLE:
                    old_rect = fitz.Rect(span["bbox"])

                    # 找同一標題列右側最近的文字。
                    # 此位置即 Commercial Invoice 每頁頁首的舊 Invoice Number，
                    # 例如 D4088154；僅處理 Commercial Invoice，不處理 Packing List。
                    header_invoice_span = None
                    nearest_x = None

                    for other_span in spans:
                        other_rect = fitz.Rect(other_span["bbox"])

                        if (
                            other_rect.x0 > old_rect.x1
                            and abs(other_span["origin"][1] - span["origin"][1]) <= 3
                        ):
                            if nearest_x is None or other_rect.x0 < nearest_x:
                                nearest_x = other_rect.x0
                                header_invoice_span = other_span

                    title_items.append(
                        {
                            "page_index": page_index,
                            "span": span,
                            "header_invoice_span": header_invoice_span,
                        }
                    )
                    break

        if not title_items:
            raise ValueError(f'找不到標題："{OLD_TITLE}"')

        # ── 3. 先刪除原文字 ───────────────────────────────────────────
        for item in title_items:
            page = doc[item["page_index"]]

            # 刪除原 Commercial Invoice 標題
            rect = fitz.Rect(item["span"]["bbox"])
            page.add_redact_annot(rect, fill=(1, 1, 1))

            # 刪除 Commercial Invoice 每一頁標題右側的舊 Invoice Number
            if item["header_invoice_span"] is not None:
                header_number_rect = fitz.Rect(
                    item["header_invoice_span"]["bbox"]
                )
                page.add_redact_annot(
                    header_number_rect,
                    fill=(1, 1, 1),
                )

        page = doc[invoice_page_index]
        page.add_redact_annot(
            fitz.Rect(commercial_span["bbox"]),
            fill=(1, 1, 1),
        )

        affected_pages = {item["page_index"] for item in title_items}
        affected_pages.add(invoice_page_index)

        for page_index in affected_pages:
            doc[page_index].apply_redactions()

        # ── 4. 寫入新的 Invoice 標題 ─────────────────────────────────
        for item in title_items:
            page = doc[item["page_index"]]
            span = item["span"]
            old_rect = fitz.Rect(span["bbox"])

            font_size = float(span["size"])
            baseline_y = float(span["origin"][1])
            center_x = (old_rect.x0 + old_rect.x1) / 2
            # 頁首舊 Invoice Number 已移除，新標題可使用原本右側空間。
            right_limit = page.rect.width - 90

            while font_size >= 9:
                text_width = fitz.get_text_length(
                    NEW_TITLE,
                    fontname="hebo",
                    fontsize=font_size,
                )

                start_x = center_x - text_width / 2

                if start_x >= 20 and start_x + text_width <= right_limit:
                    break

                font_size -= 0.2

            page.insert_text(
                (start_x, baseline_y),
                NEW_TITLE,
                fontname="hebo",
                fontsize=font_size,
                color=(0, 0, 0),
            )

        # ── 5. 以 Vendor Invoice Number 取代 Commercial Invoice Number ─
        page = doc[invoice_page_index]

        font_size = float(commercial_span["size"])
        x = float(commercial_span["origin"][0])
        y = float(commercial_span["origin"][1])
        available_width = page.rect.width - x - 25

        while font_size >= 7:
            text_width = fitz.get_text_length(
                vendor_invoice_number,
                fontname="helv",
                fontsize=font_size,
            )

            if text_width <= available_width:
                break

            font_size -= 0.2

        page.insert_text(
            (x, y),
            vendor_invoice_number,
            fontname="helv",
            fontsize=font_size,
            color=(0, 0, 0),
        )

        # ── 6. 先存暫存檔，再寫入指定輸出資料夾 ─────────────────────
        output_dir.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            suffix=".pdf",
            prefix=input_path.stem + "_tmp_",
            dir=str(output_dir),
        )
        os.close(fd)
        temp_path = Path(temp_name)

        doc.save(
            str(temp_path),
            garbage=4,
            deflate=True,
        )

    finally:
        doc.close()

    if temp_path is None or not temp_path.exists():
        raise RuntimeError("暫存 PDF 未成功建立。")

    try:
        os.replace(str(temp_path), str(output_path))
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise

    return vendor_invoice_number, len(title_items), output_path


class PdfModifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Commercial Invoice PDF 修正工具")
        self.root.geometry("900x650")
        self.root.minsize(820, 580)

        self.files = []
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        title = ttk.Label(
            main,
            text="Commercial Invoice PDF 修正工具",
            font=("Microsoft JhengHei", 16, "bold"),
        )
        title.pack(anchor="w")

        desc = ttk.Label(
            main,
            text=(
                "功能：將 Invoice 標題改為 ALL STRONG INDUSTRY USA INC COMMERCIAL INVOICE，"
                "以 VENDOR INVOICE NUMBER 取代 COMMERCIAL INVOICE NUMBER，並移除每頁 Invoice 標題右側的舊編號。"
            ),
            wraplength=780,
            font=("Microsoft JhengHei", 10),
        )
        desc.pack(anchor="w", pady=(6, 14))

        button_frame = ttk.Frame(main)
        button_frame.pack(fill="x", pady=(0, 8))

        ttk.Button(
            button_frame,
            text="選取 PDF 檔案",
            command=self.select_files,
            width=18,
        ).pack(side="left")

        ttk.Button(
            button_frame,
            text="清除清單",
            command=self.clear_files,
            width=12,
        ).pack(side="left", padx=(8, 0))

        self.file_count_var = tk.StringVar(value="已選取 0 個檔案")
        ttk.Label(
            button_frame,
            textvariable=self.file_count_var,
            font=("Microsoft JhengHei", 10),
        ).pack(side="right")

        list_frame = ttk.LabelFrame(main, text="待處理檔案", padding=8)
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            list_frame,
            font=("Consolas", 10),
            selectmode=tk.EXTENDED,
        )
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.listbox.yview,
        )
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        output_frame = ttk.LabelFrame(
            main,
            text="處理後檔案存放路徑",
            padding=8,
        )
        output_frame.pack(fill="x", pady=(12, 0))

        output_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_dir,
            state="readonly",
            font=("Microsoft JhengHei", 10),
        )
        output_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(
            output_frame,
            text="選擇資料夾",
            command=self.select_output_dir,
            width=14,
        ).pack(side="left", padx=(8, 0))

        progress_frame = ttk.Frame(main)
        progress_frame.pack(fill="x", pady=(12, 0))

        self.progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
        )
        self.progress.pack(fill="x")

        self.status_var = tk.StringVar(value="請先選取 PDF 檔案及輸出資料夾。")
        ttk.Label(
            main,
            textvariable=self.status_var,
            font=("Microsoft JhengHei", 10),
        ).pack(anchor="w", pady=(6, 10))

        style = ttk.Style()
        style.configure(
            "Run.TButton",
            font=("Microsoft JhengHei", 13, "bold"),
            padding=(12, 14),
        )

        self.run_button = ttk.Button(
            main,
            text="開始執行",
            command=self.run_batch,
            style="Run.TButton",
        )
        self.run_button.pack(fill="x", ipady=8)

        warning = ttk.Label(
            main,
            text="處理後檔案會以原檔名存入指定資料夾；若同名檔案已存在，將直接覆蓋。",
            foreground="#B00020",
            font=("Microsoft JhengHei", 9),
        )
        warning.pack(anchor="w", pady=(8, 0))

    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="請選擇一個或多個 PDF 檔案",
            filetypes=[("PDF files", "*.pdf")],
        )

        if not paths:
            return

        existing = {str(Path(p).resolve()).lower() for p in self.files}

        for path in paths:
            resolved = str(Path(path).resolve())

            if resolved.lower() not in existing:
                self.files.append(resolved)
                existing.add(resolved.lower())

        self.refresh_file_list()
        self.status_var.set("檔案已加入清單。")

    def clear_files(self):
        self.files.clear()
        self.refresh_file_list()
        self.progress["value"] = 0
        self.status_var.set("檔案清單已清除。")

    def select_output_dir(self):
        path = filedialog.askdirectory(
            title="請選擇處理後 PDF 的存放資料夾"
        )

        if not path:
            return

        self.output_dir.set(str(Path(path).resolve()))
        self.status_var.set("輸出資料夾已設定。")

    def refresh_file_list(self):
        self.listbox.delete(0, tk.END)

        for path in self.files:
            self.listbox.insert(tk.END, path)

        self.file_count_var.set(f"已選取 {len(self.files)} 個檔案")

    def run_batch(self):
        if not self.files:
            messagebox.showwarning(
                "尚未選取檔案",
                "請先選取要處理的 PDF 檔案。",
            )
            return

        if not self.output_dir.get():
            messagebox.showwarning(
                "尚未選擇輸出資料夾",
                "請先選擇處理後檔案的存放位置。",
            )
            return

        # 防止不同來源資料夾中有同名 PDF，輸出時彼此覆蓋。
        file_names = [Path(path).name.lower() for path in self.files]
        duplicate_names = sorted({
            name for name in file_names
            if file_names.count(name) > 1
        })

        if duplicate_names:
            messagebox.showerror(
                "檔名重複",
                "選取的 PDF 中有同名檔案，無法安全輸出到同一資料夾：\n\n"
                + "\n".join(duplicate_names)
            )
            return

        confirm = messagebox.askyesno(
            "確認執行",
            f"即將處理 {len(self.files)} 個 PDF。\n\n"
            f"輸出位置：\n{self.output_dir.get()}\n\n"
            "若輸出資料夾已有同名檔案，將直接覆蓋。\n"
            "是否繼續？",
        )

        if not confirm:
            return

        self.run_button.config(state="disabled")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0

        success_count = 0
        failed_items = []

        try:
            for index, path in enumerate(self.files, start=1):
                file_name = Path(path).name
                self.status_var.set(
                    f"處理中 {index}/{len(self.files)}：{file_name}"
                )
                self.root.update_idletasks()

                try:
                    vendor_invoice_number, title_count, output_path = modify_pdf(
                        path,
                        self.output_dir.get(),
                    )
                    success_count += 1

                except Exception as exc:
                    failed_items.append(
                        f"{file_name}\n  {exc}"
                    )

                self.progress["value"] = index
                self.root.update_idletasks()

        finally:
            self.run_button.config(state="normal")

        if failed_items:
            self.status_var.set(
                f"完成：成功 {success_count} 個，失敗 {len(failed_items)} 個。"
            )

            messagebox.showwarning(
                "批次處理完成",
                f"成功：{success_count} 個\n"
                f"失敗：{len(failed_items)} 個\n\n"
                "失敗檔案：\n\n" + "\n\n".join(failed_items),
            )
        else:
            self.status_var.set(
                f"完成：{success_count} 個 PDF 已全部修正並輸出完成。"
            )

            messagebox.showinfo(
                "完成",
                f"共 {success_count} 個 PDF 已完成修正。\n\n"
                f"輸出位置：\n{self.output_dir.get()}",
            )


def main():
    root = tk.Tk()
    PdfModifierApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
