# ============================================================
#  ui/app.py — Giao diện Tkinter
#  Chỉ xử lý hiển thị; mọi logic nghiệp vụ gọi qua backend/
# ============================================================

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from function.exporter import save_csv, save_excel
from function.pdf_reader import extract_from_pdf, find_pdf_files
from config import (
    APP_GEOMETRY,
    APP_TITLE,
    COLOR_BG,
    COLOR_BTN_BLUE,
    COLOR_BTN_GRAY,
    COLOR_BTN_GREEN,
    COLOR_BTN_ORANGE,
    COLOR_PRIMARY,
    COLOR_ROW_ERROR,
    COLOR_ROW_EVEN,
    DEFAULT_OUTPUT_FOLDER,
    DEFAULT_PDF_FOLDER,
    OUTPUT_COLUMNS,
    OUTPUT_CSV_NAME,
    OUTPUT_XLSX_NAME,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.resizable(True, True)
        self.configure(bg=COLOR_BG)
        self.results_data: list[dict] = []
        self._build_ui()

    # ── UI construction ──────────────────────────────────────

    def _build_ui(self):
        self._build_titlebar()
        self._build_folder_row("  📁  Thư mục chứa file PDF  ", DEFAULT_PDF_FOLDER,  "folder_var", "_browse_folder",  top_pad=14)
        self._build_folder_row("  💾  Thư mục lưu kết quả  ",   DEFAULT_OUTPUT_FOLDER, "output_var", "_browse_output", top_pad=6)
        self._build_buttons()
        self._build_table()
        self._build_statusbar()

    def _build_titlebar(self):
        bar = tk.Frame(self, bg=COLOR_PRIMARY, height=55)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(
            bar, text="📄  Invoice Extractor",
            bg=COLOR_PRIMARY, fg="white",
            font=("Arial", 16, "bold"),
        ).pack(side="left", padx=18, pady=12)

    def _build_folder_row(self, label: str, default: str, var_attr: str, cmd_attr: str, top_pad: int):
        frame = tk.LabelFrame(
            self, text=label,
            bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_PRIMARY,
            padx=10, pady=8,
        )
        frame.pack(fill="x", padx=16, pady=(top_pad, 6))

        string_var = tk.StringVar(value=default)
        setattr(self, var_attr, string_var)

        tk.Entry(frame, textvariable=string_var, font=("Arial", 9), width=70).pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        tk.Button(
            frame, text="Chọn...",
            command=getattr(self, cmd_attr),
            bg=COLOR_BTN_BLUE, fg="white",
            font=("Arial", 9, "bold"),
            relief="flat", cursor="hand2", padx=10,
        ).pack(side="left")

    def _build_buttons(self):
        frame = tk.Frame(self, bg=COLOR_BG)
        frame.pack(fill="x", padx=16, pady=8)

        self.scan_btn = tk.Button(
            frame, text="🔍  Quét & Xử lý PDF",
            command=self._start_scan,
            bg=COLOR_BTN_GREEN, fg="white",
            font=("Arial", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=8,
        )
        self.scan_btn.pack(side="left", padx=(0, 10))

        self.export_btn = tk.Button(
            frame, text="📊  Xuất Excel",
            command=self._export_excel,
            bg=COLOR_BTN_ORANGE, fg="white",
            font=("Arial", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=8,
            state="disabled",
        )
        self.export_btn.pack(side="left")

        tk.Button(
            frame, text="🗑  Xóa",
            command=self._clear_results,
            bg=COLOR_BTN_GRAY, fg="white",
            font=("Arial", 11),
            relief="flat", cursor="hand2", padx=14, pady=8,
        ).pack(side="right")

    def _build_table(self):
        frame = tk.LabelFrame(
            self, text="  📋  Kết quả  ",
            bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_PRIMARY,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=(4, 6))

        col_widths = {"File": 260, "Invoice Date": 120, "Invoice Number": 150, "Shipment ID / ASN": 150}

        self.tree = ttk.Treeview(frame, columns=OUTPUT_COLUMNS, show="headings", height=10)
        for col in OUTPUT_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 140), anchor="w" if col == "File" else "center")

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background=COLOR_PRIMARY)
        style.configure("Treeview", font=("Arial", 9), rowheight=22)

        sb_y = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        sb_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.tag_configure("evenrow", background=COLOR_ROW_EVEN)
        self.tree.tag_configure("oddrow",  background="#FFFFFF")
        self.tree.tag_configure("error",   background=COLOR_ROW_ERROR)

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="Sẵn sàng.")
        tk.Label(
            self, textvariable=self.status_var,
            bg=COLOR_PRIMARY, fg="white",
            font=("Arial", 9), anchor="w", padx=12,
        ).pack(fill="x", side="bottom")

    # ── Event handlers ───────────────────────────────────────

    def _browse_folder(self):
        path = filedialog.askopenfilename(title="Chọn thư mục PDF")
        if path:
            self.folder_var.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if path:
            self.output_var.set(path)

    def _start_scan(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục chứa file PDF!")
            return
        if not os.path.exists(folder):
            messagebox.showerror("Lỗi", f"Thư mục không tồn tại:\n{folder}")
            return
        self.scan_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self._clear_results(keep_status=True)
        threading.Thread(target=self._scan_worker, args=(folder,), daemon=True).start()

    def _scan_worker(self, folder: str):
        """Chạy trên thread riêng; cập nhật UI qua self.after()."""
        pdf_files = find_pdf_files(folder)
        if not pdf_files:
            self.after(0, lambda: self._set_status(f"❌  Không tìm thấy file PDF trong: {folder}"))
            self.after(0, lambda: self.scan_btn.config(state="normal"))
            return

        total = len(pdf_files)
        self.after(0, lambda: self._set_status(f"⏳  Đang xử lý {total} file PDF..."))

        results: list[dict] = []
        for idx, pdf_path in enumerate(pdf_files, start=1):
            name = os.path.basename(pdf_path)
            self.after(0, lambda i=idx, t=total, n=name: self._set_status(f"⏳  [{i}/{t}]  Đang đọc: {n}"))

            record = extract_from_pdf(pdf_path)
            if record:
                results.append(record)
                self.after(0, lambda r=record, i=idx: self._append_row(r, i))

        self.results_data = results
        ok        = sum(1 for r in results if "error" not in r)
        err_count = len(results) - ok
        skipped   = total - len(results)

        msg = f"✅  Xong! {ok} file có dữ liệu"
        if err_count: msg += f"  |  ⚠ {err_count} lỗi đọc file"
        if skipped:   msg += f"  |  ℹ {skipped} file không có trang Commercial Invoice"

        self.after(0, lambda: self._set_status(msg))
        self.after(0, lambda: self.scan_btn.config(state="normal"))
        if results:
            self.after(0, lambda: self.export_btn.config(state="normal"))

    def _append_row(self, record: dict, idx: int):
        tag = "error" if "error" in record else ("evenrow" if idx % 2 == 0 else "oddrow")
        values = tuple(
            record.get(col, record.get("error", "") if col == "Invoice Date" else "")
            for col in OUTPUT_COLUMNS
        )
        self.tree.insert("", "end", values=values, tags=(tag,))

    def _export_excel(self):
        if not self.results_data:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
            return

        output_dir = self.output_var.get().strip() or DEFAULT_OUTPUT_FOLDER
        os.makedirs(output_dir, exist_ok=True)

        csv_path  = os.path.join(output_dir, OUTPUT_CSV_NAME)
        xlsx_path = os.path.join(output_dir, OUTPUT_XLSX_NAME)
        clean     = [r for r in self.results_data if "error" not in r]

        try:
            save_csv(clean,  csv_path)
            save_excel(clean, xlsx_path)
            messagebox.showinfo(
                "Xuất thành công ✅",
                f"Đã lưu:\n\n📄  CSV:   {csv_path}\n📊  Excel: {xlsx_path}",
            )
            self._set_status(f"✅  Đã xuất: {xlsx_path}")
        except Exception as exc:
            messagebox.showerror("Lỗi xuất file", str(exc))

    def _clear_results(self, keep_status: bool = False):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results_data = []
        self.export_btn.config(state="disabled")
        if not keep_status:
            self._set_status("Sẵn sàng.")

    def _set_status(self, msg: str):
        self.status_var.set(msg)