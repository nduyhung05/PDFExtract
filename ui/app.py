# ============================================================
#  ui/app.py — Giao diện chính v3
# ============================================================

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from function.exporter import save_csv, save_excel
from function.pdf_reader import extract_from_pdf, search_pdf_by_names
from config import (
    APP_GEOMETRY, APP_TITLE,
    COLOR_BG, COLOR_BTN_BLUE, COLOR_BTN_GRAY, COLOR_BTN_GREEN,
    COLOR_BTN_ORANGE, COLOR_BTN_RED, COLOR_PANEL_BG, COLOR_PRIMARY,
    COLOR_ROW_ERROR, COLOR_ROW_EVEN,
    DEFAULT_OUTPUT_FOLDER, DEFAULT_PDF_FOLDER,
    OUTPUT_COLUMNS, OUTPUT_CSV_NAME, OUTPUT_XLSX_NAME,
)

ICON_PDF    = "PDF"
ICON_OK     = "OK"
ICON_ERR    = "ERR"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(880, 640)
        self.resizable(True, True)
        self.configure(bg=COLOR_BG)

        self._queued_paths: list[str] = []
        self._results: list[dict]     = []
        self._placeholder_active      = False

        self._build_ui()

    # ── UI construction ──────────────────────────────────────

    def _build_ui(self):
        self._build_titlebar()

        main = tk.Frame(self, bg=COLOR_BG)
        main.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        self._build_left_panel(main)
        self._build_right_panel(main)
        self._build_result_table()
        self._build_action_bar()
        self._build_statusbar()

    def _build_titlebar(self):
        bar = tk.Frame(self, bg=COLOR_PRIMARY, height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="  Invoice Extractor",
                 bg=COLOR_PRIMARY, fg="white", font=("Arial", 15, "bold"),
                 ).pack(side="left", padx=16, pady=10)
        tk.Label(bar, text="UA Shipping Doc",
                 bg=COLOR_PRIMARY, fg="#A8C4E0", font=("Arial", 10),
                 ).pack(side="left", pady=10)

    def _build_left_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  File PDF cho xu ly  ",
                              bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_PRIMARY,
                              padx=8, pady=6)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        # Folder row
        folder_row = tk.Frame(frame, bg=COLOR_BG)
        folder_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        folder_row.columnconfigure(0, weight=1)

        tk.Label(folder_row, text="Thu muc tim kiem:", bg=COLOR_BG,
                 font=("Arial", 8, "bold"), fg="#555").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar(value=DEFAULT_PDF_FOLDER)
        tk.Entry(folder_row, textvariable=self.folder_var,
                 font=("Arial", 8), width=38).grid(row=1, column=0, sticky="ew", padx=(0, 4))
        tk.Button(folder_row, text="...", command=self._browse_folder,
                  bg=COLOR_BTN_BLUE, fg="white", font=("Arial", 8, "bold"),
                  relief="flat", cursor="hand2", padx=6).grid(row=1, column=1)

        # Listbox
        list_frame = tk.Frame(frame, bg=COLOR_BG)
        list_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(
            list_frame, font=("Courier", 9), selectmode=tk.EXTENDED,
            bg="white", fg="#1a1a1a",
            selectbackground=COLOR_PRIMARY, selectforeground="white",
            activestyle="none", bd=1, relief="solid", highlightthickness=0)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=sb.set)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # Buttons
        btn_row = tk.Frame(frame, bg=COLOR_BG)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        tk.Button(btn_row, text="+ Chon file", command=self._add_files_dialog,
                  bg=COLOR_BTN_GREEN, fg="white", font=("Arial", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=4).pack(side="left", padx=(0, 4))
        tk.Button(btn_row, text="Xoa chon", command=self._remove_selected,
                  bg=COLOR_BTN_RED, fg="white", font=("Arial", 9),
                  relief="flat", cursor="hand2", padx=8, pady=4).pack(side="left", padx=(0, 4))
        tk.Button(btn_row, text="Xoa tat ca", command=self._clear_queue,
                  bg=COLOR_BTN_GRAY, fg="white", font=("Arial", 9),
                  relief="flat", cursor="hand2", padx=8, pady=4).pack(side="left")

        self._queue_count_var = tk.StringVar(value="0 file")
        tk.Label(btn_row, textvariable=self._queue_count_var,
                 bg=COLOR_BG, fg=COLOR_PRIMARY, font=("Arial", 9, "bold")).pack(side="right")

    def _build_right_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  Tim file theo ten / ASN  ",
                              bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_PRIMARY,
                              padx=8, pady=6)
        frame.grid(row=0, column=1, sticky="nsew")
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        tk.Label(frame, text="Nhap moi ten / ASN mot dong\n(chinh xac hoac mot phan):",
                 bg=COLOR_BG, font=("Arial", 9), fg="#444", justify="left",
                 ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        text_frame = tk.Frame(frame, bg=COLOR_BG)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.name_input = tk.Text(
            text_frame, font=("Courier", 9), width=28, height=10,
            bg="white", fg="#AAAAAA",
            bd=1, relief="solid", wrap="none",
            highlightthickness=1, highlightcolor=COLOR_PRIMARY)
        sb2 = ttk.Scrollbar(text_frame, orient="vertical", command=self.name_input.yview)
        self.name_input.configure(yscrollcommand=sb2.set)
        self.name_input.grid(row=0, column=0, sticky="nsew")
        sb2.grid(row=0, column=1, sticky="ns")

        self.name_input.insert("1.0", "VD:\nASN1034657\nASN1034658\n1034659")
        self._placeholder_active = True
        self.name_input.bind("<FocusIn>",  self._clear_placeholder)
        self.name_input.bind("<FocusOut>", self._restore_placeholder)

        search_row = tk.Frame(frame, bg=COLOR_BG)
        search_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        tk.Button(search_row, text="Tim & them vao queue",
                  command=self._search_and_add,
                  bg=COLOR_BTN_BLUE, fg="white", font=("Arial", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=5,
                  ).pack(fill="x")

        self._search_result_var = tk.StringVar(value="")
        tk.Label(frame, textvariable=self._search_result_var,
                 bg=COLOR_BG, fg="#1F7A3E", font=("Arial", 8),
                 wraplength=220, justify="left",
                 ).grid(row=3, column=0, sticky="w", pady=(4, 0))

    def _build_result_table(self):
        frame = tk.LabelFrame(self, text="  Ket qua da xu ly  ",
                              bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_PRIMARY)
        frame.pack(fill="both", expand=True, padx=12, pady=(6, 4))

        col_conf = {
            "File":              {"width": 240, "anchor": "w"},
            "Invoice Date":      {"width": 110, "anchor": "center"},
            "Invoice Number":    {"width": 140, "anchor": "center"},
            "Shipment ID / ASN": {"width": 140, "anchor": "center"},
        }
        display_cols = OUTPUT_COLUMNS + ["Status"]
        self.tree = ttk.Treeview(frame, columns=display_cols, show="headings", height=8)
        for col in display_cols:
            cfg = col_conf.get(col, {"width": 100, "anchor": "center"})
            self.tree.heading(col, text=col)
            self.tree.column(col, width=cfg["width"], anchor=cfg["anchor"], stretch=True)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"))
        style.configure("Treeview", font=("Arial", 9), rowheight=22)

        sb_y = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        sb_x = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.tag_configure("ok",    background="#FFFFFF")
        self.tree.tag_configure("even",  background=COLOR_ROW_EVEN)
        self.tree.tag_configure("error", background=COLOR_ROW_ERROR)

    def _build_action_bar(self):
        bar = tk.Frame(self, bg=COLOR_BG)
        bar.pack(fill="x", padx=12, pady=(0, 4))

        self.process_btn = tk.Button(
            bar, text="Xu ly file trong queue",
            command=self._start_process,
            bg=COLOR_BTN_GREEN, fg="white", font=("Arial", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=7, state="disabled")
        self.process_btn.pack(side="left", padx=(0, 8))

        self.export_btn = tk.Button(
            bar, text="Xuat Excel",
            command=self._export_excel,
            bg=COLOR_BTN_ORANGE, fg="white", font=("Arial", 11, "bold"),
            relief="flat", cursor="hand2", padx=18, pady=7, state="disabled")
        self.export_btn.pack(side="left", padx=(0, 8))

        out_frame = tk.Frame(bar, bg=COLOR_BG)
        out_frame.pack(side="right")
        tk.Label(out_frame, text="Luu tai:", bg=COLOR_BG,
                 font=("Arial", 8), fg="#555").pack(side="left", padx=(0, 4))
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT_FOLDER)
        tk.Entry(out_frame, textvariable=self.output_var,
                 font=("Arial", 8), width=28).pack(side="left", padx=(0, 4))
        tk.Button(out_frame, text="...", command=self._browse_output,
                  bg=COLOR_BTN_BLUE, fg="white", font=("Arial", 8, "bold"),
                  relief="flat", cursor="hand2", padx=6).pack(side="left")

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="San sang.")
        tk.Label(self, textvariable=self.status_var,
                 bg=COLOR_PRIMARY, fg="white",
                 font=("Arial", 9), anchor="w", padx=12,
                 ).pack(fill="x", side="bottom")

    # ── Placeholder ──────────────────────────────────────────

    _PLACEHOLDER = "VD:\nASN1034657\nASN1034658\n1034659"

    def _clear_placeholder(self, _=None):
        if self._placeholder_active:
            self.name_input.delete("1.0", "end")
            self.name_input.config(fg="#1a1a1a")
            self._placeholder_active = False

    def _restore_placeholder(self, _=None):
        if not self.name_input.get("1.0", "end").strip():
            self.name_input.insert("1.0", self._PLACEHOLDER)
            self.name_input.config(fg="#AAAAAA")
            self._placeholder_active = True

    # ── Queue management ─────────────────────────────────────

    def _refresh_listbox(self):
        self.file_listbox.delete(0, "end")
        for path in self._queued_paths:
            self.file_listbox.insert("end", f"  [PDF]  {os.path.basename(path)}")
        count = len(self._queued_paths)
        self._queue_count_var.set(f"{count} file")
        self.process_btn.config(state="normal" if count > 0 else "disabled")

    def _add_paths(self, new_paths: list[str]) -> tuple[int, int, int]:
        added = skipped = replaced = 0
        existing = {os.path.basename(p): i for i, p in enumerate(self._queued_paths)}

        for path in new_paths:
            name = os.path.basename(path)
            if name in existing:
                answer = messagebox.askyesnocancel(
                    "File trung ten",
                    f"File da co trong danh sach:\n{name}\n\n"
                    "Yes = Thay the  |  No = Bo qua  |  Cancel = Dung them",
                )
                if answer is None:
                    break
                elif answer:
                    self._queued_paths[existing[name]] = path
                    replaced += 1
                else:
                    skipped += 1
            else:
                self._queued_paths.append(path)
                existing[name] = len(self._queued_paths) - 1
                added += 1

        return added, skipped, replaced

    def _add_files_dialog(self):
        paths = filedialog.askopenfilenames(
            title="Chon file PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self.folder_var.get(),
        )
        if not paths:
            return
        added, skipped, replaced = self._add_paths(list(paths))
        self._refresh_listbox()
        msg = f"Them {added} file"
        if replaced: msg += f", thay the {replaced}"
        if skipped:  msg += f", bo qua {skipped} trung"
        self._set_status(msg)

    def _remove_selected(self):
        for idx in reversed(self.file_listbox.curselection()):
            del self._queued_paths[idx]
        self._refresh_listbox()

    def _clear_queue(self):
        self._queued_paths.clear()
        self._refresh_listbox()
        self._set_status("Da xoa toan bo file trong queue.")

    # ── Search by name ───────────────────────────────────────

    def _search_and_add(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Loi", f"Thu muc khong ton tai:\n{folder}")
            return

        raw = "" if self._placeholder_active else self.name_input.get("1.0", "end")
        names = [n.strip() for n in raw.splitlines() if n.strip()]
        if not names:
            messagebox.showwarning("Canh bao", "Vui long nhap it nhat mot ten file!")
            return

        results = search_pdf_by_names(folder, names)
        found_paths: list[str] = []
        not_found:   list[str] = []

        for name, matches in results.items():
            if len(matches) == 0:
                not_found.append(name)
            elif len(matches) == 1:
                found_paths.append(matches[0])
            else:
                chosen = self._pick_from_multiple(name, matches)
                if chosen:
                    found_paths.extend(chosen)

        added, skipped, replaced = self._add_paths(found_paths)
        self._refresh_listbox()

        lines = []
        if added:     lines.append(f"Them {added}")
        if replaced:  lines.append(f"Thay the {replaced}")
        if skipped:   lines.append(f"Bo qua {skipped} trung")
        if not_found: lines.append(f"Khong tim thay: {', '.join(not_found)}")
        summary = "  |  ".join(lines) if lines else "Khong co file nao duoc them."
        self._search_result_var.set(summary)
        self._set_status("Tim kiem xong: " + summary)

    def _pick_from_multiple(self, keyword: str, matches: list[str]) -> list[str]:
        dialog = tk.Toplevel(self)
        dialog.title(f"Nhieu ket qua: {keyword}")
        dialog.geometry("500x280")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.configure(bg=COLOR_BG)

        tk.Label(dialog,
                 text=f'Tu khoa "{keyword}" khop {len(matches)} file.\nChon file can them:',
                 bg=COLOR_BG, font=("Arial", 10), justify="left",
                 ).pack(padx=14, pady=(12, 6), anchor="w")

        lb_frame = tk.Frame(dialog, bg=COLOR_BG)
        lb_frame.pack(fill="both", expand=True, padx=14)
        lb = tk.Listbox(lb_frame, selectmode=tk.MULTIPLE, font=("Courier", 9),
                        bg="white", selectbackground=COLOR_PRIMARY,
                        selectforeground="white", height=8)
        sb = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        for p in matches:
            lb.insert("end", f"  [PDF]  {os.path.basename(p)}")
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        chosen: list[str] = []

        def confirm():
            for i in lb.curselection():
                chosen.append(matches[i])
            dialog.destroy()

        btn_row = tk.Frame(dialog, bg=COLOR_BG)
        btn_row.pack(pady=10)
        tk.Button(btn_row, text="Them da chon", command=confirm,
                  bg=COLOR_BTN_GREEN, fg="white", font=("Arial", 9, "bold"),
                  relief="flat", cursor="hand2", padx=12, pady=5).pack(side="left", padx=6)
        tk.Button(btn_row, text="Bo qua", command=dialog.destroy,
                  bg=COLOR_BTN_GRAY, fg="white", font=("Arial", 9),
                  relief="flat", cursor="hand2", padx=10, pady=5).pack(side="left")

        dialog.wait_window()
        return chosen

    # ── Process queue ────────────────────────────────────────

    def _start_process(self):
        if not self._queued_paths:
            return
        self.process_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        threading.Thread(target=self._process_worker,
                         args=(list(self._queued_paths),), daemon=True).start()

    def _process_worker(self, paths: list[str]):
        total = len(paths)
        self.after(0, lambda: self._set_status(f"Dang xu ly {total} file..."))

        new_results: list[dict] = []
        for idx, path in enumerate(paths, 1):
            name = os.path.basename(path)
            self.after(0, lambda i=idx, t=total, n=name:
                       self._set_status(f"[{i}/{t}]  {n}"))
            record = extract_from_pdf(path)
            if record:
                new_results.append(record)
                row_idx = len(self._results) + len(new_results)
                self.after(0, lambda r=record, i=row_idx: self._append_row(r, i))

        self._results.extend(new_results)
        self.after(0, self._clear_queue)

        ok  = sum(1 for r in new_results if "error" not in r)
        err = len(new_results) - ok
        skip = total - len(new_results)
        msg = f"Xong {ok} file"
        if err:  msg += f"  | {err} loi"
        if skip: msg += f"  | {skip} khong co trang CI"
        self.after(0, lambda: self._set_status(msg))
        if self._results:
            self.after(0, lambda: self.export_btn.config(state="normal"))

    def _append_row(self, record: dict, idx: int):
        is_err = "error" in record
        tag = "error" if is_err else ("even" if idx % 2 == 0 else "ok")
        status = "ERR" if is_err else "OK"
        values = (
            record.get("File", ""),
            record.get("error", "") if is_err else record.get("Invoice Date", ""),
            record.get("Invoice Number", ""),
            record.get("Shipment ID / ASN", ""),
            status,
        )
        self.tree.insert("", "end", values=values, tags=(tag,))

    # ── Export ───────────────────────────────────────────────

    def _export_excel(self):
        clean = [r for r in self._results if "error" not in r]
        if not clean:
            messagebox.showwarning("Canh bao", "Khong co du lieu hop le de xuat!")
            return
        output_dir = self.output_var.get().strip() or DEFAULT_OUTPUT_FOLDER
        os.makedirs(output_dir, exist_ok=True)
        export_data = [{k: v for k, v in r.items() if not k.startswith("_")} for r in clean]
        csv_path  = os.path.join(output_dir, OUTPUT_CSV_NAME)
        xlsx_path = os.path.join(output_dir, OUTPUT_XLSX_NAME)
        try:
            save_csv(export_data,  csv_path)
            save_excel(export_data, xlsx_path)
            messagebox.showinfo("Xuat thanh cong",
                f"Da luu {len(export_data)} dong:\n\nCSV:   {csv_path}\nExcel: {xlsx_path}")
            self._set_status(f"Xuat xong {len(export_data)} dong -> {xlsx_path}")
        except Exception as exc:
            messagebox.showerror("Loi xuat file", str(exc))

    # ── Misc ─────────────────────────────────────────────────

    def _browse_folder(self):
        p = filedialog.askdirectory(title="Chon thu muc PDF", initialdir=self.folder_var.get())
        if p:
            self.folder_var.set(p)

    def _browse_output(self):
        p = filedialog.askdirectory(title="Chon thu muc luu ket qua")
        if p:
            self.output_var.set(p)

    def _set_status(self, msg: str):
        self.status_var.set(msg)