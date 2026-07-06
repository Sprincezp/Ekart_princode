#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VLOOKUP Pro - Advanced Excel Matching Tool
Match by any column and fill data automatically
One-click, fast, zero errors
"""

import os
import sys
import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from datetime import datetime
import traceback

try:
    import pandas as pd
    import openpyxl
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Missing Libraries",
        "Required libraries not found.\n\n"
        "Run this command in terminal:\n"
        "pip install pandas openpyxl")
    root.destroy()
    sys.exit(1)


class VLookupPro:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VLOOKUP Pro - Excel Matcher")
        self.root.geometry("1000x700+100+50")
        self.root.configure(bg="#0f1a15")
        self.root.minsize(900, 650)

        try:
            self.root.iconbitmap(default="")
        except:
            pass

        # ---------- state ----------
        self.file1_path = ""
        self.file2_path = ""
        self.df1 = None
        self.df2 = None
        self.result_df = None
        self._processing = False

        # ---------- colors ----------
        self.C = {
            "bg": "#0f1a15",
            "bg2": "#142420",
            "bg3": "#1a3028",
            "fg": "#e8f0eb",
            "fg2": "#8fa899",
            "fg3": "#5a7467",
            "accent": "#10b981",
            "accent2": "#34d399",
            "err": "#ef4444",
            "warn": "#f59e0b",
            "border": "#234038",
        }

        self._style_ttk()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # =================================================================
    # STYLING
    # =================================================================
    def _style_ttk(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".",
            background=self.C["bg"],
            foreground=self.C["fg"],
            fieldbackground=self.C["bg2"],
            troughcolor=self.C["bg3"],
            selectbackground=self.C["accent"],
            selectforeground="#000",
            font=("Segoe UI", 10),
        )
        style.configure("TLabel", background=self.C["bg"], foreground=self.C["fg"])
        style.configure("TFrame", background=self.C["bg"])
        style.configure("TButton",
            background=self.C["accent"], foreground="#000",
            borderwidth=0, focusthroughcolor="",
            font=("Segoe UI", 10, "bold"), padding=(16, 8))
        style.map("TButton",
            background=[("active", self.C["accent2"]), ("disabled", self.C["bg3"])],
            foreground=[("disabled", self.C["fg3"])])
        style.configure("Accent.TButton",
            background=self.C["accent"], foreground="#000",
            font=("Segoe UI", 12, "bold"), padding=(24, 12))
        style.map("Accent.TButton",
            background=[("active", self.C["accent2"]), ("disabled", self.C["bg3"])])
        style.configure("Outline.TButton",
            background=self.C["bg2"], foreground=self.C["fg"],
            borderwidth=1, relief="solid", font=("Segoe UI", 10))
        style.map("Outline.TButton",
            background=[("active", self.C["bg3"])])
        style.configure("Status.TLabel",
            background=self.C["bg"], foreground=self.C["fg2"], font=("Segoe UI", 9))
        style.configure("Title.TLabel",
            background=self.C["bg"], foreground=self.C["fg"],
            font=("Segoe UI", 16, "bold"))
        style.configure("Heading.TLabel",
            background=self.C["bg"], foreground=self.C["accent"],
            font=("Segoe UI", 11, "bold"))
        style.configure("Stats.TLabel",
            background=self.C["bg2"], foreground=self.C["accent"],
            font=("Segoe UI", 20, "bold"))
        style.configure("StatsSub.TLabel",
            background=self.C["bg2"], foreground=self.C["fg2"],
            font=("Segoe UI", 9))
        style.configure("Red.TLabel",
            background=self.C["bg2"], foreground=self.C["err"],
            font=("Segoe UI", 20, "bold"))
        style.configure("Warn.TLabel",
            background=self.C["bg2"], foreground=self.C["warn"],
            font=("Segoe UI", 20, "bold"))

        style.configure("Treeview",
            background=self.C["bg2"], foreground=self.C["fg"],
            fieldbackground=self.C["bg2"], borderwidth=0,
            rowheight=28, font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
            background=self.C["bg3"], foreground=self.C["fg"],
            font=("Segoe UI", 9, "bold"), borderwidth=0)
        style.map("Treeview.Heading",
            background=[("active", self.C["border"])])
        style.map("Treeview",
            background=[("selected", self.C["accent"])],
            foreground=[("selected", "#000")])

        style.configure("Vertical.TScrollbar",
            background=self.C["bg3"], troughcolor=self.C["bg"],
            borderwidth=0, arrowcolor=self.C["fg2"])
        style.configure("Horizontal.TProgressbar",
            background=self.C["accent"], troughcolor=self.C["bg3"],
            borderwidth=0, thickness=20)

    def _make_label(self, parent, **kw):
        return tk.Label(parent, bg=self.C["bg"], fg=self.C["fg"],
                       font=("Segoe UI", 10), **kw)

    def _make_heading(self, parent, **kw):
        return tk.Label(parent, bg=self.C["bg"], fg=self.C["accent"],
                       font=("Segoe UI", 12, "bold"), **kw)

    def _make_frame(self, parent, bg=None, **kw):
        return tk.Frame(parent, bg=bg or self.C["bg"], **kw)

    def _card(self, parent, **kw):
        f = tk.Frame(parent, bg=self.C["bg2"], bd=0, highlightthickness=0, **kw)
        f.config(highlightbackground=self.C["border"], highlightcolor=self.C["border"])
        return f

    # =================================================================
    # UI BUILD
    # =================================================================
    def _build_ui(self):
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # ---------- header ----------
        hdr = self._make_frame(self.root)
        hdr.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 4))

        ico = tk.Label(hdr, text="⚡", bg=self.C["bg"], fg=self.C["accent"],
                       font=("Segoe UI", 24))
        ico.pack(side="left", padx=(0, 8))

        tt = tk.Label(hdr, text="VLOOKUP Pro",
                      bg=self.C["bg"], fg=self.C["fg"],
                      font=("Segoe UI", 20, "bold"))
        tt.pack(side="left")

        sub = tk.Label(hdr, text="Excel Data Matcher",
                       bg=self.C["bg"], fg=self.C["fg2"],
                       font=("Segoe UI", 10))
        sub.pack(side="left", padx=(12, 0), pady=(6, 0))

        # ---------- scrollable main container ----------
        canvas = tk.Canvas(self.root, bg=self.C["bg"], highlightthickness=0, bd=0)
        canvas.grid(row=1, column=0, sticky="nsew", padx=16, pady=(4, 12))

        v_scroll = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        v_scroll.grid(row=1, column=1, sticky="ns", pady=(4, 12))
        canvas.configure(yscrollcommand=v_scroll.set)

        main = tk.Frame(canvas, bg=self.C["bg"])
        canvas_window = canvas.create_window((0, 0), window=main, anchor="nw")

        def _configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _configure_canvas)

        def _configure_main(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        main.bind("<Configure>", _configure_main)

        def _on_mousewheel(event):
            canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        # -- top row: left panel + right panel --
        top_row = self._make_frame(main)
        top_row.pack(fill="both", expand=True)

        left = self._make_frame(top_row, width=480)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        left.pack_propagate(False)

        self._build_file_section(left, 1, "Source File (jahan se data lena hai)",
                                "📂", self.C["accent"])
        self._build_file_section(left, 2, "Target File (jahan bharna hai)",
                                "🎯", self.C["warn"])

        right = self._make_frame(top_row)
        right.pack(side="right", fill="both", expand=True, padx=(8, 0))

        self._build_settings(right)
        self._build_preview(right)
        self._build_stats(right)
        self._build_actions(right)

    # ---------- file section ----------
    def _build_file_section(self, parent, n, title, icon, color):
        card = self._card(parent)
        card.pack(fill="x", pady=(0, 8), padx=0)

        inner = tk.Frame(card, bg=self.C["bg2"])
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        row1 = tk.Frame(inner, bg=self.C["bg2"])
        row1.pack(fill="x")

        lbl_icon = tk.Label(row1, text=icon, bg=self.C["bg2"], fg=color,
                           font=("Segoe UI", 14))
        lbl_icon.pack(side="left", padx=(0, 6))

        lbl_title = tk.Label(row1, text=title, bg=self.C["bg2"], fg=self.C["fg"],
                            font=("Segoe UI", 11, "bold"))
        lbl_title.pack(side="left")

        # file path label
        path_frame = tk.Frame(inner, bg=self.C["bg2"])
        path_frame.pack(fill="x", pady=(8, 0))

        path_lbl = tk.Label(path_frame, text="No file selected",
                           bg=self.C["bg2"], fg=self.C["fg3"],
                           font=("Segoe UI", 9), anchor="w")
        path_lbl.pack(side="left", fill="x", expand=True)

        btn_browse = ttk.Button(path_frame, text="Browse",
                               style="TButton",
                               command=lambda n=n: self._browse_file(n))
        btn_browse.pack(side="right", padx=(8, 0))

        # column mapping row
        map_frame = tk.Frame(inner, bg=self.C["bg2"])
        map_frame.pack(fill="x", pady=(10, 0))

        # Match column
        tk.Label(map_frame, text="Match Col:", bg=self.C["bg2"],
                fg=self.C["fg2"], font=("Segoe UI", 9)).pack(side="left")
        match_var = tk.StringVar()
        match_cb = ttk.Combobox(map_frame, textvariable=match_var,
                                state="readonly", width=22,
                                font=("Segoe UI", 9))
        match_cb.pack(side="left", padx=(4, 12))

        # Data column
        tk.Label(map_frame, text="Data Col:", bg=self.C["bg2"],
                fg=self.C["fg2"], font=("Segoe UI", 9)).pack(side="left")
        data_var = tk.StringVar()
        data_cb = ttk.Combobox(map_frame, textvariable=data_var,
                               state="readonly", width=22,
                               font=("Segoe UI", 9))
        data_cb.pack(side="left", padx=(4, 0))

        # store refs
        setattr(self, f"_path_lbl_{n}", path_lbl)
        setattr(self, f"_match_var_{n}", match_var)
        setattr(self, f"_match_cb_{n}", match_cb)
        setattr(self, f"_data_var_{n}", data_var)
        setattr(self, f"_data_cb_{n}", data_cb)
        setattr(self, f"_browse_btn_{n}", btn_browse)

    # ---------- settings ----------
    def _build_settings(self, parent):
        card = self._card(parent)
        card.pack(fill="x", pady=(0, 6))

        inner = tk.Frame(card, bg=self.C["bg2"])
        inner.pack(fill="x", padx=12, pady=8)

        self._make_heading(inner, text="⚙ Settings").pack(anchor="w")

        row = tk.Frame(inner, bg=self.C["bg2"])
        row.pack(fill="x", pady=(6, 0))

        tk.Label(row, text="Match Type:", bg=self.C["bg2"],
                fg=self.C["fg2"], font=("Segoe UI", 9)).pack(side="left")
        self._match_type = ttk.Combobox(row, values=[
            "Exact Match", "Contains", "Starts With", "Ends With"],
            state="readonly", width=16, font=("Segoe UI", 9))
        self._match_type.current(0)
        self._match_type.pack(side="left", padx=(6, 16))

        self._case_var = tk.IntVar()
        cb = tk.Checkbutton(row, text="Case Sensitive",
                           variable=self._case_var,
                           bg=self.C["bg2"], fg=self.C["fg2"],
                           selectcolor=self.C["bg2"],
                           activebackground=self.C["bg2"],
                           activeforeground=self.C["fg"],
                           font=("Segoe UI", 9))
        cb.pack(side="left", padx=(0, 16))

        tk.Label(row, text="No Match:", bg=self.C["bg2"],
                fg=self.C["fg2"], font=("Segoe UI", 9)).pack(side="left")
        self._nomatch_var = ttk.Combobox(row, values=[
            "Leave Empty", "Fill N/A", "Custom Value"],
            state="readonly", width=14, font=("Segoe UI", 9))
        self._nomatch_var.current(0)
        self._nomatch_var.pack(side="left", padx=(6, 0))

        self._custom_val = tk.Entry(row, bg=self.C["bg2"], fg=self.C["fg"],
                                    insertbackground=self.C["fg"],
                                    font=("Segoe UI", 9), width=12,
                                    relief="flat", bd=0)
        self._custom_val.pack(side="left", padx=(4, 0))
        self._custom_val.insert(0, "-")
        self._custom_val.config(state="disabled")

        def on_nomatch(*_):
            self._custom_val.config(
                state="normal" if self._nomatch_var.get() == "Custom Value" else "disabled")
        self._nomatch_var.bind("<<ComboboxSelected>>", on_nomatch)

    # ---------- preview ----------
    def _build_preview(self, parent):
        card = self._card(parent)
        card.pack(fill="both", expand=True, pady=(0, 6))

        inner = tk.Frame(card, bg=self.C["bg2"])
        inner.pack(fill="both", expand=True, padx=12, pady=8)

        hdr = tk.Frame(inner, bg=self.C["bg2"])
        hdr.pack(fill="x")
        self._make_heading(hdr, text="👁 Preview").pack(side="left")

        self._preview_info = tk.Label(hdr, text="",
            bg=self.C["bg2"], fg=self.C["fg3"], font=("Segoe UI", 9))
        self._preview_info.pack(side="left", padx=(12, 0))

        # treeview container
        tv_frame = tk.Frame(inner, bg=self.C["bg2"])
        tv_frame.pack(fill="both", expand=True, pady=(6, 0))

        vsb = ttk.Scrollbar(tv_frame, orient="vertical")
        hsb = ttk.Scrollbar(tv_frame, orient="horizontal")

        self._tree = ttk.Treeview(tv_frame,
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            height=8, selectmode="none")
        vsb.config(command=self._tree.yview)
        hsb.config(command=self._tree.xview)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self._tree.pack(side="left", fill="both", expand=True)

    # ---------- stats ----------
    def _build_stats(self, parent):
        card = self._card(parent)
        card.pack(fill="x", pady=(0, 6))

        inner = tk.Frame(card, bg=self.C["bg2"])
        inner.pack(fill="x", padx=12, pady=8)

        self._make_heading(inner, text="📊 Results").pack(anchor="w")

        grid = tk.Frame(inner, bg=self.C["bg2"])
        grid.pack(fill="x", pady=(6, 0))

        stats = [
            ("sv_total", "Total", self.C["fg"]),
            ("sv_matched", "Matched", self.C["accent"]),
            ("sv_unmatched", "Unmatched", self.C["err"]),
            ("sv_multi", "Multiple", self.C["warn"]),
        ]
        for key, lbl, clr in stats:
            box = tk.Frame(grid, bg=self.C["bg"], bd=0,
                          highlightbackground=self.C["border"],
                          highlightthickness=1)
            box.pack(side="left", fill="x", expand=True, padx=2)

            val_lbl = tk.Label(box, text="0", bg=self.C["bg"], fg=clr,
                              font=("Segoe UI", 18, "bold"))
            val_lbl.pack(pady=(6, 0))

            tk.Label(box, text=lbl, bg=self.C["bg"], fg=self.C["fg3"],
                    font=("Segoe UI", 9)).pack(pady=(0, 6))

            setattr(self, f"_{key}", val_lbl)

        self._status_lbl = tk.Label(inner, text="Ready. Load both files to begin.",
            bg=self.C["bg2"], fg=self.C["fg3"], font=("Segoe UI", 9), anchor="w")
        self._status_lbl.pack(fill="x", pady=(6, 0))

    # ---------- actions ----------
    def _build_actions(self, parent):
        card = self._card(parent)
        card.pack(fill="x")

        inner = tk.Frame(card, bg=self.C["bg2"])
        inner.pack(fill="x", padx=12, pady=10)

        # progress bar
        self._progress = ttk.Progressbar(inner, style="Horizontal.TProgressbar",
                                         mode="determinate", value=0)
        self._progress.pack(fill="x", pady=(0, 10))

        btn_row = tk.Frame(inner, bg=self.C["bg2"])
        btn_row.pack(fill="x")

        self._process_btn = ttk.Button(btn_row, text="🚀 MATCH & FILL",
                                      style="Accent.TButton",
                                      command=self._start_process)
        self._process_btn.pack(side="left", padx=(0, 8))

        self._save_btn = ttk.Button(btn_row, text="💾 Save Result",
                                   style="Outline.TButton",
                                   command=self._save_result, state="disabled")
        self._save_btn.pack(side="left", padx=(0, 8))

        self._reset_btn = ttk.Button(btn_row, text="🔄 Reset",
                                    style="Outline.TButton",
                                    command=self._reset_all)
        self._reset_btn.pack(side="left")

    # =================================================================
    # FILE OPERATIONS
    # =================================================================
    def _browse_file(self, n):
        path = filedialog.askopenfilename(
            title=f"Select Excel File {n}",
            filetypes=[("All supported", "*.xlsx *.xls *.xlsm *.csv"),
                      ("Excel files", "*.xlsx *.xls *.xlsm"),
                      ("CSV files", "*.csv"),
                      ("All files", "*.*")])
        if not path:
            return
        self._load_file(n, path)

    def _load_file(self, n, path):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".csv":
                try:
                    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="cp1252")
            elif ext in (".xls", ".xlsx", ".xlsm"):
                df = pd.read_excel(path, dtype=str, keep_default_na=False, engine="openpyxl")
            else:
                raise ValueError(f"Unsupported format: {ext}")

            df.columns = df.columns.str.strip()
            df = df.loc[:, ~df.columns.duplicated()]

            if n == 1:
                self.df1 = df
                self.file1_path = path
            else:
                self.df2 = df
                self.file2_path = path

            # update path label
            lbl = getattr(self, f"_path_lbl_{n}")
            fname = os.path.basename(path)
            lbl.config(text=f"{fname}  ({len(df)} rows, {len(df.columns)} cols)")

            # populate comboboxes
            cols = list(df.columns)
            match_cb = getattr(self, f"_match_cb_{n}")
            data_cb = getattr(self, f"_data_cb_{n}")

            match_cb["values"] = cols
            data_cb["values"] = cols

            # auto-detect columns
            match_col, data_col = self._auto_detect_cols(n, cols)

            match_var = getattr(self, f"_match_var_{n}")
            match_var.set(match_col or (cols[0] if cols else ""))

            data_var = getattr(self, f"_data_var_{n}")
            data_var.set(data_col or "")

            # update preview
            self._update_preview()
            self._update_status(f"File {n} loaded: {fname}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            self._update_status(f"Error loading file {n}")

    def _auto_detect_cols(self, n, cols):
        """Auto-detect Master Waybill → match col, Image → data col."""
        PRIORITY_MATCH = [
            "master waybill", "waybill", "waybill no", "waybill number",
            "awb", "awb no", "awb number", "order id", "order no",
            "id", "invoice", "invoice no",
        ]
        PRIORITY_DATA = [
            "image", "images", "photo", "photos", "picture",
            "image url", "image link", "img", "attachment",
            "file", "document", "scan",
        ]

        lower = [c.strip().lower() for c in cols]

        # Match column
        match_col = None
        for p in PRIORITY_MATCH:
            for i, lc in enumerate(lower):
                if lc == p or lc.startswith(p) or lc.replace(" ", "").replace("_", "").replace("-", "") == p.replace(" ", ""):
                    match_col = cols[i]
                    break
            if match_col:
                break
        if not match_col and n == 2 and self.df1 is not None:
            # If file 2, try to find a column that matches file 1's match col
            match_var_1 = getattr(self, "_match_var_1", None)
            if match_var_1:
                m1 = match_var_1.get()
                if m1 and m1 in cols:
                    match_col = m1
        if not match_col and len(cols) > 0:
            match_col = cols[0]

        # Data column (prefer "Image" in file 2)
        data_col = None
        for p in PRIORITY_DATA:
            for i, lc in enumerate(lower):
                if lc == p or lc.startswith(p):
                    data_col = cols[i]
                    break
            if data_col:
                break

        return match_col, data_col

    # =================================================================
    # PREVIEW
    # =================================================================
    def _update_preview(self):
        for item in self._tree.get_children():
            self._tree.delete(item)

        # decide which DF to show
        df = None
        lbl = ""
        if self.df2 is not None:
            df = self.df2
            lbl = f"Target: {os.path.basename(self.file2_path)}"
        elif self.df1 is not None:
            df = self.df1
            lbl = f"Source: {os.path.basename(self.file1_path)}"

        if df is None:
            self._preview_info.config(text="No data to preview")
            return

        self._preview_info.config(text=lbl)

        cols = list(df.columns)
        self._tree["columns"] = cols
        self._tree["displaycolumns"] = cols[:20]  # limit display

        self._tree.heading("#0", text="#")
        self._tree.column("#0", width=40, minwidth=40, anchor="center")
        for c in cols[:20]:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=100, minwidth=60)

        for i, (_, row) in enumerate(df.head(10).iterrows()):
            vals = [str(row.get(c, ""))[:60] for c in cols[:20]]
            self._tree.insert("", "end", text=str(i + 1), values=vals)

    def _update_status(self, msg):
        self._status_lbl.config(text=msg)
        self.root.update_idletasks()

    def _set_stats(self, total=0, matched=0, unmatched=0, multi=0):
        self._sv_total.config(text=str(total))
        self._sv_matched.config(text=str(matched))
        self._sv_unmatched.config(text=str(unmatched))
        self._sv_multi.config(text=str(multi))

    # =================================================================
    # PROCESSING
    # =================================================================
    def _start_process(self):
        if self._processing:
            return

        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("Missing Files", "Please load both Excel files first.")
            return

        match1 = getattr(self, "_match_var_1").get()
        match2 = getattr(self, "_match_var_2").get()
        data2 = getattr(self, "_data_var_2").get()

        if not match1 or not match2:
            messagebox.showwarning("Missing Mapping",
                "Select Match Column for both files.")
            return

        if not data2:
            messagebox.showwarning("Missing Mapping",
                "Select Data Column in Target file (the column to fill).")
            return

        self._processing = True
        self._process_btn.config(state="disabled", text="⏳ Processing...")
        self._progress["value"] = 0

        # Run in a thread so UI stays responsive
        t = threading.Thread(target=self._process, daemon=True)
        t.start()

    def _process(self):
        try:
            match1 = getattr(self, "_match_var_1").get()
            match2 = getattr(self, "_match_var_2").get()
            data2 = getattr(self, "_data_var_2").get()

            match_type = self._match_type.get()
            case_sensitive = self._case_var.get() == 1
            nomatch = self._nomatch_var.get()
            custom_val = self._custom_val.get()

            self._update_status("Building match index...")
            self._progress["value"] = 5

            # Prepare source (file 1) - it has the data to copy FROM
            # We need to find which column in file 1 has the data to fill
            # Auto-detect: if file 1 has a column with same name as data2, use it
            data1 = None
            if data2 in self.df1.columns:
                data1 = data2
            else:
                # Try to find a similar column in source
                lower_d2 = data2.strip().lower()
                for c in self.df1.columns:
                    if c.strip().lower() == lower_d2:
                        data1 = c
                        break
                if not data1:
                    # Use the first non-match column in source
                    for c in self.df1.columns:
                        if c != match1:
                            data1 = c
                            break

            # Build lookup dictionary from file 2 (target - has the actual data)
            self._update_status("Indexing target data...")
            lookup = {}
            for i, (_, row) in enumerate(self.df2.iterrows()):
                key = str(row.get(match2, "")).strip()
                if not case_sensitive:
                    key = key.lower()
                val = row.get(data2, "")
                if key:
                    if key not in lookup:
                        lookup[key] = []
                    lookup[key].append(val)

                if i % 500 == 0 and i > 0:
                    self._progress["value"] = min(20, 5 + int(15 * i / len(self.df2)))
                    self._update_status(f"Indexing... {i}/{len(self.df2)} rows")

            self._progress["value"] = 25
            self._update_status("Matching & filling data...")

            # Process file 1 (source) - fill in the data column
            result = self.df1.copy()
            if data1 and data1 in result.columns:
                # If data1 column already exists, we'll update it
                pass
            else:
                # Create the new column
                if data1:
                    result[data1] = ""

            matched = 0
            unmatched = 0
            multi = 0
            total = len(result)

            for i, (idx, row) in enumerate(result.iterrows()):
                key = str(row.get(match1, "")).strip()
                if not case_sensitive:
                    key = key.lower()

                if key in lookup:
                    vals = lookup[key]
                    if len(vals) == 1:
                        result.at[idx, data1] = vals[0]
                        matched += 1
                    else:
                        # Multiple matches
                        if match_type == "Exact Match":
                            result.at[idx, data1] = vals[0]
                        else:
                            result.at[idx, data1] = ", ".join(str(v) for v in vals)
                        matched += 1
                        multi += 1
                else:
                    # Partial matching for non-exact types
                    if match_type != "Exact Match":
                        found = self._partial_match(key, lookup, match_type, case_sensitive)
                        if found:
                            result.at[idx, data1] = found[0] if len(found) == 1 else ", ".join(str(v) for v in found)
                            matched += 1
                            if len(found) > 1:
                                multi += 1
                        else:
                            result.at[idx, data1] = self._no_match_val(nomatch, custom_val)
                            unmatched += 1
                    else:
                        result.at[idx, data1] = self._no_match_val(nomatch, custom_val)
                        unmatched += 1

                if i % 100 == 0:
                    pct = 25 + int(65 * i / total)
                    self._progress["value"] = min(90, pct)
                    self._update_status(f"Processing... {i}/{total} rows ({matched} matched)")

            self.result_df = result

            # Update stats
            self._progress["value"] = 100
            self._update_status(
                f"Done! {matched} matched ({multi} multi), {unmatched} unmatched"
            )
            self._set_stats(total, matched, unmatched, multi)

            # Show result in preview
            self._show_result_preview()

            self._save_btn.config(state="normal")
            self._process_btn.config(state="normal", text="🚀 MATCH & FILL")
            self._processing = False

        except Exception as e:
            self._update_status(f"Error: {str(e)}")
            self._process_btn.config(state="normal", text="🚀 MATCH & FILL")
            self._processing = False
            traceback.print_exc()
            messagebox.showerror("Processing Error", str(e))

    def _partial_match(self, key, lookup, match_type, case_sensitive):
        """Find partial matches across all lookup keys."""
        results = []
        for lk, vals in lookup.items():
            compare_key = lk if case_sensitive else lk.lower()
            compare_target = key if case_sensitive else key.lower()

            ok = False
            if match_type == "Contains":
                ok = compare_target in compare_key or compare_key in compare_target
            elif match_type == "Starts With":
                ok = compare_target.startswith(compare_key)
            elif match_type == "Ends With":
                ok = compare_target.endswith(compare_key)

            if ok:
                results.extend(vals)

        return results if results else None

    def _no_match_val(self, nomatch, custom_val):
        if nomatch == "Fill N/A":
            return "N/A"
        elif nomatch == "Custom Value":
            return custom_val
        return ""

    def _show_result_preview(self):
        if self.result_df is None:
            return

        for item in self._tree.get_children():
            self._tree.delete(item)

        df = self.result_df
        cols = list(df.columns)
        self._tree["columns"] = cols
        self._tree["displaycolumns"] = cols[:20]

        self._tree.heading("#0", text="#")
        self._tree.column("#0", width=40, minwidth=40, anchor="center")
        for c in cols[:20]:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=100, minwidth=60)

        for i, (_, row) in enumerate(df.head(10).iterrows()):
            vals = [str(row.get(c, ""))[:60] for c in cols[:20]]
            self._tree.insert("", "end", text=str(i + 1), values=vals)

        fname = os.path.basename(self.file1_path)
        self._preview_info.config(text=f"Result (first 10 rows) — {fname}")

    # =================================================================
    # SAVE
    # =================================================================
    def _save_result(self):
        if self.result_df is None:
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=f"VLOOKUP_Result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        if not path:
            return

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".csv":
                self.result_df.to_csv(path, index=False)
            else:
                self.result_df.to_excel(path, index=False, engine="openpyxl")
            self._update_status(f"Saved: {os.path.basename(path)}")
            messagebox.showinfo("Success", f"File saved successfully!\n{path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    # =================================================================
    # RESET
    # =================================================================
    def _reset_all(self):
        self.df1 = None
        self.df2 = None
        self.result_df = None
        self.file1_path = ""
        self.file2_path = ""
        self._processing = False

        for n in (1, 2):
            getattr(self, f"_path_lbl_{n}").config(text="No file selected")
            getattr(self, f"_match_var_{n}").set("")
            getattr(self, f"_data_var_{n}").set("")
            getattr(self, f"_match_cb_{n}")["values"] = []
            getattr(self, f"_data_cb_{n}")["values"] = []

        for item in self._tree.get_children():
            self._tree.delete(item)

        self._progress["value"] = 0
        self._set_stats(0, 0, 0, 0)
        self._save_btn.config(state="disabled")
        self._process_btn.config(state="normal", text="🚀 MATCH & FILL")
        self._preview_info.config(text="")
        self._update_status("Reset complete. Load files to begin.")

    def _on_close(self):
        if self._processing:
            if not messagebox.askyesno("Quit?", "Processing in progress. Quit anyway?"):
                return
        self.root.destroy()

    # =================================================================
    # RUN
    # =================================================================
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = VLookupPro()
    app.run()
