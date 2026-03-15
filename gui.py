import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import categorizer
import data_manager
import charts

CATEGORIES_PATH = "categories.json"
DATA_DIR = "data"
YELLOW = "#fffacd"


class ExpenseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CIBC Expense Tracker")
        self.geometry("1000x680")
        self.resizable(True, True)

        self.rules = categorizer.load_rules(CATEGORIES_PATH)
        self._preview_df = None   # DataFrame staged in Import tab
        self._current_df = None   # DataFrame loaded in Transactions tab
        self._chart_canvas = None

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=6, pady=6)

        self._tab_import = ttk.Frame(notebook)
        self._tab_transactions = ttk.Frame(notebook)
        self._tab_charts = ttk.Frame(notebook)

        notebook.add(self._tab_import, text="  Import  ")
        notebook.add(self._tab_transactions, text="  Transactions  ")
        notebook.add(self._tab_charts, text="  Charts  ")

        self._build_import_tab()
        self._build_transactions_tab()
        self._build_charts_tab()

    # ------------------------------------------------------------------
    # TAB 1 — IMPORT
    # ------------------------------------------------------------------
    def _build_import_tab(self):
        top = ttk.Frame(self._tab_import)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Button(top, text="Load CSV", command=self._load_csv).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Load Excel", command=self._load_excel).pack(side="left", padx=(0, 6))
        self._import_btn = ttk.Button(top, text="Import to Data", command=self._import_data, state="disabled")
        self._import_btn.pack(side="left")

        self._import_status = ttk.Label(top, text="")
        self._import_status.pack(side="left", padx=10)

        cols = ("Date", "Description", "Debit", "Credit", "Category")
        frame = ttk.Frame(self._tab_import)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        self._preview_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self._preview_tree.yview)
        hsb.config(command=self._preview_tree.xview)

        col_widths = {"Date": 90, "Description": 380, "Debit": 80, "Credit": 80, "Category": 130}
        for col in cols:
            self._preview_tree.heading(col, text=col)
            self._preview_tree.column(col, width=col_widths[col], anchor="w")

        self._preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._preview_tree.tag_configure("uncategorized", background=YELLOW)

    def _load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            df = data_manager.parse_cibc_csv(path)
            df = data_manager.apply_categories(df, self.rules)
            self._preview_df = df
            self._populate_preview_tree(df)
            self._import_btn.config(state="normal")
            self._import_status.config(text=f"Loaded {len(df)} rows from {path.split('/')[-1]}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not parse CSV:\n{e}")

    def _load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not path:
            return
        try:
            saved = data_manager.import_excel(path, DATA_DIR)
            messagebox.showinfo("Excel Import", f"Imported {len(saved)} sheet(s) from Excel.")
            self._import_status.config(text=f"Excel imported: {len(saved)} sheet(s)")
            self._refresh_month_selectors()
        except Exception as e:
            messagebox.showerror("Error", f"Could not import Excel:\n{e}")

    def _populate_preview_tree(self, df, max_rows=None):
        for row in self._preview_tree.get_children():
            self._preview_tree.delete(row)
        subset = df if max_rows is None else df.head(max_rows)
        for _, r in subset.iterrows():
            tag = "uncategorized" if r.get("Category", "") == "Uncategorized" else ""
            vals = (r.get("Date", ""), r.get("Description", ""),
                    r.get("Debit", ""), r.get("Credit", ""), r.get("Category", ""))
            self._preview_tree.insert("", "end", values=vals, tags=(tag,))

    def _import_data(self):
        if self._preview_df is None:
            return
        try:
            path = data_manager.save_month(self._preview_df, DATA_DIR)
            self._import_status.config(text=f"Saved to {path}")
            self._import_btn.config(state="disabled")
            self._refresh_month_selectors()
            messagebox.showinfo("Imported", f"Data saved to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save data:\n{e}")

    # ------------------------------------------------------------------
    # TAB 2 — TRANSACTIONS
    # ------------------------------------------------------------------
    def _build_transactions_tab(self):
        top = ttk.Frame(self._tab_transactions)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Month:").pack(side="left")
        self._month_var = tk.StringVar()
        self._month_combo = ttk.Combobox(top, textvariable=self._month_var, width=12, state="readonly")
        self._month_combo.pack(side="left", padx=(4, 10))
        self._month_combo.bind("<<ComboboxSelected>>", self._on_month_selected)

        ttk.Button(top, text="Save Changes", command=self._save_changes).pack(side="left")

        self._tx_status = ttk.Label(top, text="")
        self._tx_status.pack(side="left", padx=10)

        cols = ("Date", "Description", "Debit", "Credit", "Category")
        frame = ttk.Frame(self._tab_transactions)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        self._tx_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self._tx_tree.yview)
        hsb.config(command=self._tx_tree.xview)

        col_widths = {"Date": 90, "Description": 370, "Debit": 80, "Credit": 80, "Category": 140}
        for col in cols:
            self._tx_tree.heading(col, text=col)
            self._tx_tree.column(col, width=col_widths[col], anchor="w")

        self._tx_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._tx_tree.tag_configure("uncategorized", background=YELLOW)
        self._tx_tree.bind("<Double-1>", self._on_tx_double_click)

        self._refresh_month_selectors()

    def _refresh_month_selectors(self):
        months = data_manager.list_months(DATA_DIR)
        self._month_combo["values"] = months
        if months and not self._month_var.get():
            self._month_var.set(months[-1])
            self._load_month_data(months[-1])
        # Also refresh chart month selector if it exists
        if hasattr(self, "_chart_month_combo"):
            self._chart_month_combo["values"] = ["All time"] + months
            if not self._chart_month_var.get():
                self._chart_month_var.set("All time")

    def _on_month_selected(self, _event=None):
        ym = self._month_var.get()
        if ym:
            self._load_month_data(ym)

    def _load_month_data(self, year_month):
        df = data_manager.load_month(year_month, DATA_DIR)
        self._current_df = df
        self._populate_tx_tree(df)
        self._tx_status.config(text=f"{len(df)} transactions")

    def _populate_tx_tree(self, df):
        for row in self._tx_tree.get_children():
            self._tx_tree.delete(row)
        for _, r in df.iterrows():
            tag = "uncategorized" if str(r.get("Category", "")) == "Uncategorized" else ""
            vals = (r.get("Date", ""), r.get("Description", ""),
                    r.get("Debit", ""), r.get("Credit", ""), r.get("Category", ""))
            self._tx_tree.insert("", "end", values=vals, tags=(tag,))

    def _on_tx_double_click(self, event):
        """Open an inline editor for the Category cell."""
        region = self._tx_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self._tx_tree.identify_column(event.x)
        col_num = int(col_id.replace("#", "")) - 1
        cols = ("Date", "Description", "Debit", "Credit", "Category")
        if cols[col_num] != "Category":
            return

        item = self._tx_tree.identify_row(event.y)
        if not item:
            return

        # Get bounding box of the cell
        bbox = self._tx_tree.bbox(item, col_id)
        if not bbox:
            return

        current_val = self._tx_tree.item(item, "values")[col_num]

        all_categories = list(self.rules.keys()) + ["Uncategorized", "Payment"]
        var = tk.StringVar(value=current_val)
        combo = ttk.Combobox(self._tx_tree, textvariable=var, values=all_categories, state="readonly")
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        combo.focus_set()

        def apply(_event=None):
            new_val = var.get()
            vals = list(self._tx_tree.item(item, "values"))
            vals[col_num] = new_val
            tag = "uncategorized" if new_val == "Uncategorized" else ""
            self._tx_tree.item(item, values=vals, tags=(tag,))
            combo.destroy()

        combo.bind("<<ComboboxSelected>>", apply)
        combo.bind("<FocusOut>", lambda _e: combo.destroy())

    def _save_changes(self):
        ym = self._month_var.get()
        if not ym or self._current_df is None:
            return
        cols = ("Date", "Description", "Debit", "Credit", "Category")
        rows = [self._tx_tree.item(iid, "values") for iid in self._tx_tree.get_children()]
        df = pd.DataFrame(rows, columns=cols)
        import os
        import data_manager as dm
        path = os.path.join(DATA_DIR, f"{ym}.csv")
        df.to_csv(path, index=False)
        self._current_df = df
        self._tx_status.config(text=f"Saved {len(df)} rows to {path}")

    # ------------------------------------------------------------------
    # TAB 3 — CHARTS
    # ------------------------------------------------------------------
    def _build_charts_tab(self):
        top = ttk.Frame(self._tab_charts)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Month:").pack(side="left")
        self._chart_month_var = tk.StringVar(value="All time")
        self._chart_month_combo = ttk.Combobox(
            top, textvariable=self._chart_month_var, width=12, state="readonly",
            values=["All time"] + data_manager.list_months(DATA_DIR),
        )
        self._chart_month_combo.pack(side="left", padx=(4, 14))

        ttk.Button(top, text="Category Breakdown", command=lambda: self._show_chart("category")).pack(side="left", padx=3)
        ttk.Button(top, text="Monthly Trend",       command=lambda: self._show_chart("trend")).pack(side="left", padx=3)
        ttk.Button(top, text="Top Merchants",       command=lambda: self._show_chart("merchants")).pack(side="left", padx=3)

        self._chart_frame = ttk.Frame(self._tab_charts)
        self._chart_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

    def _get_chart_df(self):
        ym = self._chart_month_var.get()
        if ym == "All time":
            return data_manager.load_all_months(DATA_DIR)
        return data_manager.load_month(ym, DATA_DIR)

    def _show_chart(self, chart_type):
        df = self._get_chart_df()
        if df.empty:
            messagebox.showinfo("No Data", "No transaction data found. Import a CSV first.")
            return

        if chart_type == "category":
            fig = charts.category_bar(df)
        elif chart_type == "trend":
            all_df = data_manager.load_all_months(DATA_DIR)
            fig = charts.monthly_trend(all_df)
        else:
            fig = charts.top_merchants(df)

        # Clear previous chart
        for widget in self._chart_frame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(fig, master=self._chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._chart_canvas = canvas
