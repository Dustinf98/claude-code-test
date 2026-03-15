import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import categorizer
import data_manager
import charts

CATEGORIES_PATH = "categories.json"
DATA_DIR = "data"

# ── Catppuccin Mocha palette ──────────────────────────────────────────────────
BG_BASE     = "#1e1e2e"
BG_MANTLE   = "#181825"
BG_SURFACE  = "#313244"
BG_SURFACE2 = "#45475a"
FG_TEXT     = "#cdd6f4"
FG_SUBTEXT  = "#a6adc8"
ACCENT      = "#89b4fa"
GREEN       = "#a6e3a1"
RED         = "#f38ba8"
YELLOW      = "#f9e2af"

CATEGORY_COLORS = {
    "Groceries":      "#2d6a4f",
    "Gas":            "#6b2737",
    "Restaurants":    "#7c4f1e",
    "Shopping":       "#1a4a6b",
    "Subscriptions":  "#4a2d6b",
    "Activities":     "#1a5f5f",
    "Gym":            "#5c4200",
    "Car Repairs":    "#3d3d3d",
    "Miscellaneous":  "#2e2e2e",
    "Phone":          "#1a4a4a",
    "Insurance":      "#1e3a5f",
    "Rent":           "#4a1a5f",
    "Internet Stuff": "#1a3a3a",
    "Payment":        "#1a4d2e",
    "Uncategorized":  "#5c1a1a",
}

NAV_ITEMS = [
    ("📥", "Import",       "import"),
    ("📋", "Transactions", "transactions"),
    ("📈", "Charts",       "charts"),
]


def _configure_tree_tags(tree):
    """Apply category background/foreground color tags to a Treeview."""
    for cat, bg in CATEGORY_COLORS.items():
        fg = RED if cat == "Uncategorized" else FG_TEXT
        tree.tag_configure(f"cat_{cat}", background=bg, foreground=fg)
    # Legacy tag kept for any pre-existing code paths
    tree.tag_configure("uncategorized", background=CATEGORY_COLORS["Uncategorized"], foreground=RED)


def _category_tag(cat):
    """Return the Treeview row tag for a given category string."""
    cat = str(cat)
    if cat == "Uncategorized":
        return "uncategorized"
    return f"cat_{cat}" if cat in CATEGORY_COLORS else ""


class ExpenseApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("CIBC Expense Tracker")
        self.geometry("1100x720")
        self.resizable(True, True)
        self.configure(bg=BG_BASE)

        self.rules = categorizer.load_rules(CATEGORIES_PATH)
        self._preview_df = None
        self._current_df = None
        self._chart_canvas = None
        self._nav_items = {}
        self._active_nav = None

        self._apply_styles()
        self._build_layout()
        self._refresh_month_selectors()
        self._navigate("import")

    # ── Treeview style overrides ───────────────────────────────────────────────
    def _apply_styles(self):
        style = ttk.Style()
        style.configure("Treeview",
            background=BG_SURFACE,
            foreground=FG_TEXT,
            fieldbackground=BG_SURFACE,
            rowheight=28,
            font=("Segoe UI", 10),
        )
        style.configure("Treeview.Heading",
            background=BG_MANTLE,
            foreground=ACCENT,
            font=("Segoe UI", 10, "bold"),
        )
        style.map("Treeview",
            background=[("selected", ACCENT)],
            foreground=[("selected", BG_BASE)],
        )

    # ── Top-level layout ───────────────────────────────────────────────────────
    def _build_layout(self):
        # Sidebar
        sidebar = tk.Frame(self, bg=BG_MANTLE, width=160)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="💸 Expenses", bg=BG_MANTLE, fg=FG_TEXT,
                 font=("Segoe UI", 11, "bold"), pady=20).pack(fill="x")
        tk.Frame(sidebar, bg=BG_SURFACE2, height=1).pack(fill="x", padx=10, pady=(0, 4))

        for icon, label, key in NAV_ITEMS:
            self._build_nav_item(sidebar, icon, label, key)

        tk.Frame(sidebar, bg=BG_SURFACE2, height=1).pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(sidebar, text="v1.0", bg=BG_MANTLE, fg=FG_SUBTEXT,
                 font=("Segoe UI", 8), pady=8).pack(side="bottom")

        # Content area
        content = tk.Frame(self, bg=BG_BASE)
        content.pack(side="left", fill="both", expand=True)

        self._frame_import       = tk.Frame(content, bg=BG_BASE)
        self._frame_transactions = tk.Frame(content, bg=BG_BASE)
        self._frame_charts       = tk.Frame(content, bg=BG_BASE)

        self._build_import_tab()
        self._build_transactions_tab()
        self._build_charts_tab()

    def _build_nav_item(self, sidebar, icon, label, key):
        item_frame = tk.Frame(sidebar, bg=BG_MANTLE, cursor="hand2")
        item_frame.pack(fill="x")

        border = tk.Frame(item_frame, width=4, bg=BG_MANTLE)
        border.pack(side="left", fill="y")

        inner = tk.Frame(item_frame, bg=BG_MANTLE)
        inner.pack(side="left", fill="both", expand=True)

        icon_lbl = tk.Label(inner, text=icon, bg=BG_MANTLE, fg=FG_SUBTEXT,
                            font=("Segoe UI", 13), padx=10, pady=14)
        icon_lbl.pack(side="left")

        text_lbl = tk.Label(inner, text=label, bg=BG_MANTLE, fg=FG_SUBTEXT,
                            font=("Segoe UI", 10))
        text_lbl.pack(side="left")

        self._nav_items[key] = {
            "frame": item_frame, "border": border,
            "inner": inner, "icon": icon_lbl, "text": text_lbl,
        }

        for widget in (item_frame, border, inner, icon_lbl, text_lbl):
            widget.bind("<Button-1>", lambda e, k=key: self._navigate(k))
            widget.bind("<Enter>",    lambda e, k=key: self._nav_hover(k, True))
            widget.bind("<Leave>",    lambda e, k=key: self._nav_hover(k, False))

    def _nav_hover(self, key, entering):
        if key == self._active_nav:
            return
        item = self._nav_items[key]
        bg = BG_SURFACE if entering else BG_MANTLE
        fg = FG_TEXT    if entering else FG_SUBTEXT
        for w in (item["frame"], item["inner"], item["icon"], item["text"]):
            w.configure(bg=bg)
        item["icon"].configure(fg=fg)
        item["text"].configure(fg=fg)

    def _navigate(self, key):
        if self._active_nav:
            old = self._nav_items[self._active_nav]
            old["border"].configure(bg=BG_MANTLE)
            for w in (old["frame"], old["inner"], old["icon"], old["text"]):
                w.configure(bg=BG_MANTLE)
            old["icon"].configure(fg=FG_SUBTEXT)
            old["text"].configure(fg=FG_SUBTEXT)

        self._active_nav = key
        item = self._nav_items[key]
        item["border"].configure(bg=ACCENT)
        for w in (item["frame"], item["inner"], item["icon"], item["text"]):
            w.configure(bg=BG_SURFACE)
        item["icon"].configure(fg=ACCENT)
        item["text"].configure(fg=ACCENT)

        frame_map = {
            "import":       self._frame_import,
            "transactions": self._frame_transactions,
            "charts":       self._frame_charts,
        }
        for k, frame in frame_map.items():
            if k == key:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()

    # ── TAB 1 — IMPORT ────────────────────────────────────────────────────────
    def _build_import_tab(self):
        top = tk.Frame(self._frame_import, bg=BG_BASE)
        top.pack(fill="x", padx=14, pady=10)

        ttk.Button(top, text="Load CSV", command=self._load_csv,
                   bootstyle="info-outline").pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Load Excel", command=self._load_excel,
                   bootstyle="secondary-outline").pack(side="left", padx=(0, 6))
        self._import_btn = ttk.Button(top, text="Import to Data",
                                      command=self._import_data,
                                      state="disabled", bootstyle="success")
        self._import_btn.pack(side="left")

        self._import_status = tk.Label(top, text="", bg=BG_BASE, fg=FG_SUBTEXT,
                                       font=("Segoe UI", 9))
        self._import_status.pack(side="left", padx=10)

        cols = ("Date", "Description", "Debit", "Credit", "Category")
        frame = tk.Frame(self._frame_import, bg=BG_BASE)
        frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        self._preview_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self._preview_tree.yview)
        hsb.config(command=self._preview_tree.xview)

        col_widths = {"Date": 90, "Description": 400, "Debit": 90, "Credit": 90, "Category": 150}
        for col in cols:
            self._preview_tree.heading(col, text=col)
            self._preview_tree.column(col, width=col_widths[col], anchor="w")

        self._preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        _configure_tree_tags(self._preview_tree)

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
            cat = str(r.get("Category", ""))
            vals = (r.get("Date", ""), r.get("Description", ""),
                    r.get("Debit", ""), r.get("Credit", ""), cat)
            self._preview_tree.insert("", "end", values=vals, tags=(_category_tag(cat),))

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

    # ── TAB 2 — TRANSACTIONS ──────────────────────────────────────────────────
    def _build_transactions_tab(self):
        top = tk.Frame(self._frame_transactions, bg=BG_BASE)
        top.pack(fill="x", padx=14, pady=10)

        tk.Label(top, text="Month:", bg=BG_BASE, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side="left")
        self._month_var = tk.StringVar()
        self._month_combo = ttk.Combobox(top, textvariable=self._month_var,
                                         width=12, state="readonly")
        self._month_combo.pack(side="left", padx=(4, 10))
        self._month_combo.bind("<<ComboboxSelected>>", self._on_month_selected)

        ttk.Button(top, text="Save Changes", command=self._save_changes,
                   bootstyle="primary").pack(side="left")

        self._tx_status = tk.Label(top, text="", bg=BG_BASE, fg=FG_SUBTEXT,
                                   font=("Segoe UI", 9))
        self._tx_status.pack(side="left", padx=10)

        # Summary stats bar
        self._stats_bar = self._build_stats_bar(self._frame_transactions)

        # Treeview
        cols = ("Date", "Description", "Debit", "Credit", "Category")
        frame = tk.Frame(self._frame_transactions, bg=BG_BASE)
        frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        self._tx_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self._tx_tree.yview)
        hsb.config(command=self._tx_tree.xview)

        col_widths = {"Date": 90, "Description": 390, "Debit": 90, "Credit": 90, "Category": 150}
        for col in cols:
            self._tx_tree.heading(col, text=col)
            self._tx_tree.column(col, width=col_widths[col], anchor="w")

        self._tx_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        _configure_tree_tags(self._tx_tree)
        self._tx_tree.bind("<Double-1>", self._on_tx_double_click)

    def _build_stats_bar(self, parent):
        bar = tk.Frame(parent, bg=BG_BASE)
        bar.pack(fill="x", padx=14, pady=(0, 10))

        stats = {}
        for i, (key, title) in enumerate([
            ("spent",   "Total Spent"),
            ("top_cat", "Top Category"),
            ("count",   "Transactions"),
        ]):
            card = tk.Frame(bar, bg=BG_SURFACE, padx=16, pady=10)
            card.grid(row=0, column=i, padx=(0, 10), sticky="nsew")

            val_lbl = tk.Label(card, text="—", bg=BG_SURFACE, fg=ACCENT,
                               font=("Segoe UI", 16, "bold"))
            val_lbl.pack()
            tk.Label(card, text=title, bg=BG_SURFACE, fg=FG_SUBTEXT,
                     font=("Segoe UI", 9)).pack()
            stats[key] = val_lbl

        bar.columnconfigure(0, weight=1)
        bar.columnconfigure(1, weight=1)
        bar.columnconfigure(2, weight=1)
        return stats

    def _update_stats_bar(self):
        if self._current_df is None or self._current_df.empty:
            return
        df = self._current_df.copy()
        df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)

        if "Category" in df.columns:
            expenses = df[df["Category"] != "Payment"]
        else:
            expenses = df

        total = expenses["Debit"].sum()
        count = len(df)

        expenses_pos = expenses[expenses["Debit"] > 0]
        if not expenses_pos.empty and "Category" in expenses_pos.columns:
            top_cat = expenses_pos.groupby("Category")["Debit"].sum().idxmax()
        else:
            top_cat = "—"

        self._stats_bar["spent"].config(text=f"${total:,.2f}")
        self._stats_bar["top_cat"].config(text=str(top_cat))
        self._stats_bar["count"].config(text=str(count))

    def _refresh_month_selectors(self):
        months = data_manager.list_months(DATA_DIR)
        self._month_combo["values"] = months
        if months and not self._month_var.get():
            self._month_var.set(months[-1])
            self._load_month_data(months[-1])
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
        self._update_stats_bar()

    def _populate_tx_tree(self, df):
        for row in self._tx_tree.get_children():
            self._tx_tree.delete(row)
        for _, r in df.iterrows():
            cat = str(r.get("Category", ""))
            vals = (r.get("Date", ""), r.get("Description", ""),
                    r.get("Debit", ""), r.get("Credit", ""), cat)
            self._tx_tree.insert("", "end", values=vals, tags=(_category_tag(cat),))

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

        bbox = self._tx_tree.bbox(item, col_id)
        if not bbox:
            return

        current_val = self._tx_tree.item(item, "values")[col_num]
        all_categories = list(self.rules.keys()) + ["Uncategorized", "Payment"]
        var = tk.StringVar(value=current_val)
        combo = ttk.Combobox(self._tx_tree, textvariable=var,
                             values=all_categories, state="readonly")
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        combo.focus_set()

        def apply(_event=None):
            new_val = var.get()
            vals = list(self._tx_tree.item(item, "values"))
            vals[col_num] = new_val
            self._tx_tree.item(item, values=vals, tags=(_category_tag(new_val),))
            combo.destroy()

        combo.bind("<<ComboboxSelected>>", apply)
        combo.bind("<FocusOut>", lambda _e: combo.destroy())

    def _save_changes(self):
        ym = self._month_var.get()
        if not ym or self._current_df is None:
            return
        import os
        cols = ("Date", "Description", "Debit", "Credit", "Category")
        rows = [self._tx_tree.item(iid, "values") for iid in self._tx_tree.get_children()]
        df = pd.DataFrame(rows, columns=cols)
        path = os.path.join(DATA_DIR, f"{ym}.csv")
        df.to_csv(path, index=False)
        self._current_df = df
        self._tx_status.config(text=f"Saved {len(df)} rows to {path}")
        self._update_stats_bar()

    # ── TAB 3 — CHARTS ────────────────────────────────────────────────────────
    def _build_charts_tab(self):
        top = tk.Frame(self._frame_charts, bg=BG_BASE)
        top.pack(fill="x", padx=14, pady=10)

        tk.Label(top, text="Month:", bg=BG_BASE, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side="left")
        self._chart_month_var = tk.StringVar(value="All time")
        self._chart_month_combo = ttk.Combobox(
            top, textvariable=self._chart_month_var, width=12, state="readonly",
            values=["All time"] + data_manager.list_months(DATA_DIR),
        )
        self._chart_month_combo.pack(side="left", padx=(4, 14))

        ttk.Button(top, text="Category Breakdown", bootstyle="info-outline",
                   command=lambda: self._show_chart("category")).pack(side="left", padx=3)
        ttk.Button(top, text="Monthly Trend", bootstyle="info-outline",
                   command=lambda: self._show_chart("trend")).pack(side="left", padx=3)
        ttk.Button(top, text="Top Merchants", bootstyle="info-outline",
                   command=lambda: self._show_chart("merchants")).pack(side="left", padx=3)

        self._chart_frame = tk.Frame(self._frame_charts, bg=BG_BASE)
        self._chart_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

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

        for widget in self._chart_frame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(fig, master=self._chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._chart_canvas = canvas
