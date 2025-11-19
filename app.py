#!/usr/bin/env python3
import os, json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd

from core import process_file, per_staff_per_day

DND_AVAILABLE = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

APP_TITLE = "Billable Hours Aggregator"
APP_WIDTH = 1120
APP_HEIGHT = 760

DEFAULT_THEMES = {
    "light": {
        "parent_even": "#f7f7fb",
        "parent_odd":  "#ffffff",
        "child":       "#eaeaf2",
        "bg":          "#ffffff",
        "selection":   "#cde4ff",
        "text":        "#000000"
    }
}

SETTINGS_FILE = "settings.json"

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                theme_name = data.get("theme_name", "light")
                if theme_name not in DEFAULT_THEMES:
                    theme_name = "light"
                colors = data.get("colors", DEFAULT_THEMES.get(theme_name, DEFAULT_THEMES["light"]))
                return {"theme_name": theme_name, "colors": colors}
        except Exception:
            pass
    return {"theme_name": "light", "colors": DEFAULT_THEMES["light"].copy()}

def save_settings(theme_name, colors):
    data = {"theme_name": theme_name, "colors": colors}
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

class SettingsDialog(tk.Toplevel):
    def __init__(self, master, current_theme, current_colors, on_apply):
        super().__init__(master)
        self.title("Settings — Colors")
        self.resizable(False, False)
        self.on_apply = on_apply
        self.current_theme = tk.StringVar(value=current_theme)
        self.colors = current_colors.copy()
        self.vars = {}

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        row = 0
        ttk.Label(frm, text="Base Theme:").grid(row=row, column=0, sticky="w")
        theme_cb = ttk.Combobox(frm, values=list(DEFAULT_THEMES.keys()), state="readonly", textvariable=self.current_theme, width=12)
        theme_cb.grid(row=row, column=1, sticky="w", padx=(6,0))
        theme_cb.bind("<<ComboboxSelected>>", self.on_theme_change)
        row += 1
        ttk.Separator(frm, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=6)
        row += 1

        for key, label in [
            ("parent_even", "Parent Row (even)"),
            ("parent_odd",  "Parent Row (odd)"),
            ("child",       "Child Row (per-day)"),
            ("bg",          "Background"),
            ("selection",   "Selection"),
            ("text",        "Text"),
        ]:
            ttk.Label(frm, text=label + ":").grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=self.colors.get(key))
            self.vars[key] = var
            ent = ttk.Entry(frm, textvariable=var, width=15)
            ent.grid(row=row, column=1, sticky="w")
            def make_pick(k=key, v=var):
                def _pick():
                    initial = v.get()
                    (rgb, hexv) = colorchooser.askcolor(initialcolor=initial, title=f"Pick {k} color")
                    if hexv:
                        v.set(hexv)
                return _pick
            btn = ttk.Button(frm, text="Pick...", command=make_pick())
            btn.grid(row=row, column=2, sticky="w", padx=6)
            row += 1

        ttk.Separator(frm, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=8)
        row += 1

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=3, sticky="e")
        ttk.Button(btns, text="Reset to Theme", command=self.reset_to_theme).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Apply", command=self.apply).pack(side=tk.LEFT, padx=4)

        self.grab_set()
        self.transient(master)

    def on_theme_change(self, _evt=None):
        base = DEFAULT_THEMES.get(self.current_theme.get(), DEFAULT_THEMES["light"])
        for k, v in self.vars.items():
            v.set(base[k])

    def reset_to_theme(self):
        self.on_theme_change()

    def apply(self):
        theme_name = self.current_theme.get()
        colors = {k: v.get() for k, v in self.vars.items()}
        self.on_apply(theme_name, colors)
        self.destroy()

class App(tk.Tk if not DND_AVAILABLE else TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(900, 560)
        self.summary_df = None
        self.details_df = None
        self.filtered_df = None
        self.input_path = None

        s = load_settings()
        self.theme_name = s["theme_name"]
        self.colors = s["colors"]

        self._setup_style()
        self._build_ui()
        self._apply_colors()

    def _setup_style(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("default")
        except Exception:
            pass
        self.style.configure("TFrame", background=self.colors.get("bg", "#ffffff"))
        self.style.configure("TLabel", background=self.colors.get("bg", "#ffffff"), foreground=self.colors.get("text", "#000000"))
        self.style.configure("Treeview",
                        background="#ffffff",
                        fieldbackground="#ffffff",
                        rowheight=26,
                        borderwidth=0)
        self.style.configure("Treeview.Heading", font=("", 10, "bold"))
        self.style.map("Treeview",
                  background=[("selected", self.colors.get("selection", "#cde4ff"))],
                  foreground=[("selected", self.colors.get("text", "#000000"))])

    def _apply_colors(self):
        self.configure(background=self.colors["bg"])
        if hasattr(self, "tree"):
            self.tree.tag_configure("parent_even", background=self.colors["parent_even"], foreground=self.colors["text"])
            self.tree.tag_configure("parent_odd", background=self.colors["parent_odd"], foreground=self.colors["text"])
            self.tree.tag_configure("child", background=self.colors["child"], foreground=self.colors["text"])
            self.rebuild_tree()

    def _build_ui(self):
        # Menu
        menubar = tk.Menu(self)
        settings_menu = tk.Menu(menubar, tearoff=0)
        theme_menu = tk.Menu(settings_menu, tearoff=0)
        theme_menu.add_command(label="Light", command=lambda: self.set_theme("light"))
        settings_menu.add_cascade(label="Theme", menu=theme_menu)
        settings_menu.add_command(label="Edit Colors...", command=self.open_settings_dialog)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        self.config(menu=menubar)

        root = ttk.Frame(self, padding=0)
        root.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Top bar
        top = ttk.Frame(root, padding=8)
        top.pack(side=tk.TOP, fill=tk.X)

        left = ttk.Frame(top)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        browse_btn = ttk.Button(left, text="Browse File...", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT)

        self.drop_label = ttk.Label(left, text=("Drag & drop a CSV/XLSX here" if DND_AVAILABLE else "Drag & drop requires 'tkinterdnd2'. Use Browse."),
                                    padding=6, relief=tk.SOLID)
        self.drop_label.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)

        mid = ttk.Frame(top)
        mid.pack(side=tk.LEFT, padx=8)
        ttk.Label(mid, text="Search:").grid(row=0, column=0, sticky="w")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_search_sort())
        search_entry = ttk.Entry(mid, textvariable=self.search_var, width=24)
        search_entry.grid(row=0, column=1, padx=(4,8))
        ttk.Label(mid, text="Sort:").grid(row=0, column=2, sticky="w")
        self.sort_var = tk.StringVar(value="Name A→Z")
        sort_cb = ttk.Combobox(mid, textvariable=self.sort_var, values=["Name A→Z","Minutes ↓"], width=12, state="readonly")
        sort_cb.grid(row=0, column=3)
        # Normalize any corrupted labels in sort combobox to clean ASCII
        try:
            sort_cb.configure(values=("Name A-Z", "Minutes (desc)"))
            self.sort_var.set("Name A-Z")
        except Exception:
            pass
        sort_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_search_sort())

        right = ttk.Frame(top)
        right.pack(side=tk.RIGHT)
        self.export_excel_btn = ttk.Button(right, text="Export Excel", command=self.export_excel, state=tk.DISABLED)
        self.export_excel_btn.grid(row=0, column=0, padx=4)
        self.export_csv_btn = ttk.Button(right, text="Export CSVs", command=self.export_csv, state=tk.DISABLED)
        self.export_csv_btn.grid(row=0, column=1, padx=4)
        self.export_staff_btn = ttk.Button(right, text="Export Selected Staff Details", command=self.export_selected_staff, state=tk.DISABLED)
        self.export_staff_btn.grid(row=0, column=2, padx=4)

        ttk.Separator(root, orient="horizontal").pack(fill=tk.X)

        # Table area
        table_outer = ttk.Frame(root, padding=8)
        table_outer.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        table_frame = ttk.Frame(table_outer)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        cols = ("Staff Name", "Total Minutes", "Total Hours", "Total Units", "Appt Start", "Appt End")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("Staff Name", anchor=tk.W, width=280, stretch=True)
        self.tree.column("Total Minutes", anchor=tk.W, width=120, stretch=False)
        self.tree.column("Total Hours", anchor=tk.W, width=120, stretch=False)
        self.tree.column("Total Units", anchor=tk.W, width=110, stretch=False)
        self.tree.column("Appt Start", anchor=tk.W, width=110, stretch=False)
        self.tree.column("Appt End", anchor=tk.W, width=110, stretch=False)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Treeview tags (colors applied in _apply_colors)
        self.tree.tag_configure("parent_even")
        self.tree.tag_configure("parent_odd")
        self.tree.tag_configure("child")

        self.tree.bind("<Double-1>", self.on_double_click)

        ttk.Separator(root, orient="horizontal").pack(fill=tk.X)

        # Footer
        footer = ttk.Frame(root, padding=(8,6))
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        self.footer_var = tk.StringVar(value="Totals: 0 minutes | 0.00 hours | 0 units")
        footer_label = ttk.Label(footer, textvariable=self.footer_var, anchor="w")
        footer_label.pack(side=tk.LEFT)

        # Status bar
        self.status = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status, anchor="e", padding=(8,6))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self._on_drop)

    # Theme / Settings
    def set_theme(self, name):
        self.theme_name = name if name in DEFAULT_THEMES else "light"
        self.colors = DEFAULT_THEMES[self.theme_name].copy()
        save_settings(self.theme_name, self.colors)
        self._setup_style()
        self._apply_colors()

    def open_settings_dialog(self):
        def on_apply(theme_name, colors):
            # Sanitize theme to light only
            self.theme_name = theme_name if theme_name in DEFAULT_THEMES else "light"
            self.colors = colors
            save_settings(self.theme_name, self.colors)
            self._setup_style()
            self._apply_colors()
        SettingsDialog(self, self.theme_name, self.colors, on_apply)

    # File load / UI behaviors
    def browse_file(self):
        path = filedialog.askopenfilename(title="Select CSV or Excel file",
                                          filetypes=[("Excel", "*.xlsx;*.xlsm;*.xltx;*.xltm"),
                                                     ("CSV", "*.csv"),
                                                     ("All files", "*.*")])
        if not path:
            return
        self.load_and_display(Path(path))

    def _on_drop(self, event):
        raw = event.data
        items = []
        cur = ""
        in_brace = False
        for ch in raw:
            if ch == "{":
                in_brace = True
                cur = ""
            elif ch == "}":
                in_brace = False
                items.append(cur)
                cur = ""
            elif ch == " " and not in_brace:
                if cur:
                    items.append(cur)
                    cur = ""
            else:
                cur += ch
        if cur:
            items.append(cur)
        if not items:
            return
        self.load_and_display(Path(items[0]))

    def load_and_display(self, path: Path):
        try:
            self.status.set(f"Processing: {path.name}..."); self.update_idletasks()
            summary, details = process_file(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{e}")
            self.status.set("Ready"); return
        self.input_path = path
        self.summary_df = summary
        self.details_df = details
        self.apply_search_sort()
        self.export_excel_btn.config(state=tk.NORMAL)
        self.export_csv_btn.config(state=tk.NORMAL)
        self.export_staff_btn.config(state=tk.NORMAL)
        self.status.set(f"Loaded {path.name} - {len(self.summary_df)} staff")

    def apply_search_sort(self):
        if self.summary_df is None: return
        df = self.summary_df.copy()
        q = self.search_var.get().strip().lower()
        if q:
            df = df[df["Staff Name"].str.lower().str.contains(q, na=False)]
        # Robust sort handling independent of label text
        opt_val = (self.sort_var.get() or "").lower()
        if "minute" in opt_val:
            df = df.sort_values(["Total_Minutes", "Staff Name"], ascending=[False, True])
        else:
            df = df.sort_values(["Staff Name"], ascending=[True])
        self.filtered_df = df
        self.rebuild_tree()
        self.update_footer()
        return
        if self.sort_var.get() == "Minutes ↓":
            df = df.sort_values(["Total_Minutes","Staff Name"], ascending=[False, True])
        else:
            df = df.sort_values(["Staff Name"], ascending=[True])
        self.filtered_df = df
        self.rebuild_tree()
        self.update_footer()

    def rebuild_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.filtered_df is None: return
        for i, (_, row) in enumerate(self.filtered_df.iterrows()):
            tag = "parent_even" if (i % 2 == 0) else "parent_odd"
            parent_id = self.tree.insert("", "end", values=(
                row["Staff Name"],
                int(row["Total_Minutes"]),
                f"{row['Total_Hours']:.2f}",
                int(row["Total_Units"]),
                "",
                ""
            ), tags=(tag, "parent"))
            self.tree.insert(parent_id, "end", values=("", "", "", "", "", ""), tags=("child",))

    def on_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        parent = item_id
        children = self.tree.get_children(parent)
        if children:
            vals = self.tree.item(children[0], "values")
            if any(vals):
                for c in children: self.tree.delete(c)
                self.tree.insert(parent, "end", values=("", "", "", "", "", ""), tags=("child",))
                return
            else:
                self.tree.delete(children[0])
        vals = self.tree.item(parent, "values")
        if not vals: return
        staff_name = vals[0]
        daily = per_staff_per_day(self.details_df, staff_name)
        # Insert a small note row if this staff has incomplete sessions excluded
        try:
            cnt = 0
            if self.details_df is not None and not self.details_df.empty and "Incomplete_Excluded_Count" in self.details_df.columns:
                _rows = self.details_df[self.details_df["Staff Name"].astype(str) == str(staff_name)]
                if not _rows.empty:
                    _val = pd.to_numeric(_rows["Incomplete_Excluded_Count"], errors='coerce').max()
                    cnt = int(_val) if pd.notna(_val) else 0
            if cnt > 0:
                self.tree.insert(parent, "end", values=(f"  Note: {cnt} incomplete session(s) were excluded from billing", "", "", "", "", ""), tags=("child",))
        except Exception:
            pass
        if daily.empty:
            self.tree.insert(parent, "end", values=("  (no dated rows)", "", "", "", "", ""), tags=("child",)); return
        for _, r in daily.iterrows():
            self.tree.insert(parent, "end", values=(
                f"  {r['Appt. Date']}",
                int(r["Total_Minutes"]),
                f"{r['Total_Hours']:.2f}",
                int(r["Total_Units"]),
                r.get("Appt Start",""),
                r.get("Appt End","")
            ), tags=("child",))

    def ensure_outdir(self):
        if self.input_path is None: raise RuntimeError("No input file loaded")
        outdir = self.input_path.parent / (self.input_path.stem + "_out")
        outdir.mkdir(exist_ok=True)
        return outdir

    def export_excel(self):
        try:
            outdir = self.ensure_outdir()
            xlsx = outdir / "results.xlsx"
            with pd.ExcelWriter(xlsx, engine="openpyxl") as xl:
                self.filtered_df.to_excel(xl, index=False, sheet_name="Summary")
                self.details_df.to_excel(xl, index=False, sheet_name="Details")
            messagebox.showinfo("Export", f"Wrote: {xlsx}")
            self.status.set(f"Exported Excel to {xlsx}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def export_csv(self):
        try:
            outdir = self.ensure_outdir()
            self.filtered_df.to_csv(outdir / "summary.csv", index=False)
            self.details_df.to_csv(outdir / "details.csv", index=False)
            messagebox.showinfo("Export", f"Wrote: {outdir / 'summary.csv'}\n{outdir / 'details.csv'}")
            self.status.set(f"Exported CSVs to {outdir}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def export_selected_staff(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Export Staff", "Select a staff row first."); return
        item_id = sel[0]
        parent = self.tree.parent(item_id) or item_id
        vals = self.tree.item(parent, "values")
        if not vals or not vals[0]:
            messagebox.showwarning("Export Staff", "Select a staff row (not a date row)."); return
        staff_name = vals[0]
        try:
            outdir = self.ensure_outdir()
            daily = per_staff_per_day(self.details_df, staff_name)
            daily.to_csv(outdir / f"{staff_name}_by_day.csv", index=False)
            staff_rows = self.details_df[self.details_df["Staff Name"] == staff_name]
            staff_rows.to_csv(outdir / f"{staff_name}_details.csv", index=False)
            messagebox.showinfo("Export Staff", f"Exported:\n{outdir / (staff_name + '_by_day.csv')}\n{outdir / (staff_name + '_details.csv')}")
            self.status.set(f"Exported staff details for {staff_name}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def update_footer(self):
        if self.filtered_df is None or self.filtered_df.empty:
            self.footer_var.set("Totals: 0 minutes | 0.00 hours | 0 units"); return
        mins = int(self.filtered_df["Total_Minutes"].sum())
        hours = round(mins / 60.0, 2)
        units = int(round(mins / 15.0))
        self.footer_var.set(f"Totals: {mins} minutes | {hours:.2f} hours | {units} units")

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
