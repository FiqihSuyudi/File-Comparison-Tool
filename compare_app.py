import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd


# =========================
# CONFIG
# =========================
NUMERIC_TOLERANCE = 0.01
NUMERIC_COLUMNS = {"Stunden", "Zuschlag", "Normalarbeitszeit", "KST"}


# =========================
# HELPERS
# =========================
def norm(x):
    if pd.isna(x):
        return None
    return str(x).strip().lower()


def norm_date(x):
    if pd.isna(x):
        return None
    dt = pd.to_datetime(x, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.date()


def norm_number(x):
    if pd.isna(x):
        return None

    s = str(x).strip()
    if s == "":
        return None

    s = s.replace("−", "-").replace("—", "-").replace("–", "-")
    s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return None


def numbers_equal(a, b, tolerance=NUMERIC_TOLERANCE):
    av = norm_number(a)
    bv = norm_number(b)

    if av is None and bv is None:
        return True
    if av is None or bv is None:
        return False

    return abs(av - bv) <= tolerance


def norm_by_column(col_name, value):
    col_lower = str(col_name).lower()

    if "datum" in col_lower or "date" in col_lower:
        return norm_date(value)

    if col_name in NUMERIC_COLUMNS:
        return norm_number(value)

    return norm(value)


def values_equal(col_name, a, b):
    col_lower = str(col_name).lower()

    if "datum" in col_lower or "date" in col_lower:
        return norm_date(a) == norm_date(b)

    if col_name in NUMERIC_COLUMNS:
        return numbers_equal(a, b)

    return norm(a) == norm(b)


def load_file(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=object, sep=";")
    return pd.read_excel(path, dtype=object)


def file_size_kb(path):
    try:
        return round(os.path.getsize(path) / 1024, 2)
    except Exception:
        return None


# =========================
# MAIN APP
# =========================
class CompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Compare Tool PRO")
        self.root.geometry("1450x820")
        self.root.minsize(1150, 680)

        self.file1_path = None
        self.file2_path = None
        self.df1 = None
        self.df2 = None
        self.result_df = None
        self.column_vars = {}

        self.build_ui()

    # =========================
    # UI
    # =========================
    def build_ui(self):
        top = tk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=8)

        tk.Button(top, text="Load File 1", command=self.load_file1, width=14).pack(side="left")
        self.lbl1 = tk.Label(top, text="No file selected", anchor="w")
        self.lbl1.pack(side="left", padx=(8, 20))

        tk.Button(top, text="Load File 2", command=self.load_file2, width=14).pack(side="left")
        self.lbl2 = tk.Label(top, text="No file selected", anchor="w")
        self.lbl2.pack(side="left", padx=(8, 20))

        diag = tk.LabelFrame(self.root, text="Diagnostics")
        diag.pack(fill="x", padx=10, pady=(0, 8))

        self.diag1 = tk.Label(diag, text="File 1 → name: -, rows: -, cols: -, size: - KB", anchor="w")
        self.diag1.pack(fill="x", padx=8, pady=4)

        self.diag2 = tk.Label(diag, text="File 2 → name: -, rows: -, cols: -, size: - KB", anchor="w")
        self.diag2.pack(fill="x", padx=8, pady=(0, 6))

        main = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        main.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        left = tk.Frame(main, width=340)
        main.add(left)

        tk.Label(left, text="Columns", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        self.smart_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left, text="Smart Matching", variable=self.smart_var).pack(anchor="w", pady=(0, 6))

        helper = tk.Label(
            left,
            text="Ignored columns are removed before compare.\n"
                 "Date keys are normalized as dates.\n"
                 "Numeric columns use tolerance matching.\n"
                 "Preview hides ignored columns and keeps keys first.",
            justify="left",
            fg="gray30"
        )
        helper.pack(anchor="w", pady=(0, 8))

        quick_btns = tk.Frame(left)
        quick_btns.pack(fill="x", pady=(0, 8))

        tk.Button(quick_btns, text="Clear Keys", command=self.clear_keys).pack(side="left", padx=(0, 6))
        tk.Button(quick_btns, text="Clear Ignore", command=self.clear_ignore).pack(side="left", padx=(0, 6))
        tk.Button(quick_btns, text="Clear All", command=self.clear_all_checks).pack(side="left")

        columns_container = tk.Frame(left)
        columns_container.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(columns_container, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(columns_container, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")

        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", self.on_scroll_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        right = tk.Frame(main)
        main.add(right)

        tk.Label(right, text="Preview / Compare Result", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        tree_frame = tk.Frame(right)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=tree_scroll_y.set)

        tree_scroll_x = ttk.Scrollbar(right, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.pack(fill="x")
        self.tree.configure(xscrollcommand=tree_scroll_x.set)

        bottom = tk.Frame(self.root, bd=1, relief="groove")
        bottom.pack(fill="x", side="bottom")

        left_bottom = tk.Frame(bottom)
        left_bottom.pack(side="left", padx=10, pady=8)

        tk.Button(left_bottom, text="Compare", command=self.compare, width=16).pack(side="left", padx=(0, 10))
        tk.Button(left_bottom, text="Export Result", command=self.export, width=16).pack(side="left")

        self.status = tk.Label(bottom, text="Ready", anchor="w")
        self.status.pack(side="left", padx=20)

        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=250)
        self.progress.pack(side="right", padx=10, pady=8)

    # =========================
    # UI helpers
    # =========================
    def on_scroll_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def clear_keys(self):
        for v in self.column_vars.values():
            v["key"].set(False)
        self.refresh_preview_if_possible()

    def clear_ignore(self):
        for v in self.column_vars.values():
            v["ignore"].set(False)
        self.refresh_preview_if_possible()

    def clear_all_checks(self):
        for v in self.column_vars.values():
            v["key"].set(False)
            v["ignore"].set(False)
        self.refresh_preview_if_possible()

    def update_diagnostics(self):
        if self.file1_path and self.df1 is not None:
            self.diag1.config(
                text=(
                    f"File 1 → name: {os.path.basename(self.file1_path)}, "
                    f"rows: {len(self.df1):,}, cols: {len(self.df1.columns)}, "
                    f"size: {file_size_kb(self.file1_path)} KB"
                )
            )

        if self.file2_path and self.df2 is not None:
            self.diag2.config(
                text=(
                    f"File 2 → name: {os.path.basename(self.file2_path)}, "
                    f"rows: {len(self.df2):,}, cols: {len(self.df2.columns)}, "
                    f"size: {file_size_kb(self.file2_path)} KB"
                )
            )

    # =========================
    # Load files
    # =========================
    def load_file1(self):
        path = filedialog.askopenfilename(
            filetypes=[("Supported files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            self.file1_path = path
            self.df1 = load_file(path)
            self.lbl1.config(text=os.path.basename(path))
            self.update_columns()
            self.update_diagnostics()
            self.refresh_preview_if_possible()
            self.status.config(text="File 1 loaded")

    def load_file2(self):
        path = filedialog.askopenfilename(
            filetypes=[("Supported files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            self.file2_path = path
            self.df2 = load_file(path)
            self.lbl2.config(text=os.path.basename(path))
            self.update_columns()
            self.update_diagnostics()
            self.refresh_preview_if_possible()
            self.status.config(text="File 2 loaded")

    # =========================
    # Columns
    # =========================
    def update_columns(self):
        if self.df1 is None:
            return

        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        old_state = {}
        for col, vals in self.column_vars.items():
            old_state[col] = {
                "key": vals["key"].get(),
                "ignore": vals["ignore"].get()
            }

        self.column_vars.clear()

        if self.df2 is not None:
            cols = list(dict.fromkeys(list(self.df1.columns) + list(self.df2.columns)))
        else:
            cols = list(self.df1.columns)

        header = tk.Frame(self.scroll_frame)
        header.pack(fill="x", pady=(0, 4))
        tk.Label(header, text="Column", width=24, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(header, text="Key", width=8, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(header, text="Ignore", width=8, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")

        for col in cols:
            row = tk.Frame(self.scroll_frame)
            row.pack(fill="x", pady=1)

            tk.Label(row, text=col, width=24, anchor="w").pack(side="left")

            key_var = tk.BooleanVar(value=old_state.get(col, {}).get("key", False))
            ignore_var = tk.BooleanVar(value=old_state.get(col, {}).get("ignore", False))

            key_var.trace_add("write", lambda *args: self.refresh_preview_if_possible())
            ignore_var.trace_add("write", lambda *args: self.refresh_preview_if_possible())

            tk.Checkbutton(row, variable=key_var).pack(side="left", padx=(8, 18))
            tk.Checkbutton(row, variable=ignore_var).pack(side="left")

            self.column_vars[col] = {"key": key_var, "ignore": ignore_var}

    def get_selected_keys_and_ignore(self):
        keys = [c for c, v in self.column_vars.items() if v["key"].get()]
        ignore = [c for c, v in self.column_vars.items() if v["ignore"].get()]
        return keys, ignore

    # =========================
    # Preview
    # =========================
    def get_preview_dataframe(self):
        if self.df1 is None:
            return None

        keys, ignore = self.get_selected_keys_and_ignore()

        df = self.df1.copy()
        df = df.drop(columns=ignore, errors="ignore")

        if self.df2 is not None:
            df2_cols = set(self.df2.columns) - set(ignore)
            common_cols = [c for c in df.columns if c in df2_cols]
        else:
            common_cols = list(df.columns)

        ordered_keys = [k for k in keys if k in common_cols]
        ordered_rest = [c for c in common_cols if c not in ordered_keys]
        final_order = ordered_keys + ordered_rest

        if final_order:
            return df[final_order].copy()
        return df.copy()

    def refresh_preview_if_possible(self):
        preview_df = self.get_preview_dataframe()
        if preview_df is not None:
            self.preview_df(preview_df)

    def preview_df(self, df):
        self.tree.delete(*self.tree.get_children())

        cols = list(df.columns)
        self.tree["columns"] = cols

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w")

        for _, row in df.head(100).iterrows():
            self.tree.insert("", "end", values=list(row))

    # =========================
    # Second pass matching
    # =========================
    def build_loose_signature(self, row, compare_cols):
        vals = []
        for c in compare_cols:
            vals.append(norm_by_column(c, row.get(f"{c}_F1", row.get(f"{c}_F2", None))))
        return tuple(vals)

    def second_pass_reconcile(self, merged_df, compare_cols):
        matched_rows = merged_df[merged_df["status"].isin(["same", "different"])].copy()
        left_only = merged_df[merged_df["status"] == "only_in_file1"].copy()
        right_only = merged_df[merged_df["status"] == "only_in_file2"].copy()

        if left_only.empty or right_only.empty:
            return pd.concat([matched_rows, left_only, right_only], ignore_index=True)

        right_used = set()
        reconciled = []

        right_signatures = {}
        for ridx, r in right_only.iterrows():
            sig = self.build_loose_signature(r, compare_cols)
            right_signatures.setdefault(sig, []).append(ridx)

        for _, lrow in left_only.iterrows():
            lsig = self.build_loose_signature(lrow, compare_cols)

            candidate_idxs = right_signatures.get(lsig, [])
            chosen = None
            for ridx in candidate_idxs:
                if ridx not in right_used:
                    chosen = ridx
                    break

            if chosen is not None:
                rrow = right_only.loc[chosen].copy()
                right_used.add(chosen)

                new_row = lrow.copy()

                for col in right_only.columns:
                    if col.endswith("_F2") or col in ["_merge", "status", "diff_columns"]:
                        new_row[col] = rrow.get(col)

                new_row["_merge"] = "both"
                new_row["status"] = "same"
                new_row["diff_columns"] = ""
                reconciled.append(new_row)
            else:
                reconciled.append(lrow)

        for ridx, rrow in right_only.iterrows():
            if ridx not in right_used:
                reconciled.append(rrow)

        reconciled_df = pd.DataFrame(reconciled)
        final_df = pd.concat([matched_rows, reconciled_df], ignore_index=True)
        return final_df

    # =========================
    # Compare
    # =========================
    def compare(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showerror("Error", "Please load both files first.")
            return

        keys, ignore = self.get_selected_keys_and_ignore()

        if not keys:
            messagebox.showerror("Error", "Please select at least one key.")
            return

        conflict = [k for k in keys if k in ignore]
        if conflict:
            messagebox.showerror(
                "Error",
                f"These columns are selected as both Key and Ignore:\n{conflict}"
            )
            return

        self.progress.start()
        self.status.config(text="Comparing...")
        self.root.update_idletasks()

        try:
            df1 = self.df1.copy()
            df2 = self.df2.copy()

            # 1. drop ignored first
            df1 = df1.drop(columns=ignore, errors="ignore")
            df2 = df2.drop(columns=ignore, errors="ignore")

            # 2. validate keys
            missing_keys_1 = [k for k in keys if k not in df1.columns]
            missing_keys_2 = [k for k in keys if k not in df2.columns]

            if missing_keys_1 or missing_keys_2:
                messagebox.showerror(
                    "Error",
                    f"Some selected keys are missing after ignore step.\n"
                    f"Missing in File 1: {missing_keys_1}\n"
                    f"Missing in File 2: {missing_keys_2}"
                )
                return

            # 3. reorder columns
            common_cols = [c for c in df1.columns if c in df2.columns]
            ordered_keys = [k for k in keys if k in common_cols]
            ordered_rest = [c for c in common_cols if c not in ordered_keys]
            final_order = ordered_keys + ordered_rest

            df1 = df1[final_order].copy()
            df2 = df2[final_order].copy()

            # 4. normalize keys by type
            for k in ordered_keys:
                df1[k] = df1[k].map(lambda x: norm_by_column(k, x))
                df2[k] = df2[k].map(lambda x: norm_by_column(k, x))

            # 5. first pass
            if self.smart_var.get():
                merged = df1.merge(
                    df2,
                    on=ordered_keys,
                    how="outer",
                    indicator=True,
                    suffixes=("_F1", "_F2")
                )
            else:
                df1["pair_idx"] = df1.groupby(ordered_keys).cumcount()
                df2["pair_idx"] = df2.groupby(ordered_keys).cumcount()

                merged = df1.merge(
                    df2,
                    on=ordered_keys + ["pair_idx"],
                    how="outer",
                    indicator=True,
                    suffixes=("_F1", "_F2")
                )

            compare_exclude = ordered_keys + ["pair_idx"]
            compare_cols = [c for c in final_order if c not in compare_exclude]

            status = []
            diff_columns = []

            for _, r in merged.iterrows():
                if r["_merge"] == "left_only":
                    status.append("only_in_file1")
                    diff_columns.append("")
                    continue

                if r["_merge"] == "right_only":
                    status.append("only_in_file2")
                    diff_columns.append("")
                    continue

                diffs = []
                for c in compare_cols:
                    if not values_equal(c, r.get(f"{c}_F1"), r.get(f"{c}_F2")):
                        diffs.append(c)

                if diffs:
                    status.append("different")
                    diff_columns.append(", ".join(diffs))
                else:
                    status.append("same")
                    diff_columns.append("")

            merged["status"] = status
            merged["diff_columns"] = diff_columns

            # 6. second pass
            merged = self.second_pass_reconcile(merged, compare_cols)

            # 7. result columns
            result_cols = []

            for k in ordered_keys:
                if k in merged.columns:
                    result_cols.append(k)

            if "pair_idx" in merged.columns:
                result_cols.append("pair_idx")

            for c in compare_cols:
                left_col = f"{c}_F1"
                right_col = f"{c}_F2"
                if left_col in merged.columns:
                    result_cols.append(left_col)
                if right_col in merged.columns:
                    result_cols.append(right_col)

            result_cols += ["_merge", "status", "diff_columns"]

            existing_cols = [c for c in result_cols if c in merged.columns]
            self.result_df = merged[existing_cols].copy()

            self.preview_df(self.result_df)

            counts = self.result_df["status"].value_counts().to_dict()
            self.status.config(text=f"Done | {counts}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.config(text="Error")
        finally:
            self.progress.stop()

    # =========================
    # Export
    # =========================
    def export(self):
        if self.result_df is None:
            messagebox.showerror("Error", "No compare result to export.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")]
        )
        if not path:
            return

        try:
            self.result_df.to_excel(path, index=False)
            messagebox.showinfo("Saved", "Export successful.")
            self.status.config(text=f"Exported: {path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = CompareApp(root)
    root.mainloop()