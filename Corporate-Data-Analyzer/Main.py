import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

import ttkbootstrap as tb
from tkinter import ttk


class CorporateDataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analyzer - Excel & CSV Analytics Tool")
        self.root.geometry("1180x880")
        self.root.minsize(1100, 750)
        self.root.configure(bg="#F8FAFC")

        # -------------------- State --------------------
        self.file_path = tk.StringVar()
        self.group_by_var = tk.StringVar()
        self.agg_var = tk.StringVar(value="Sum")
        self.value_var = tk.StringVar()
        self.export_format_var = tk.StringVar(value="Excel (.xlsx)")
        self.chart_type_var = tk.StringVar(value="Column")
        self.top_n_var = tk.StringVar(value="10")

        self.df = None
        self.report_df = None
        self.current_figure = None

        self.text_columns = []
        self.numeric_columns = []

        # UI Components State
        self.report_tree = None
        self.report_status_var = None

        self.agg_map = {
            "Sum": "sum",
            "Mean": "mean",
            "Average": "mean",
            "Max": "max",
            "Min": "min",
            "Count": "count",
            "Median": "median",
        }

        self.chart_types = ["Bar", "Column", "Line", "Pie"]
        self.export_formats = ["Excel (.xlsx)", "CSV (.csv)"]

        self._configure_styles()
        self._build_ui()

    # -------------------- Styling --------------------
    def _configure_styles(self):
        style = tb.Style("flatly")

        bg = "#F8FAFC"
        text = "#0F172A"
        muted = "#475569"
        accent = "#4F46E5"

        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, font=("Segoe UI", 10), foreground=text)
        style.configure("Title.TLabel", background=bg, font=("Segoe UI", 22, "bold"), foreground=text)
        style.configure("Header.TLabel", background=bg, font=("Segoe UI", 11, "bold"), foreground=muted)
        style.configure("Section.TLabel", background=bg, font=("Segoe UI", 10, "bold"), foreground=muted)

        style.configure("TLabelframe", background=bg, foreground=text, font=("Segoe UI", 10, "bold"))
        style.configure("TLabelframe.Label", background=bg, foreground=text, font=("Segoe UI", 10, "bold"))

        style.configure("TButton", font=("Segoe UI", 9, "bold"), padding=(12, 8))
        style.configure("TCombobox", padding=4)
        style.configure("TEntry", padding=5)

        style.configure(
            "Treeview",
            rowheight=30,
            font=("Segoe UI", 9),
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#E2E8F0",
            foreground=text,
        )

        self.metric_color = accent

    # -------------------- UI --------------------
    def _build_ui(self):
        self._build_header()
        self._build_file_section()
        self._build_info_section()
        self._build_controls_section()
        
        # Draw the bottom bar FIRST so it locks to the bottom edge
        self._build_status_bar() 
        
        # Now draw the expanding table so it fills the remaining gap
        self._build_embedded_report_section()

    def _build_header(self):
        header = ttk.Frame(self.root, padding=(16, 14, 16, 6))
        header.pack(fill="x")

        ttk.Label(
            header,
            text="Corporate Data Analyzer - Excel & CSV Analytics Tool",
            style="Title.TLabel"
        ).pack(anchor="center")

        ttk.Label(
            header,
            text="Clean, analyze, group, and visualize Excel/CSV data in a polished desktop dashboard.",
            style="Header.TLabel"
        ).pack(anchor="center", pady=(4, 0))

    def _build_file_section(self):
        frame = ttk.LabelFrame(self.root, text="File Selection & Reading", padding=10)
        frame.pack(fill="x", padx=14, pady=(0, 10))

        ttk.Label(frame, text="Select CSV/Excel:").grid(row=0, column=0, sticky="w", padx=(4, 8), pady=4)
        self.file_entry = ttk.Entry(frame, textvariable=self.file_path, width=90)
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(frame, text="Browse", command=self.browse_file, bootstyle="primary").grid(
            row=0, column=2, padx=8, pady=4
        )
        ttk.Button(frame, text="Read", command=self.read_file, bootstyle="success").grid(
            row=0, column=3, padx=(0, 4), pady=4
        )
        ttk.Button(frame, text="Reset Paths", command=self.reset_paths, bootstyle="secondary").grid(
            row=0, column=4, padx=(8, 4), pady=4
        )

        frame.columnconfigure(1, weight=1)

    def _build_info_section(self):
        frame = ttk.LabelFrame(self.root, text="File Info", padding=10)
        frame.pack(fill="x", padx=14, pady=(0, 10))

        cards = ttk.Frame(frame)
        cards.pack(fill="x", pady=(0, 8))

        self.rows_card = self._create_metric_card(cards, "Rows", "-")
        self.cols_card = self._create_metric_card(cards, "Columns", "-")
        self.numeric_card = self._create_metric_card(cards, "Numeric Columns", "-")
        self.text_card = self._create_metric_card(cards, "Text Columns", "-")

        self.rows_card.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.cols_card.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.numeric_card.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.text_card.pack(side="left", fill="x", expand=True)

        headings_box = ttk.Frame(frame)
        headings_box.pack(fill="x", pady=(3, 0))

        ttk.Label(headings_box, text="Column Headings", style="Section.TLabel").pack(anchor="w", pady=(0, 5))

        text_container = ttk.Frame(headings_box)
        text_container.pack(fill="x")

        self.info_text = tk.Text(
            text_container,
            height=3,
            wrap="word",
            font=("Consolas", 10),
            relief="solid",
            borderwidth=1,
            background="#FFFFFF",
            foreground="#0F172A",
        )
        self.info_text.pack(side="left", fill="x", expand=True)

        info_scroll = ttk.Scrollbar(text_container, orient="vertical", command=self.info_text.yview)
        info_scroll.pack(side="right", fill="y")
        self.info_text.configure(yscrollcommand=info_scroll.set)

        self._set_info_text("Select a CSV or Excel file and click Read.")

    def _create_metric_card(self, parent, title, value):
        card = ttk.LabelFrame(parent, text=title, padding=12)
        label = ttk.Label(card, text=value, font=("Segoe UI", 18, "bold"), foreground=self.metric_color)
        label.pack()
        card.metric_label = label
        return card

    def _build_controls_section(self):
        frame = ttk.LabelFrame(self.root, text="Build Report (GroupBy + Aggregation)", padding=10)
        frame.pack(fill="x", padx=14, pady=(0, 10))

        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(4, 8))

        ttk.Label(row1, text="Group By (Text column):").pack(side="left", padx=(4, 8))
        self.group_by_cb = ttk.Combobox(row1, textvariable=self.group_by_var, state="readonly", width=28)
        self.group_by_cb.pack(side="left", padx=(0, 16))

        ttk.Label(row1, text="Aggregation:").pack(side="left", padx=(0, 8))
        self.agg_cb = ttk.Combobox(
            row1,
            textvariable=self.agg_var,
            values=list(self.agg_map.keys()),
            state="readonly",
            width=14
        )
        self.agg_cb.pack(side="left", padx=(0, 16))

        ttk.Label(row1, text="Value (Numeric column):").pack(side="left", padx=(0, 8))
        self.value_cb = ttk.Combobox(row1, textvariable=self.value_var, state="readonly", width=28)
        self.value_cb.pack(side="left")

        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(0, 4))

        ttk.Button(row2, text="Preview Report", command=self.preview_report, bootstyle="primary").pack(
            side="left", padx=(4, 8)
        )

        ttk.Label(row2, text="Export as:").pack(side="left", padx=(8, 6))
        self.export_cb = ttk.Combobox(
            row2,
            textvariable=self.export_format_var,
            values=self.export_formats,
            state="readonly",
            width=14
        )
        self.export_cb.pack(side="left", padx=(0, 8))

        ttk.Button(row2, text="Export Report", command=self.export_report, bootstyle="success").pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(row2, text="Reset Report", command=self.clear_report, bootstyle="secondary").pack(
            side="left", padx=(0, 8)
        )

        ttk.Label(row2, text="Top N:").pack(side="left", padx=(10, 6))
        self.top_n_spin = ttk.Spinbox(row2, from_=3, to=100, textvariable=self.top_n_var, width=6)
        self.top_n_spin.pack(side="left", padx=(0, 8))

        ttk.Label(row2, text="Chart Type:").pack(side="left", padx=(10, 6))
        self.chart_cb = ttk.Combobox(
            row2,
            textvariable=self.chart_type_var,
            values=self.chart_types,
            state="readonly",
            width=12
        )
        self.chart_cb.pack(side="left", padx=(0, 8))

        ttk.Button(row2, text="Preview Chart", command=self.preview_chart, bootstyle="info").pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(row2, text="Export Chart (PNG)", command=self.export_chart, bootstyle="warning").pack(
            side="left", padx=(0, 8)
        )

    def _build_embedded_report_section(self):
        """Builds a clean, single report table at the bottom of the main window."""
        preview_frame = ttk.LabelFrame(self.root, text="Analytical Report Preview", padding=10)
        preview_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        header = ttk.Frame(preview_frame)
        header.pack(fill="x", pady=(0, 5))

        self.report_status_var = tk.StringVar(value="Generate a report to preview it here.")
        ttk.Label(header, textvariable=self.report_status_var, style="Section.TLabel").pack(anchor="w")

        table_frame = ttk.Frame(preview_frame)
        table_frame.pack(fill="both", expand=True)

        self.report_tree = ttk.Treeview(table_frame, show="headings")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.report_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.report_tree.xview)
        self.report_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.report_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def _build_status_bar(self):
        # Create a container frame for the bottom bar and GLUE it to the bottom
        status_frame = ttk.Frame(self.root, padding=(12, 6))
        status_frame.pack(side="bottom", fill="x")

        # Add a subtle line separator right above the status bar for a clean look
        sep = ttk.Separator(self.root, orient="horizontal")
        sep.pack(side="bottom", fill="x")

        # 1. Dynamic status text on the Left
        self.status_var = tk.StringVar(value="Ready.")
        status_lbl = ttk.Label(status_frame, textvariable=self.status_var, anchor="w")
        status_lbl.pack(side="left")

        # 2. Developer credit on the Right
        dev_lbl = ttk.Label(
            status_frame, 
            text="Developed by Himanshu Kumar - Netaji Subhas University Of Technology", 
            font=("Segoe UI", 10, "italic"), 
            foreground="#64748b"  
        )
        dev_lbl.pack(side="right")

    # -------------------- Helpers --------------------
    def set_status(self, text):
        self.status_var.set(text)
        self.root.update_idletasks()

    def _set_info_text(self, text):
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("1.0", text)
        self.info_text.configure(state="disabled")

    def _update_metric_card(self, card, value):
        card.metric_label.config(text=value)

    def _input_folder(self):
        if not self.file_path.get().strip():
            return ""
        return os.path.dirname(os.path.abspath(self.file_path.get().strip()))

    def _safe_numeric_convert(self, series: pd.Series) -> pd.Series:
        cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
        return pd.to_numeric(cleaned, errors="coerce")

    def _get_top_n(self):
        try:
            n = int(self.top_n_var.get())
            return max(3, min(n, 1000))
        except Exception:
            return 10

    def _ensure_file_loaded(self):
        if self.df is None:
            messagebox.showerror("No Data", "Please click Read after selecting a file.")
            return False
        return True

    def _ensure_report_ready(self):
        if self.report_df is None or self.report_df.empty:
            messagebox.showerror("No Report", "Please generate a report first using Preview Report.")
            return False
        return True

    # -------------------- File actions --------------------
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Supported", "*.xlsx *.xls *.csv"),
            ]
        )
        if path:
            self.file_path.set(path)
            self.df = None
            self.report_df = None
            self.text_columns = []
            self.numeric_columns = []
            self.group_by_var.set("")
            self.value_var.set("")
            self.clear_report()
            self._set_info_text("Select a CSV or Excel file and click Read.")
            self._update_metric_card(self.rows_card, "-")
            self._update_metric_card(self.cols_card, "-")
            self._update_metric_card(self.numeric_card, "-")
            self._update_metric_card(self.text_card, "-")
            self.set_status("File selected. Click Read to load data.")

    def reset_paths(self):
        self.file_path.set("")
        self.df = None
        self.report_df = None
        self.text_columns = []
        self.numeric_columns = []
        self.group_by_var.set("")
        self.value_var.set("")
        self._set_info_text("Select a CSV or Excel file and click Read.")
        self._update_metric_card(self.rows_card, "-")
        self._update_metric_card(self.cols_card, "-")
        self._update_metric_card(self.numeric_card, "-")
        self._update_metric_card(self.text_card, "-")
        self.clear_report()
        self.set_status("Paths reset.")

    def read_file(self):
        path = self.file_path.get().strip()

        if not path:
            messagebox.showerror("Missing File", "Please select a CSV or Excel file first.")
            return

        if not os.path.exists(path):
            messagebox.showerror("Invalid File", "The selected file does not exist.")
            return

        ext = os.path.splitext(path)[1].lower()

        try:
            if ext == ".csv":
                try:
                    df = pd.read_csv(path)
                except Exception:
                    df = None
                    for enc in ("utf-8-sig", "latin1", "cp1252"):
                        try:
                            df = pd.read_csv(path, encoding=enc)
                            break
                        except Exception:
                            continue
                    if df is None:
                        raise
            elif ext in (".xlsx", ".xls"):
                df = pd.read_excel(path, sheet_name=0)
            else:
                messagebox.showerror("Unsupported File", "Please select a CSV or Excel file.")
                return
        except Exception as e:
            messagebox.showerror("Read Error", f"Could not read the file.\n\n{e}")
            return

        if df is None or df.empty:
            messagebox.showwarning("Empty File", "The selected file contains no data.")
            return

        self.df = df.copy()
        self.report_df = None
        self.clear_report()

        self.text_columns, self.numeric_columns = self._detect_columns(self.df)
        self._update_dropdowns()
        self._update_file_info()

        self.set_status(f"Loaded: {os.path.basename(path)}")
        messagebox.showinfo("Success", "File loaded successfully.")

    def _detect_columns(self, df):
        text_cols = []
        numeric_cols = []

        for col in df.columns:
            s = df[col]

            if pd.api.types.is_numeric_dtype(s):
                numeric_cols.append(col)
                continue

            if pd.api.types.is_datetime64_any_dtype(s):
                text_cols.append(col)
                continue

            non_null = s.dropna().astype(str).str.strip()
            if non_null.empty:
                text_cols.append(col)
                continue

            coerced = pd.to_numeric(non_null, errors="coerce")
            numeric_ratio = coerced.notna().mean()

            if numeric_ratio >= 0.80:
                numeric_cols.append(col)
            else:
                text_cols.append(col)

        return text_cols, numeric_cols

    def _update_dropdowns(self):
        self.group_by_cb["values"] = self.text_columns
        self.value_cb["values"] = self.numeric_columns

        self.group_by_var.set(self.text_columns[0] if self.text_columns else "")
        self.value_var.set(self.numeric_columns[0] if self.numeric_columns else "")
        self.agg_var.set("Sum")
        self.chart_type_var.set("Column")
        self.export_format_var.set("Excel (.xlsx)")
        self.top_n_var.set("10")

    def _update_file_info(self):
        rows, cols = self.df.shape
        headings = list(map(str, self.df.columns))

        info_text = (
            f"Rows: {rows}\n"
            f"Columns: {cols}\n"
            f"Text Columns: {len(self.text_columns)}\n"
            f"Numeric Columns: {len(self.numeric_columns)}\n\n"
            f"Column Headings:\n- " + "\n- ".join(headings)
        )
        self._set_info_text(info_text)
        self._update_metric_card(self.rows_card, str(rows))
        self._update_metric_card(self.cols_card, str(cols))
        self._update_metric_card(self.numeric_card, str(len(self.numeric_columns)))
        self._update_metric_card(self.text_card, str(len(self.text_columns)))

    # -------------------- Report generation --------------------
    def preview_report(self):
        if not self._ensure_file_loaded():
            return

        group_col = self.group_by_var.get().strip()
        agg_name = self.agg_var.get().strip()
        value_col = self.value_var.get().strip()

        if not group_col:
            messagebox.showerror("Missing Selection", "Please select a Group By column.")
            return
        if not agg_name:
            messagebox.showerror("Missing Selection", "Please select an Aggregation method.")
            return
        if not value_col:
            messagebox.showerror("Missing Selection", "Please select a Value column.")
            return

        if group_col not in self.text_columns:
            messagebox.showerror("Invalid Selection", "Group By column must be a text column.")
            return
        if value_col not in self.numeric_columns:
            messagebox.showerror("Invalid Selection", "Value column must be a numeric column.")
            return

        try:
            working_df = self.df.copy()

            if not pd.api.types.is_numeric_dtype(working_df[value_col]):
                working_df[value_col] = self._safe_numeric_convert(working_df[value_col])

            # Apply Data Cleaning to Text Categories before grouping
            working_df[group_col] = working_df[group_col].astype(str).str.strip().str.title()

            agg_func = self.agg_map.get(agg_name)
            if agg_func is None:
                messagebox.showerror("Invalid Aggregation", "Selected aggregation is not supported.")
                return

            report = (
                working_df.groupby(group_col, dropna=False, as_index=False)[value_col]
                .agg(agg_func)
            )

            if report.empty:
                messagebox.showwarning("No Results", "The report returned no data.")
                return

            report = report.sort_values(by=value_col, ascending=False, kind="mergesort").reset_index(drop=True)
            self.report_df = report

            self.render_report_tab()
            self.set_status("Report created successfully.")
        except Exception as e:
            messagebox.showerror("Report Error", f"Could not generate report.\n\n{e}")

    def render_report_tab(self):
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

        cols = list(self.report_df.columns)
        self.report_tree["columns"] = cols

        for col in cols:
            self.report_tree.heading(col, text=str(col))
            self.report_tree.column(col, width=280, anchor="w")

        self.report_tree.tag_configure("odd", background="#FFFFFF")
        self.report_tree.tag_configure("even", background="#F8FAFC")

        for idx, (_, row) in enumerate(self.report_df.iterrows()):
            values = []
            for v in row.tolist():
                if pd.isna(v):
                    values.append("")
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    values.append(f"{v:,.2f}")
                else:
                    values.append(str(v))
            tag = "even" if idx % 2 == 0 else "odd"
            self.report_tree.insert("", "end", values=values, tags=(tag,))

        row_count = len(self.report_df)
        self.report_status_var.set(f"Showing report with {row_count} rows.")

    def clear_report(self):
        self.report_df = None
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
        self.report_status_var.set("Generate a report to preview it here.")

    # -------------------- Exports --------------------
    def export_report(self):
        if not self._ensure_report_ready():
            return
        if not self._ensure_file_loaded():
            return

        input_dir = self._input_folder()
        if not input_dir:
            messagebox.showerror("Export Error", "Could not determine input folder.")
            return

        base_name = os.path.splitext(os.path.basename(self.file_path.get().strip()))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fmt = self.export_format_var.get()

        try:
            if fmt == "Excel (.xlsx)":
                out_path = os.path.join(input_dir, f"Report_{base_name}_{ts}.xlsx")
                self.report_df.to_excel(out_path, index=False, engine="openpyxl")
            elif fmt == "CSV (.csv)":
                out_path = os.path.join(input_dir, f"Report_{base_name}_{ts}.csv")
                self.report_df.to_csv(out_path, index=False)
            else:
                messagebox.showerror("Invalid Format", "Please select a valid export format.")
                return

            messagebox.showinfo("Export Successful", f"Report saved successfully.\n\n{out_path}")
            self.set_status("Report exported successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export report.\n\n{e}")

    # -------------------- Chart pop-up window --------------------
    def preview_chart(self):
        if not self._ensure_report_ready():
            return

        chart_type = self.chart_type_var.get().strip()
        if chart_type not in self.chart_types:
            messagebox.showerror("Invalid Selection", "Please select a valid chart type.")
            return

        try:
            self._draw_chart(chart_type)
            self.set_status("Chart generated successfully in a new window.")
        except Exception as e:
            messagebox.showerror("Chart Error", f"Could not generate chart.\n\n{e}")

    def _draw_chart(self, chart_type):
        # Create a large standalone window for the chart
        chart_win = tk.Toplevel(self.root)
        chart_win.title(f"{chart_type} Chart Visualizer")
        chart_win.geometry("950x650")
        chart_win.configure(bg="#ffffff")

        x_col = self.report_df.columns[0]
        y_col = self.report_df.columns[1]

        top_n = self._get_top_n()
        plot_df = self.report_df.head(top_n).copy()

        labels = plot_df[x_col].astype(str).tolist()
        values = pd.to_numeric(plot_df[y_col], errors="coerce").fillna(0).tolist()

        fig = Figure(figsize=(9, 6), dpi=100)
        ax = fig.add_subplot(111)

        if chart_type == "Bar":
            ax.barh(labels[::-1], values[::-1])
            ax.set_xlabel(y_col)
            ax.set_ylabel(x_col)
            ax.invert_yaxis()
        elif chart_type == "Column":
            ax.bar(labels, values)
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=40, ha="right")
        elif chart_type == "Line":
            ax.plot(labels, values, marker="o", linewidth=2)
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=40, ha="right")
        elif chart_type == "Pie":
            if len(labels) > 8:
                labels = labels[:8]
                values = values[:8]
            ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
            ax.axis("equal")
        else:
            raise ValueError("Unsupported chart type.")

        ax.set_title(f"{chart_type} Chart: {y_col} by {x_col}", fontsize=14, fontweight="bold", pad=20)
        
        # Because it is a large standalone window, tight_layout works flawlessly here
        fig.tight_layout()

        self.current_figure = fig

        canvas = FigureCanvasTkAgg(fig, master=chart_win)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(fill="both", expand=True, padx=10, pady=10)

        toolbar = NavigationToolbar2Tk(canvas, chart_win)
        toolbar.update()

    def export_chart(self):
        if not self._ensure_report_ready():
            return

        if self.current_figure is None:
            messagebox.showerror("No Chart", "Please click Preview Chart before exporting.")
            return

        if not self._ensure_file_loaded():
            return

        input_dir = self._input_folder()
        if not input_dir:
            messagebox.showerror("Export Error", "Could not determine input folder.")
            return

        base_name = os.path.splitext(os.path.basename(self.file_path.get().strip()))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(input_dir, f"Chart_{base_name}_{ts}.png")

        try:
            self.current_figure.savefig(out_path, dpi=200, bbox_inches="tight")
            messagebox.showinfo("Export Successful", f"Chart saved successfully.\n\n{out_path}")
            self.set_status("Chart exported successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export chart.\n\n{e}")


if __name__ == "__main__":
    root = tb.Window(themename="flatly")
    app = CorporateDataAnalyzerApp(root)
    root.mainloop()