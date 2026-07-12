"""
risk_analyzer_gui.py
=====================
A desktop GUI tool that:
  1. Opens an Excel file containing a list of software libraries
     (this is INPUT data - the program does NOT generate this file)
  2. Classifies every library as SAFE or CRITICAL based on 3 simple checks:
        - is it a known-broken version (has a security bug)?
        - does it use a risky license?
        - has it not been updated in 2+ years?
  3. Shows the results on screen in a table
  4. Lets you filter the table to show only Safe or only Critical libraries
  5. Draws a chart of the results, right there in the window

Nothing is saved to disk automatically - everything happens on screen.

EXPECTED EXCEL COLUMNS (case-insensitive, any order):
    App ID | App Name | Library | Version | License |
    Dependency Type | Last Updated | Distributed

Run:
    python3 risk_analyzer_gui.py

(Use make_sample_input.py first if you don't have your own Excel file yet -
it creates "dependency_data.xlsx" with sample data you can open.)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, datetime

import pandas as pd

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

TODAY = date(2026, 7, 11)

# A small reference list of library+version pairs known to have security bugs.
# (In a real company tool, this would be a real vulnerability database.)
KNOWN_BROKEN_VERSIONS = {
    ("log4j-core", "2.14.1"): "CVE-2021-44228",
    ("jackson-databind", "2.9.8"): "CVE-2019-12384",
    ("spring-core", "5.2.0"): "CVE-2022-22965",
    ("commons-lang3", "3.8"): "CVE-2020-15250",
    ("flask", "0.12"): "CVE-2019-1010083",
    ("left-pad", "1.0.0"): "CVE-2018-99999",
    ("openssl-wrapper", "1.1.0"): "CVE-2020-1971",
}

RISKY_LICENSES = {"GPL-3.0", "AGPL-3.0", "Unlicensed"}
SAFE_COLOR = "#27ae60"
CRITICAL_COLOR = "#c0392b"


# =============================================================================
# CLASSIFICATION LOGIC  (reads a row, decides Safe or Critical)
# =============================================================================

def classify_row(row):
    """
    Takes one dependency row (a dict-like) and returns a dict with the
    classification result. Kept deliberately simple: 3 yes/no checks,
    if ANY of them are a problem -> Critical, otherwise -> Safe.
    """
    library = str(row.get("Library", "")).strip()
    version = str(row.get("Version", "")).strip()
    license_name = str(row.get("License", "")).strip()
    distributed = str(row.get("Distributed", "Yes")).strip().lower() in ("yes", "y", "true", "1")

    reasons = []

    # 1) known security bug?
    cve = KNOWN_BROKEN_VERSIONS.get((library, version))
    if cve:
        reasons.append(f"Known vulnerability ({cve})")

    # 2) risky license? (GPL/AGPL only matters if the app is actually distributed)
    if license_name in RISKY_LICENSES:
        if license_name == "Unlicensed" or distributed:
            reasons.append(f"Risky license ({license_name})")

    # 3) not updated in 2+ years?
    last_updated_raw = row.get("Last Updated")
    years_old = None
    if pd.notna(last_updated_raw):
        try:
            if isinstance(last_updated_raw, (datetime, date)):
                last_updated = last_updated_raw.date() if isinstance(last_updated_raw, datetime) else last_updated_raw
            else:
                last_updated = datetime.fromisoformat(str(last_updated_raw)[:10]).date()
            years_old = round((TODAY - last_updated).days / 365, 1)
            if years_old >= 2:
                reasons.append(f"Outdated ({years_old} years since last update)")
        except (ValueError, TypeError):
            pass

    status = "Critical" if reasons else "Safe"
    return {"status": status, "reasons": "; ".join(reasons) if reasons else "No issues found",
            "years_old": years_old}


def load_and_classify(filepath):
    """Reads the Excel file and returns a list of classified dependency dicts."""
    df = pd.read_excel(filepath)
    # normalize column names so "app id", "App_ID", "App ID" all work the same
    df.columns = [str(c).strip() for c in df.columns]

    results = []
    for _, row in df.iterrows():
        classification = classify_row(row)
        results.append({
            "App ID": row.get("App ID", ""),
            "App Name": row.get("App Name", ""),
            "Library": row.get("Library", ""),
            "Version": row.get("Version", ""),
            "License": row.get("License", ""),
            "Type": row.get("Dependency Type", ""),
            "Status": classification["status"],
            "Reason": classification["reasons"],
        })
    return results


# =============================================================================
# GUI APPLICATION
# =============================================================================

class RiskAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Software Dependency Risk Analyzer")
        self.root.geometry("1150x700")

        self.all_results = []      # every row, unfiltered
        self.loaded_filename = tk.StringVar(value="No file loaded")

        self._build_top_bar()
        self._build_filter_bar()
        self._build_main_area()
        self._build_status_bar()

    # ---------------------------------------------------------------
    # UI construction
    # ---------------------------------------------------------------
    def _build_top_bar(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Open Excel File...", command=self.open_file).pack(side="left")
        ttk.Label(top, textvariable=self.loaded_filename, foreground="#555").pack(side="left", padx=12)

    def _build_filter_bar(self):
        bar = ttk.Frame(self.root, padding=(10, 0))
        bar.pack(fill="x")

        ttk.Label(bar, text="View:").pack(side="left")
        self.filter_choice = tk.StringVar(value="All")
        for label in ("All", "Safe", "Critical"):
            ttk.Radiobutton(bar, text=label, value=label, variable=self.filter_choice,
                             command=self.apply_filter).pack(side="left", padx=6)

        ttk.Label(bar, text="   App:").pack(side="left", padx=(20, 0))
        self.app_filter_choice = tk.StringVar(value="All Apps")
        self.app_dropdown = ttk.Combobox(bar, textvariable=self.app_filter_choice,
                                          values=["All Apps"], state="readonly", width=20)
        self.app_dropdown.pack(side="left", padx=6)
        self.app_dropdown.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

    def _build_main_area(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        # left: table of results
        table_frame = ttk.Frame(main)
        table_frame.pack(side="left", fill="both", expand=True)

        columns = ("App ID", "App Name", "Library", "Version", "License", "Type", "Status", "Reason")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=25)
        for col in columns:
            self.tree.heading(col, text=col)
            width = 220 if col == "Reason" else 100
            self.tree.column(col, width=width, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.tree.tag_configure("Safe", background="#eafaf1")
        self.tree.tag_configure("Critical", background="#fdecea")

        # right: chart
        chart_frame = ttk.Frame(main, width=380)
        chart_frame.pack(side="right", fill="y", padx=(10, 0))
        chart_frame.pack_propagate(False)

        self.figure = plt.Figure(figsize=(4, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self._draw_empty_chart()

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="Load an Excel file to begin.")
        status = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", padding=6)
        status.pack(fill="x", side="bottom")

    # ---------------------------------------------------------------
    # Actions
    # ---------------------------------------------------------------
    def open_file(self):
        filepath = filedialog.askopenfilename(
            title="Select dependency Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not filepath:
            return

        try:
            self.all_results = load_and_classify(filepath)
        except Exception as exc:
            messagebox.showerror("Could not read file", str(exc))
            return

        self.loaded_filename.set(filepath.split("/")[-1])

        app_names = sorted({str(r["App Name"]) for r in self.all_results if r["App Name"]})
        self.app_dropdown["values"] = ["All Apps"] + app_names
        self.app_filter_choice.set("All Apps")
        self.filter_choice.set("All")

        self.apply_filter()

    def apply_filter(self):
        if not self.all_results:
            return

        status_filter = self.filter_choice.get()
        app_filter = self.app_filter_choice.get()

        filtered = self.all_results
        if status_filter != "All":
            filtered = [r for r in filtered if r["Status"] == status_filter]
        if app_filter != "All Apps":
            filtered = [r for r in filtered if str(r["App Name"]) == app_filter]

        # refresh table
        self.tree.delete(*self.tree.get_children())
        for r in filtered:
            values = (r["App ID"], r["App Name"], r["Library"], r["Version"],
                      r["License"], r["Type"], r["Status"], r["Reason"])
            self.tree.insert("", "end", values=values, tags=(r["Status"],))

        self._update_chart(filtered)
        self._update_status(filtered)

    def _update_status(self, filtered):
        total = len(self.all_results)
        safe_total = sum(1 for r in self.all_results if r["Status"] == "Safe")
        critical_total = sum(1 for r in self.all_results if r["Status"] == "Critical")
        self.status_var.set(
            f"Showing {len(filtered)} of {total} libraries   |   "
            f"Overall: {safe_total} Safe, {critical_total} Critical"
        )

    # ---------------------------------------------------------------
    # Chart
    # ---------------------------------------------------------------
    def _draw_empty_chart(self):
        self.ax.clear()
        self.ax.text(0.5, 0.5, "Load an Excel file\nto see the chart",
                      ha="center", va="center", fontsize=11, color="#888")
        self.ax.axis("off")
        self.canvas.draw()

    def _update_chart(self, filtered):
        self.ax.clear()

        if self.app_filter_choice.get() == "All Apps":
            # Chart 1: overall Safe vs Critical counts (based on currently visible status filter)
            safe_count = sum(1 for r in filtered if r["Status"] == "Safe")
            critical_count = sum(1 for r in filtered if r["Status"] == "Critical")
            labels, counts, colors = [], [], []
            if self.filter_choice.get() in ("All", "Safe"):
                labels.append("Safe"); counts.append(safe_count); colors.append(SAFE_COLOR)
            if self.filter_choice.get() in ("All", "Critical"):
                labels.append("Critical"); counts.append(critical_count); colors.append(CRITICAL_COLOR)

            self.ax.bar(labels, counts, color=colors)
            self.ax.set_title("Libraries by Classification", fontsize=10)
            self.ax.set_ylabel("Count")
            for i, c in enumerate(counts):
                self.ax.text(i, c, str(c), ha="center", va="bottom", fontsize=9)
        else:
            # a specific app is selected -> show Safe vs Critical just for that app
            safe_count = sum(1 for r in filtered if r["Status"] == "Safe")
            critical_count = sum(1 for r in filtered if r["Status"] == "Critical")
            self.ax.pie([safe_count, critical_count] if (safe_count + critical_count) else [1],
                        labels=["Safe", "Critical"] if (safe_count + critical_count) else ["No data"],
                        colors=[SAFE_COLOR, CRITICAL_COLOR] if (safe_count + critical_count) else ["#ccc"],
                        autopct=lambda p: f"{p:.0f}%" if p > 0 else "",
                        startangle=90)
            self.ax.set_title(f"{self.app_filter_choice.get()}: Safe vs Critical", fontsize=10)

        self.figure.tight_layout()
        self.canvas.draw()


# =============================================================================
# MAIN
# =============================================================================

def main():
    root = tk.Tk()
    app = RiskAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
