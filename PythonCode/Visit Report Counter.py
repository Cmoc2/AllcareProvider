# -*- coding: utf-8 -*-
"""
Created on Tue Dec 30 14:53:14 2025

@author: ChristianOrtiz

Visit Counts by User Type (GUI)
- Asks for MR#, SOC Date, DC Date, and Source File in a single popup window.
- Validates inputs (MR# numeric, dates mm/dd/yyyy, DC >= SOC).
- Reads CSV or Excel and counts visits between SOC and DC inclusive.
- Groups counts by 'User Type' and displays results in a popup.

Columns expected (at minimum): 'MR#', 'User Type', and a date column (default: 'Form Date').
You can change DATE_COLUMN_NAME below if you want to use 'Day of Visit' instead.

Dependencies:
- pandas
- openpyxl (for .xlsx)
- xlrd (for .xls)
- tkinter (bundled with Python)

Run:
    python visit_counts_by_user_type.py
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import pandas as pd

# --------------------- Configuration ---------------------
# Prefer to add your source file here if you don't want to pick it each time:
SOURCE_FILE_DEFAULT = ""  # e.g., r"C:\Path\to\your\file.xlsx"

# Choose which column represents the visit date in your dataset:
# Common options based on your headers: 'Form Date' or 'Day of Visit'
DATE_COLUMN_NAME = "Form Date"  # Change to "Day of Visit" if that's the actual visit date.
MRN_COLUMN_NAME = "MR#"
USER_TYPE_COLUMN_NAME = "User Type"

# ---------------------------------------------------------

def parse_mmddyyyy(date_str: str, field_label: str) -> datetime:
    """Parse a date string in mm/dd/yyyy format with validation and helpful errors."""
    try:
        return datetime.strptime(date_str.strip(), "%m/%d/%Y")
    except ValueError:
        raise ValueError(f"{field_label} must be in mm/dd/yyyy format. You entered: {date_str!r}")

def parse_optional_mmddyyyy(date_str: str, default_dt: datetime, field_label: str) -> datetime:
    """
    Parse a date string in mm/dd/yyyy format if provided; otherwise return default_dt.
    Raises ValueError if a non-empty string is invalid format.
    """
    s = date_str.strip()
    if s == "":
        return default_dt
    try:
        return datetime.strptime(s, "%m/%d/%Y")
    except ValueError:
        raise ValueError(f"{field_label} must be in mm/dd/yyyy format or left blank. You entered: {date_str!r}")


def require_columns(df: pd.DataFrame, needed_cols):
    """Ensure required columns exist; raise a user-friendly error otherwise."""
    missing = [c for c in needed_cols if c not in df.columns]
    if missing:
        raise KeyError(
            "Missing required column(s): "
            + ", ".join(missing)
            + "\n\nColumns found:\n- " + "\n- ".join(df.columns)
        )

def read_source_file(path: str) -> pd.DataFrame:
    """Read CSV or Excel into a DataFrame, choosing appropriate engine."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Source file not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".csv":
            return pd.read_csv(path)
        elif ext == ".xlsx":
            return pd.read_excel(path, sheet_name="Visit Report Data", engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(path, engine="xlrd")
        else:
            raise ValueError("Unsupported file type. Please use .csv, .xlsx, or .xls")
    except Exception as e:
        raise RuntimeError(f"Failed to read the source file:\n{e}")

def count_visits_by_user_type(df: pd.DataFrame, mrn: int, soc: datetime, dc: datetime) -> pd.Series:
    """
    Filter dataframe for MRN and date range, then group by 'User Type' and count.
    Dates are inclusive: [SOC, DC].
    """
    # Validate required columns first
    require_columns(df, [MRN_COLUMN_NAME, USER_TYPE_COLUMN_NAME, DATE_COLUMN_NAME])

    # Coerce MR# to numeric for robust matching
    mrn_series = pd.to_numeric(df[MRN_COLUMN_NAME], errors="coerce")

    # Parse the date column to datetime
    date_series = pd.to_datetime(df[DATE_COLUMN_NAME], errors="coerce")

    # Filter by MRN match and valid dates within range
    mask = (mrn_series == mrn) & (date_series.notna()) & (date_series >= soc) & (date_series <= dc)
    filtered = df.loc[mask]

    # Group by User Type and count
    # Normalize/handle missing user type as 'Unknown'
    user_type = filtered[USER_TYPE_COLUMN_NAME].fillna("Unknown").astype(str).str.strip()
    counts = user_type.value_counts().sort_index()

    return counts

class VisitCounterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Visit Counter by User Type")

        # Prevent scaling issues on high DPI
        try:
            self.call('tk', 'scaling', 1.2)
        except Exception:
            pass

        self.resizable(False, False)

        # Variables
        self.var_mrn = tk.StringVar(value="")
        self.var_soc = tk.StringVar(value="")
        self.var_dc = tk.StringVar(value="")
        self.var_source = tk.StringVar(value=SOURCE_FILE_DEFAULT)

        # Build UI
        self._build_form()
        
        # Cache
        self.df_cache = None
        self.df_cache_source = None

    def _build_form(self):
        pad = 8

        frm = ttk.Frame(self, padding=16)
        frm.grid(row=0, column=0, sticky="nsew")

        # MR#
        ttk.Label(frm, text="MR# (numeric):").grid(row=0, column=0, sticky="w", padx=pad, pady=(pad, 2))
        ttk.Entry(frm, textvariable=self.var_mrn, width=30).grid(row=0, column=1, sticky="w", padx=pad, pady=(pad, 2))

        # SOC Date
        ttk.Label(frm, text="SOC Date (mm/dd/yyyy):").grid(row=1, column=0, sticky="w", padx=pad, pady=2)
        ttk.Entry(frm, textvariable=self.var_soc, width=30).grid(row=1, column=1, sticky="w", padx=pad, pady=2)

        # DC Date
        ttk.Label(frm, text="DC Date (mm/dd/yyyy) - optional:").grid(row=2, column=0, sticky="w", padx=pad, pady=2)
        ttk.Entry(frm, textvariable=self.var_dc, width=30).grid(row=2, column=1, sticky="w", padx=pad, pady=2)

        # Source file
        ttk.Label(frm, text="Source file (.csv/.xlsx/.xls):").grid(row=3, column=0, sticky="w", padx=pad, pady=2)
        source_entry = ttk.Entry(frm, textvariable=self.var_source, width=48)
        source_entry.grid(row=3, column=1, sticky="w", padx=pad, pady=2)
        ttk.Button(frm, text="Browse…", command=self._browse_file).grid(row=3, column=2, sticky="w", padx=pad, pady=2)

        # Info line about date column in use
        info_text = f"Using date column: '{DATE_COLUMN_NAME}'. Change in script if needed."
        ttk.Label(frm, text=info_text, foreground="#555").grid(row=4, column=0, columnspan=3, sticky="w", padx=pad, pady=(2, pad))

        # Actions
        ttk.Button(frm, text="Run", command=self._run).grid(row=5, column=1, sticky="e", padx=pad, pady=(2, pad))
        ttk.Button(frm, text="Quit", command=self.destroy).grid(row=5, column=2, sticky="w", padx=pad, pady=(2, pad))

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Source File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.var_source.set(path)

    def _run(self):
        # Validate MRN
        mrn_str = self.var_mrn.get().strip()
        if not mrn_str.isdigit():
            messagebox.showerror("Invalid MR#", "MR# must be numeric (digits only).")
            return
        mrn = int(mrn_str)

        # Validate dates
        try:
            soc = parse_mmddyyyy(self.var_soc.get(), "SOC Date")
        # If DC is blank, use today's date
            today_local = datetime.today()
            dc = parse_optional_mmddyyyy(self.var_dc.get(), default_dt=today_local, field_label="DC Date")
        except ValueError as e:
            messagebox.showerror("Invalid Date", str(e))
            return


        if dc < soc:
            messagebox.showerror("Invalid Range", "DC Date must be the same or after SOC Date.")
            return

        # Validate source file
        source_path = self.var_source.get().strip()
        if not source_path:
            messagebox.showerror("Missing Source File", "Please provide or browse for the source file.")
            return

        # Read and process (with caching)
        try:
            # Reuse cache if same source path and df already loaded
            if self.df_cache is not None and self.df_cache_source == source_path:
                df = self.df_cache
            else:
                df = read_source_file(source_path)
                # Cache for subsequent runs in this session
                self.df_cache = df
                self.df_cache_source = source_path
        
            counts = count_visits_by_user_type(df, mrn, soc, dc)
        except (FileNotFoundError, KeyError, ValueError, RuntimeError) as e:
            messagebox.showerror("Error", str(e))
            return
        except Exception as e:
            messagebox.showerror("Unexpected Error", f"{e.__class__.__name__}: {e}")
            return


        total = int(counts.sum()) if not counts.empty else 0

        # Format results
        lines = []
        if total == 0:
            lines.append("No visits found for the specified MR# and date range.")
        else:
            for user_type, cnt in counts.items():
                lines.append(f"{user_type}: {cnt}")

        date_range_str = f"{soc.strftime('%m/%d/%Y')} – {dc.strftime('%m/%d/%Y')}"
        dc_entered = self.var_dc.get().strip() != ""
        dc_note = "" if dc_entered else " (DC defaulted to today)"
        msg = (
            f"MR#: {mrn}\n"
            f"Date column used: '{DATE_COLUMN_NAME}'\n"
            f"SOC–DC Range: {date_range_str}{dc_note}\n"
            f"Source: {source_path}\n\n"
            f"Total visits: {total}\n"
            + ("Breakdown:\n" if total > 0 else "")
            + "\n".join(lines)
        )

        messagebox.showinfo("Visit Counts by User Type", msg)

def main():
    app = VisitCounterApp()
    app.mainloop()

if __name__ == "__main__":
    main()
