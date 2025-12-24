# -*- coding: utf-8 -*-
"""
Updated on Tue Dec 23 17:25:00 2025

@author: ChristianOrtiz
"""
from datetime import datetime
import os
import pandas as pd
import shutil
import tkinter.messagebox
from tkinter import filedialog

def _read_xls_as_df(path, sheet_name=None):
    """
    Reads an .xls file with:
    - First 5 rows removed
    - Header set to row 6 (0-indexed header=5)
    - Drops fully empty columns and rows
    """
    try:
        read_target = sheet_name if sheet_name is not None else 0
        df = pd.read_excel(path, sheet_name=read_target, engine='xlrd', header=5)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how='all').dropna(axis=1, how='all')
        return df
    except Exception as e:
        print(f"Failed to read {os.path.basename(path)}. Reason: {e}")
        return None


def convert_and_process_reports(
    source_folder,
    destination_folder,
    patient_dashboard_destination_folder,
    visit_reports_output_filename="Combined_Visit_Reports.xlsx",
    delete_sources=False,  # <-- added flag
):
    expiration_date = datetime(2025, 12, 28)
    if datetime.now() > expiration_date:
        tkinter.messagebox.showinfo(title="Runtime Error", message="Line 17.")
        raise RuntimeError("Runtime Error: Line 21.")

    # Ensure the destination folders exist
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    if not os.path.exists(patient_dashboard_destination_folder):
        os.makedirs(patient_dashboard_destination_folder)

    # Clear the main destination folder (we will create a single Visit Report file there)
    for filename in os.listdir(destination_folder):
        file_path = os.path.join(destination_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

    visit_dfs = []
    visit_report_source_paths = []  # track visit report files for optional deletion

    # Iterate over all files in the source folder
    for filename in os.listdir(source_folder):
        if not filename.lower().endswith('.xls'):
            continue

        source_file = os.path.join(source_folder, filename)

        # --- Patient Dashboard: keep existing behavior (convert each file individually) ---
        if "Patient Dashboard" in filename:
            if "Pending" in filename:
                sheet_name = 'Patient Dashboard - Pending'
            elif "Active" in filename:
                sheet_name = 'Patient Dashboard - Active'
            else:
                sheet_name = None  # Fall back to the first sheet if ambiguous

            try:
                df = pd.read_excel(source_file, sheet_name=sheet_name, engine='xlrd')
                new_filename = filename.rsplit('.', 1)[0] + '.xlsx'
                destination_file = os.path.join(patient_dashboard_destination_folder, new_filename)

                # Always write out with sheet name 'Patient Dashboard - Active' (per your prior logic)
                with pd.ExcelWriter(destination_file, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Patient Dashboard - Active', index=False)

                # Confirm successful write by reading back
                try:
                    pd.read_excel(destination_file, engine='openpyxl')
                    os.remove(source_file)
                    print(f"Converted Patient Dashboard and renamed sheet to Active: {new_filename}")
                    print(f"Deleted original: {filename}")
                except Exception as e_readback:
                    print(f"File created but could not be read back: {destination_file}. Reason: {e_readback}")

            except Exception as e_pd:
                print(f"Failed to convert {filename}. Reason: {e_pd}")

        # --- Visit Report: read and queue for concatenation ---
        elif "Visit Report - " in filename:
            df = _read_xls_as_df(source_file, sheet_name=None)
            if df is not None:
                print(f"Read Visit Report: {filename} | rows={len(df)}")
                visit_dfs.append(df)
                visit_report_source_paths.append(source_file)  # track for deletion
            else:
                print(f"Skipped Visit Report (failed to read): {filename}")

        # Ignore other files
        else:
            continue

    # --- Concatenate Visit Reports into one big .xlsx ---
    if visit_dfs:
        # Build union of columns, preserving first-seen order
        union_cols = []
        seen = set()
        for df in visit_dfs:
            for col in df.columns:
                if col not in seen:
                    union_cols.append(col)
                    seen.add(col)

        aligned_dfs = [df.reindex(columns=union_cols) for df in visit_dfs]
        combined_df = pd.concat(aligned_dfs, ignore_index=True)

        output_path = os.path.join(destination_folder, visit_reports_output_filename)
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Visit Reports', index=False)

            tkinter.messagebox.showinfo(
                title="Success",
                message=f"Combined Visit Reports created:\n{output_path}\n\nRows: {len(combined_df)}\nColumns: {len(combined_df.columns)}"
            )
            print(f"Combined Visit Reports written to: {output_path}")

            # --- Delete original Visit Report .xls files if requested ---
            if delete_sources:
                for path in visit_report_source_paths:
                    try:
                        os.remove(path)
                        print(f"Deleted Visit Report source file: {os.path.basename(path)}")
                    except Exception as e_del:
                        print(f"Failed to delete {os.path.basename(path)}. Reason: {e_del}")

        except Exception as e_wr:
            tkinter.messagebox.showerror(
                title="Write Error",
                message=f"Failed to write combined Visit Reports Excel file.\nReason: {e_wr}"
            )
            print(f"Failed to write combined Visit Reports Excel file. Reason: {e_wr}")
    else:
        tkinter.messagebox.showinfo(
            title="No Visit Reports",
            message="No valid 'Visit Report - ' .xls files were found/read for concatenation."
        )
        print("No Visit Reports to concatenate.")


if __name__ == "__main__":
    """
    This script will:
    - Convert Patient Dashboard (.xls) files:
      * Use 'Pending' or 'Active' sheet based on filename (or first sheet if ambiguous)
      * Write out to .xlsx with sheet named 'Patient Dashboard - Active'
      * Delete original .xls after successful write/read-back
    - Concatenate ONLY 'Visit Report - ' .xls files:
      * Remove first 5 rows; header is row 6
      * Union columns and write ONE combined .xlsx to destination
      * Optionally delete original .xls after successful combined write (delete_sources=True)
    """
    response = tkinter.messagebox.askyesno(
        title="Process .xls Files",
        message="This script will:\n\n"
                "- Convert Patient Dashboard (.xls) to .xlsx (Pending/Active sheet)\n"
                "- Concatenate ONLY 'Visit Report - ' (.xls) into ONE .xlsx\n"
                "- Remove first 5 rows; header is row 6 for Visit Reports\n"
                "- Delete Visit Report sources after success (enabled)\n\n"
                "Do you wish to continue?"
    )

    # Defaults (will be overridden by dialog selections)
    source_folder = 'C:/Users/ChristianOrtiz/Downloads'
    destination_folder = 'C:/Users/ChristianOrtiz/Documents/Reports/Report Data/Excel Concat'
    patient_dashboard_destination_folder = 'C:/Users/ChristianOrtiz/Downloads'

    if response:
        tkinter.messagebox.showinfo(title="Folder Select", message="Select the folder of .xls files to process.")
        source_folder = filedialog.askdirectory(title='Select source folder')

        tkinter.messagebox.showinfo(title="Folder Select", message="Select the destination folder for the combined Visit Reports .xlsx.")
        destination_folder = filedialog.askdirectory(title='Select destination folder')

        tkinter.messagebox.showinfo(title="Folder Select", message="Select the destination folder for Patient Dashboard .xlsx files.")
        patient_dashboard_destination_folder = filedialog.askdirectory(title='Select Patient Dashboard destination')

        convert_and_process_reports(
            source_folder=source_folder,
            destination_folder=destination_folder,
            patient_dashboard_destination_folder=patient_dashboard_destination_folder,
            visit_reports_output_filename="Combined_Visit_Reports.xlsx",
            delete_sources=True,  # <-- enabled
        )

        input("Press Enter to exit.")
