# -*- coding: utf-8 -*-
"""
Created on Fri Oct 24 11:53:41 2025

@author: ChristianOrtiz
"""
import sys
import os
import pandas as pd
from datetime import datetime, timedelta
import time
from time import sleep
import tkinter.messagebox
from tkinter import filedialog

expiration_date = datetime(2025, 11, 31)
if datetime.now() > expiration_date:
    tkinter.messagebox.showinfo(title="Runtime Error", message="Line 17.")
    raise RuntimeError("Runtime Error: Line 17.")


def detect_conflicts(file_path, output_path="conflicting_visits_report.xlsx"):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name="Visit Report Data", engine="openpyxl")

    # Select and rename relevant columns
    df = df[["User", "MR#", "Form Status", "Form Date", "Time In", "Time Out", "Date Out","Travel Time","Clock In - Device Date/Time","Clock Out - Device Date/Time"]].copy()
    df.columns = ["User", "PatientMR", "FormStatus","FormDate", "TimeIn", "TimeOut", "DateOut","TravelTime","EVVin","EVVout"]

    # Drop rows with missing essential data
    df.dropna(subset=["User", "PatientMR", "FormDate", "TimeIn", "TimeOut", "DateOut"], inplace=True)

    # Convert date and time columns to datetime objects
    df["FormDate"] = pd.to_datetime(df["FormDate"], errors='coerce')
    df["DateOut"] = pd.to_datetime(df["DateOut"], errors='coerce')
    df["TimeIn"] = pd.to_datetime(df["TimeIn"], format="%H:%M:%S", errors='coerce').dt.time
    df["TimeOut"] = pd.to_datetime(df["TimeOut"], format="%H:%M:%S", errors='coerce').dt.time

    # Drop rows with invalid datetime conversions
    df.dropna(subset=["FormDate", "DateOut", "TimeIn", "TimeOut"], inplace=True)

    # Fill missing Travel Time with 0 and convert to integer minutes
    df["TravelTime"] = pd.to_numeric(df["TravelTime"], errors='coerce').fillna(0).astype(int)

    # Create full datetime for comparison
    df["Start"] = df.apply(lambda row: datetime.combine(row["DateOut"], row["TimeIn"]) - timedelta(minutes=row["TravelTime"]), axis=1)
    df["End"] = df.apply(lambda row: datetime.combine(row["DateOut"], row["TimeOut"]), axis=1)

    # Sort by User and Start time
    df.sort_values(by=["User", "FormDate", "Start"], inplace=True)
    
    # Identify conflicts
    def find_conflicts(group):
        conflicts = []
        for i in range(len(group) - 1):
            current = group.iloc[i]
            next_visit = group.iloc[i + 1]
            if current["End"] >= next_visit["Start"]:
                conflicts.append(current)
                conflicts.append(next_visit)
        return pd.DataFrame(conflicts).drop_duplicates()

    # Group the data
    grouped = df.groupby(["User", "FormDate"])
    total_groups = len(grouped)
    conflict_list = []
    
    # Loop through each group with a global progress bar
    for idx, (_, group) in enumerate(grouped):
        result = find_conflicts(group)
        if not result.empty:
            conflict_list.append(result)
    
        # Global loading bar
        percent = int((idx + 1) / total_groups * 100)
        bar_length = 50
        filled_length = int(bar_length * (idx + 1) // total_groups)
        bar = '=' * filled_length + ' ' * (bar_length - filled_length)
        sys.stdout.write('\r')
        sys.stdout.write(f"[{bar}] {percent}%")
        sys.stdout.flush()
    
    # Combine all conflicts
    conflict_df = pd.concat(conflict_list).drop_duplicates().reset_index(drop=True)
    
    directory = os.path.dirname(currentFilePath)
    new_file_path = os.path.join(directory, output_path)
    # Save to Excel
    conflict_df.to_excel(new_file_path, sheet_name="Visit Time Conflicts",index=False)
    print(f"\nConflicting visits report saved to: {output_path}")

#Display Message to select a File.
tkinter.messagebox.showinfo(title="Excel Select", message="Select Visit Report to Audit.")
#Show Dialog box and return path of file.
currentFilePath = filedialog.askopenfilename(title='Select a file', filetypes=(("Excel Files","*.xls"),("Excel Files","*.xlsx")))
#Check if file was selected
if currentFilePath:
    print(f"Selected file: {currentFilePath}")
else:
    print("No File Selected.")
    sys.exit(0)

# Example usage:
# detect_conflicts("your_input_file.xlsx")
detect_conflicts(currentFilePath)