# -*- coding: utf-8 -*-
"""
Created on Fri Mar 14 10:35:54 2025

@author: ChristianOrtiz
"""

import os
import pandas as pd
import shutil
import tkinter.messagebox

def convert_and_process_reports(source_folder, destination_folder):
    # Ensure the destination folder exists
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    # Clear the destination folder
    for filename in os.listdir(destination_folder):
        file_path = os.path.join(destination_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')
            
    # Iterate over all files in the source folder
    for filename in os.listdir(source_folder):
        # Check if the file name contains "Visit Report - "
        if "Visit Report - " in filename:
            # Construct full file path
            source_file = os.path.join(source_folder, filename)
            
            # If it's an .xls file, convert to .xlsx
            if filename.lower().endswith('.xls'):
                try:
                    df = pd.read_excel(source_file)
                    new_filename = filename.rsplit('.', 1)[0] + '.xlsx'
                    destination_file = os.path.join(destination_folder, new_filename)
                    df.to_excel(destination_file, index=False)
                    os.remove(source_file)
                    print(f"Converted and moved: {new_filename}")
                except Exception as e:
                    print(f"Failed to convert {filename}. Reason: {e}")
            else:
                destination_file = os.path.join(destination_folder, filename)
                shutil.move(source_file, destination_file)
                print(f"Moved: {filename}")
        elif "Patient Dashboard" in filename:
            # Construct full file path
            source_file = os.path.join(source_folder, filename)
            
            # If it's an .xls file, convert to .xlsx
            if filename.lower().endswith('.xls'):
                #Path A: Patient Dashboard - Pending
                if "Pending" in filename:
                    try:
                        df = pd.read_excel(source_file, sheet_name='Patient Dashboard - Pending', engine='xlrd')
                        new_filename = filename.rsplit('.', 1)[0] + '.xlsx'
                        destination_file = os.path.join(patient_dashboard_destination_folder, new_filename)
                        df.to_excel(destination_file, sheet_name='Patient Dashboard - Active', index=False)
                        # Confirm successful write
                        if os.path.exists(destination_file):
                            try:
                                # Try reading the new file to confirm integrity
                                pd.read_excel(destination_file)
                                os.remove(source_file)
                                print(f"Converted and renamed sheet to Active: {new_filename}")
                                print(f"Deleted {filename}")
                            except Exception as e:
                                print(f"File created but could not be read back: {destination_file}. Reason: {e}")
                        else:
                            print(f"Failed to create {destination_file}. Original file not deleted.")

                        print(f"Converted and renamed sheet to Active: {new_filename}")
                    except Exception as e:
                        print(f"Failed to convert {filename}. Reason: {e}")
                #Path B: Patient Dashboard - Active
                elif "Active" in filename:
                    try:
                        df = pd.read_excel(source_file, sheet_name='Patient Dashboard - Active', engine='xlrd')
                        new_filename = filename.rsplit('.', 1)[0] + '.xlsx'
                        destination_file = os.path.join(patient_dashboard_destination_folder, new_filename)
                        df.to_excel(destination_file, sheet_name='Patient Dashboard - Active', index=False)
                        # Confirm successful write
                        if os.path.exists(destination_file):
                            try:
                                # Try reading the new file to confirm integrity
                                pd.read_excel(destination_file)
                                os.remove(source_file)
                                print(f"Converted and renamed sheet to Active: {new_filename}")
                                print(f"Deleted {filename}")
                            except Exception as e:
                                print(f"File created but could not be read back: {destination_file}. Reason: {e}")
                        else:
                            print(f"Failed to create {destination_file}. Original file not deleted.")

                    except Exception as e:
                        print(f"Failed to convert {filename}. Reason: {e}")
   
"""
    Give a warning on what it does:
    This script will process all xls files in the downloads folder:
    - Process Patient Dashboard xls
    - Process and move Visit Reports
    Do you wish to continue?
"""
#Display Message to select a File.
response = tkinter.messagebox.askyesno(title="Process .xls Files", message="This script will process all xls files in the downloads folder:\n\n- Process Patient Dashboard\n- Process and move Visit Reports\n\nDo you wish to continue?")
source_folder = 'C:/Users/ChristianOrtiz/Downloads'
destination_folder = 'C:/Users/ChristianOrtiz/Documents/Reports/Report Data/Excel Concat'
patient_dashboard_destination_folder = 'C:/Users/ChristianOrtiz/Downloads'

if response:
    convert_and_process_reports(source_folder, destination_folder)
    wait = input("Press Enter to continue.")