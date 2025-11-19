# -*- coding: utf-8 -*-
"""
Created on Fri Feb 14 10:03:36 2025

@author: ChristianOrtiz
"""
import fnmatch
import os
import re
import sys
import tkinter.messagebox
import datetime
from tkinter import filedialog
from tkinter.filedialog import askdirectory
import pandas as pd

#Display Message to select a File.
tkinter.messagebox.showinfo(title="Excel Select", message="Select \"Patient Dashboard - Pending\" with sheetname to rename.")
#Show Dialog box and return path of file.
currentFilePath = filedialog.askopenfilename(title='Select a file', filetypes=(("Excel Files","*.xls"),("Excel Files","*.xlsx")))
#Check if file was selected
if currentFilePath:
    print(f"Selected file: {currentFilePath}")
else:
    print("No File Selected.")
    sys.exit(0)

#Get working directory of the chosen File.
directory = os.path.dirname(currentFilePath)

# Load the .xls file
df = pd.read_excel(currentFilePath, sheet_name='Patient Dashboard - Pending', engine='xlrd')

# Get the current date to append to the file name
dateToday = datetime.datetime.now().date().strftime("%m-%d-%Y")
# Save the DataFrame to a new .xlsx file with the updated sheet name
new_file_name = f'Patient Dashboard - Pending {dateToday}.xlsx'
new_file_path = os.path.join(directory, new_file_name)
df.to_excel(new_file_path, sheet_name='Patient Dashboard - Active', index=False)

#Show success results
tkinter.messagebox.showinfo(title="Sheet Rename Successful", message=f"The sheet name has been successfully changed to 'Patient Dashboard - Active' and saved to {new_file_path}.")

print(f"The sheet name has been successfully changed to 'Patient Dashboard - Active' and saved to {new_file_path}.")
#End Program
sys.exit(0)