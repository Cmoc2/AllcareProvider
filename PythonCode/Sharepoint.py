# -*- coding: utf-8 -*-
"""
Created on Fri Feb 14 11:50:05 2025

@author: ChristianOrtiz
"""
import sys
import pandas as pd     # importing pandas to write SharePoint list in excel or csv
import tkinter.messagebox
import tkinter as Tk
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askdirectory
from datetime import datetime, timedelta

import tkinter_optiondialog

"""
from shareplum import Site 
from requests_ntlm import HttpNtlmAuth

Functionality to get the Excel straight from Sharepoint
cred = HttpNtlmAuth("christian.ortiz@allcareprovider.com", "#!!!(insert password)")
site = Site('#sharePoint_url_here', auth=cred)

sp_list = site.List('#SharePoint_list name here') # this creates SharePlum object
data = sp_list.GetListItems('All Items') # this will retrieve all items from list
"""

def fileSelect(source, fileType):
    #Display Message to select a File.
    print(f"Prompt Select {fileType} file.")
    tkinter.messagebox.showinfo(title=f"{source} Select", message=f"Select {source} {fileType} file.")
    #Show Dialog box and return path of file.
    if source == "Referral Report":
        currentFilePath = filedialog.askopenfilename(
            title=f'Select {source} file', filetypes=[("CSV Files","*.csv")]
        )
        print(f"{source} file successfully selected.")
    elif source == "Sharepoint":
        currentFilePath = filedialog.askopenfilename(
            title=f'Select {source} file', filetypes=(("Excel Files","*.xlsx"),("Excel Files","*.xls"))
            )
        print(f"{source} file successfully selected.")
    #Check if file was selected
    if currentFilePath:
        print(f"Selected file: {currentFilePath}")
    else:
        print("No File Selected.")
        root.destroy()
        sys.exit(0)
        
        
    # Load the file into a dataFrame
    if fileType =="Excel":
        df = pd.read_excel(currentFilePath, engine='openpyxl')
    elif fileType =="CSV":
        df = pd.read_csv(currentFilePath, skiprows=5)

    # Convert the 'Referral Date' column to datetime format if it's not already
    df['Referral Date'] = pd.to_datetime(df['Referral Date'])
    
    # Filter the data for a specific referral date, e.g., '2025-02-10'
    print(date_to_compare)
    specific_date = pd.to_datetime(date_to_compare)
    filtered_df = df[df['Referral Date'] == specific_date]
    print(f"Inside fileSelect:\nDate compared: {specific_date}\n{filtered_df} \nLeaving File Select")
    return {"path": currentFilePath,"dataFrame": filtered_df}




#Run the dialog Box - Main Window
def run_dialog():
    #initialize tkinter Window into variable 'root'
    root = Tk()
    #Message Prompt: Allcare or Comcare?
    def run():
        values = ['Allcare','Comcare']
        dlg = tkinter_optiondialog.OptionDialog(root,'Agency Select',"For Which Agency?",values)
        agencyButton["text"] = dlg.result
        print(dlg.result)
    #Message Prompt: Select SP File
    def sharepoint_file_select():
        global sp_df, sp_df_count
        sp_df = fileSelect("Sharepoint", "Excel")
        print(sp_df)
        sp_df_count = count_rows(sp_df["dataFrame"])
        sharepointFileButton["text"] = f"{sp_df_count} rows loaded"
        print('Sharepoint file selected.')
            
    #Message Prompt: Select CSV/Referral Report File
    def ref_report_file_select():
        global myUnity_df, myUnity_df_count
        myUnity_df = fileSelect("Referral Report", "CSV")
        myUnity_df_count = count_rows(myUnity_df["dataFrame"])
        referralReportButton["text"] = f"{myUnity_df_count} rows loaded" #OR myUnity_df["path"]
        print('Ref Report Selected')
        
    def choose_date():
        # Get today's date
        today = datetime.today()
        # Calculate yesterday's date
        yesterday = today - timedelta(days=1)
        # Format the date as MM/DD/YYYY
        formatted_date = yesterday.strftime('%m/%d/%Y')
        comparisonDateButton["text"] = formatted_date
    def begin_comparison():
        root.destroy()
    
    def cancel_process():
        root.destroy()
        sys.exit(0)
    
    #Create the buttons to be used  
    agencyButton = Button(root,text='Select Agency',command=run)
    agencyButton.pack()
    sharepointFileButton = Button(root,text=sp_df['path'] if sp_df else 'Select Sharepoint File',command=sharepoint_file_select)
    sharepointFileButton.pack()
    referralReportButton = Button(root,text="Select Referral Report File",command=ref_report_file_select)
    referralReportButton.pack()
    comparisonDateButton = Button(root,text="2/27/2025",command=choose_date)
    comparisonDateButton.pack()
    compareButton = Button(root,text="Run Comparison",command=begin_comparison)
    compareButton.pack()
    cancelButton = Button(root,text="Cancel",command=cancel_process)
    cancelButton.pack()
    
    root.mainloop()

def sp_match():
    pass

def mu_match():
    pass

def count_rows(dataFrame):
    count = 0
    for row in dataFrame.iterrows():
        count+=1
    return count

def df_check():
    if sp_df == {} or myUnity_df == {}:
        tkinter.messagebox.showinfo(title="File(s) not Selected", message="Select at least two files to compare.")
        return TRUE
    else:
        return FALSE
        
def run_comparison():
    #Begin looking for matching rows. Starting with myUnity.
    matches = []
    match_count = 0
    for index, row in myUnity_df["dataFrame"].iterrows():
        match = sp_df["dataFrame"]["Patient Name"]==row["Patient"]
            
        #if there is a match:
        if not match.empty:
            matches.append(match)
            match_count+=1
            #remove the matched row from both DataFrames
            sp_df["dataFrame"] = sp_df["dataFrame"].drop(match.index)
            myUnity_df["dataFrame"] = myUnity_df["dataFrame"].drop(index)
            #Move on to the next row in the myUnity DataFrame
            continue
    #Make a DataFrame of the results
    matches_df = pd.concat(matches)
    
    # Display the filtered data
    print(sp_df)
    print(myUnity_df)
    print(f"{sp_df_count} Sharepoint Patients.\n{myUnity_df_count} myUnity Patients")
    print(f"{match_count} Patients matched.")
#initialize variable to later be tkinter frame into variable 'root'
root = None
# Place in a var -- Calculate yesterday's date -- Format the date as MM/DD/YYYY
date_to_compare = (datetime.today() - timedelta(days=1)).strftime('%m/%d/%Y')
#initialize dataFrames & count var
sp_df = {}
sp_df_count = None
myUnity_df = {}
myUnity_df_count = None
#Run the dialog Box - Main Window
run_dialog()
print('Post RUn_Dialog')
#check for Pre-Req fulfilled. if fails, (returns true) run dialog again
while df_check():
    print("in while loop")
    run_dialog()
run_comparison()