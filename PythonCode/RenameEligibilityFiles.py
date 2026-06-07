# -*- coding: utf-8 -*-
import fnmatch
import os
import re
import sys
import tkinter.messagebox
from datetime import date
from tkinter.filedialog import askdirectory


def renameFilesInDirectory(oldDirectory, newDirectory):
    os.chdir(newDirectory)
    print(os.getcwd())
    count = 0
    for file in os.listdir('.'):
        if fnmatch.fnmatch(file, 'EligibilityResponse*') and os.path.isfile(file)==True:
            count+=1
            ptName = re.search('(EligibilityResponse\d+_)(.+)(?=_\d+_)', file).group(2)
            ptName = ptName.replace("_"," ")
            ptName = re.search('(.+) \d+', ptName).group(1)
            newFile = str(ptName + " " + todayDate + " Eligibility.pdf").title()
            print (str(count) + '. ' + file)
            print(newFile)
            os.rename(file, newFile)
    os.chdir(oldDirectory)
    return count

tkinter.messagebox.showinfo(title="Folder Select", message="Select folder with files to rename.")
currentDirectory = askdirectory(title='Select folder') #Show Dialog box and return path of folder.

os.chdir(currentDirectory)

todayDate = date.today().strftime("%Y%m%d")
p = re.compile('\d_\D*_')
iter = 0
for file in os.listdir('.'):
    if fnmatch.fnmatch(file, 'EligibilityResponse*') and os.path.isfile(file)==True:
        iter+=1
        ptName = re.search('(EligibilityResponse\d+_)(.+)(?=_\d+_)', file).group(2)
        ptName = ptName.replace("_"," ")
        ptName = re.search('(.+) \d+', ptName).group(1)
        newFile = str(ptName + " " + todayDate + " Eligibility.pdf").title()
        print (str(iter) + '. ' + file)
        print(newFile)
        os.rename(file, newFile)
    ##If Selection is a EligibilityResponse Folder
    elif fnmatch.fnmatch(file, 'EligibilityResponse*') and os.path.isdir(file)==True:
        iter += renameFilesInDirectory(currentDirectory, file) #Go into directory and do the same process
tkinter.messagebox.showinfo(title="Renaming Finished", message=str(str(iter) +" files renamed."))
# old_name = r"C:\Users\ChristianOrtiz\Downloads\EligibilityResponses*.pdf"
# new_name = r"C:\Users\ChristianOrtiz\Downloads\Last_First_Date_Eligibility.pdf"

# # Renaming the file
# os.rename(old_name, new_name)
    

sys.exit()