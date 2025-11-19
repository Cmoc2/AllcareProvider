# -*- coding: utf-8 -*-
"""
Created on Fri Mar 14 10:35:54 2025

@author: ChristianOrtiz
"""

import os
import shutil

def move_visit_reports(source_folder, destination_folder):
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
        
            destination_file = os.path.join(destination_folder, filename)
            # Move the file
            shutil.move(source_file, destination_file)
            print(f"Moved: {filename}")

# usage
source_folder = 'C:/Users/ChristianOrtiz/Downloads'
destination_folder = 'C:/Users/ChristianOrtiz/Documents/Reports/Report Data/Excel Concat'
move_visit_reports(source_folder, destination_folder)