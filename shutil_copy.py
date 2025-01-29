# -*- coding: utf-8 -*-
"""
Created on Fri Jan  3 11:54:52 2025

@author: MahdaviAl
"""


import shutil

# Define source and destination paths
source_path = r'C:\python_projects\air_water_tightness_compliance_tests\report_4_2022-02-01.docx'
destination_path = r'C:\EXP\The Met- Townhomes- Entry Door Air Leakage Test (TH 20) - 2022-02-01.docx.docx'

# Copy the file
try:
    shutil.copy(source_path, destination_path)
    print(f"File copied successfully from {source_path} to {destination_path}")
except FileNotFoundError:
    print("The source file was not found.")
except PermissionError:
    print("Permission denied while copying the file.")
except Exception as e:
    print(f"An error occurred: {e}")
