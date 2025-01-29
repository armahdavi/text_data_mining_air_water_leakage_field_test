# -*- coding: utf-8 -*-
"""
Program to find errors in reading reports (docx)
Created on Tue Sep  3 10:49:54 2024

@author: MahdaviAl
"""

import pandas as pd
from docx import Document 
exec(open(r'mark_corrector.py').read())

# Step 1: Reading all water and air test files
df = pd.read_excel('file_list_reviewed_docx.xlsx', sheet_name = 'water_test_files', header = None)
df = pd.concat([df, pd.read_excel('/file_list_reviewed_brm_docx.xlsx', sheet_name = 'water_test_files', header = None)], ignore_index = True)

# Function to check if a file is readable
def check_file(file_path):
    try:
        Document(file_path)
        return True, None
    except Exception as e:
        return False, str(e)


# Apply the function to the DataFrame
df[['is_readable', 'error']] = df[0].apply(lambda x: pd.Series(check_file(x)))

# Count number of readable files
print(len(df[df['is_readable']]))
