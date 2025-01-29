# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 09:25:22 2023

@author: MahdaviAl
"""

import os
import glob
import pandas as pd

###########################################################
### Step 1: Collecting all docx files in E and T drives ###
###########################################################

# Correct / to \
exec(open('mark_corrector.py').read())

# Take the mother folder(s)
folder_path_list = [r'E:/', r'T:/']

# Take  all docx files and save in a df
docx_files = []
for folder_path in folder_path_list:
    docx_files.append(glob.glob(os.path.join(folder_path, "**/*.docx"), recursive=True))
    
docx_files = [backslash_correct(path) for path in docx_files]
docx_files = pd.Series(docx_files)
docx_files_split = docx_files.str.rsplit((r'/'), n = 1, expand = True)
docx_files_split.columns = ['Path', 'File Name']

# Make two separate sheets in the saved excel : 1) path and file together; 2) path and file separate
with pd.ExcelWriter('file_list_all_docx.xlsx') as writer:
    docx_files.to_excel(writer, index = False, header = False, sheet_name = 'files w path')
    docx_files_split.to_excel(writer, sheet_name = 'files path separate', index = False)

## NNOTE 1: Step 1 takes about 1 hr (it checks all 100s of 1000s of docx files in the two drives)

#######################################################
### Step 2: Catching test reports in all file lists ###
#######################################################

# Read word reader function files
exec(open('word_file_reader.py').read())
# docx_files = pd.read_excel('file_list_all_docx.xlsx', header = None)[0]

# Create lists for 1) all reviewed files , 2) files filtered for air tests, 3) files filtered for water tests, and 4) files failed to be read
reviewed_file_list = []
air_file_list = []
water_file_list = []
fail_list = []

i = 0

for file in list(docx_files):
    # print(file)
    if i % 500 == 0:
        print(round( ((i/len(docx_files))*100) , 1) ) # to exhibit the percentage of progress
    try:
        text = read_text_docx(file)
        reviewed_file_list.append(file)
    
        if 'field air tightness' in text:
            air_file_list.append(file)
        elif 'field water tightness' in text:
            water_file_list.append(file)
    except:
        fail_list.append(file)
        continue
    i += 1

# Convert lists created to pd.Series
reviewed_file_list_sr = pd.Series(reviewed_file_list)
fail_list_sr = pd.Series(fail_list)

air_file_list_sr = pd.Series(air_file_list)
water_file_list_sr = pd.Series(water_file_list)

Summary = {'all docx files': len(docx_files),
           'all reviewed files': len(reviewed_file_list),
           'files failed to be reviewed': len(fail_list),
           'all air test files': len(air_file_list),
           'all water test files': len(water_file_list)}

df = pd.DataFrame({'Description': Summary.keys(), 'No of Files': Summary.values()})

# Save listes created in an excel file
with pd.ExcelWriter(r'file_list_reviewed_docx.xlsx') as writer2:
    reviewed_file_list_sr.to_excel(writer2, sheet_name = 'reviewed_files', index = False, header = False)
    fail_list_sr.to_excel(writer2, sheet_name = 'failed_files', index = False, header = False)
    air_file_list_sr.to_excel(writer2, sheet_name = 'air_test_files', index = False, header = False)
    water_file_list_sr.to_excel(writer2, sheet_name = 'water_test_files', index = False, header = False)
    df.to_excel(writer2, sheet_name = 'overall_stats', index = False, header = True)

## NOTE 2: It will take aproximately 12-24 h to run Step 2
