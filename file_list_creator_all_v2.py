# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 13:06:21 2023

@author: MahdaviAl
"""

# Import essential modules, functions, and file runs

import os
import glob
import pandas as pd
from datetime import datetime
from docx import Document

def get_file_name(file_path):
    return os.path.basename(file_path)

exec(open(r'mark_corrector.py').read()) # Correcting / to \
exec(open(r'word_file_reader.py').read()) ## Reading functions that read word files

##########################################################
### Step 1: Collecting all docx files in 2003-Brampton ###
##########################################################

today = datetime.today().strftime('%y%m%d')

# Take the mother folder(s)
folder_path_list = [r'E:/', r'T:/']

# Take  all docx files and save in a df
docx_files = []
for folder_path in folder_path_list:
    docx_files = docx_files + glob.glob(os.path.join(folder_path, "**/*.docx"), recursive=True)
    
docx_files = [backslash_correct(path) for path in docx_files]
docx_files = pd.Series(docx_files)
docx_files_split = docx_files.str.rsplit((r'/'), n = 1, expand = True)
docx_files_split.columns = ['Path', 'File Name']

# Create processed folder
processed_folder = os.path.join(os.getcwd(), 'processed files', 'f{today}')
os.makedirs(processed_folder, exist_ok = True)


# Make two separate sheets in the saved excel : 1) path and file together; 2) path and file separate
with pd.ExcelWriter(os.path.join(processed_folder + f'file_list_all_docx_{today}.xlsx')) as writer:
    docx_files.to_excel(writer, index = False, header = False, sheet_name = 'files w path')
    docx_files_split.to_excel(writer, sheet_name = 'files path separate', index = False)


## NOTE 1: It will take aproximately 6-10 mins to run Step 1
    
###############################################################
### Step 2: Catching water and air tests in different lists ###
###############################################################

# Create a list of 1) all reviewed files , 2) files filtered for air test, 3) files filtered for water test, and 4) files failed to be read
reviewed_file_list = []
air_file_list = []
water_file_list = []
fail_list = []
permission_error_list = []


# Make repository folders for air and water (date-wise)
repo_folder_air = os.path.join(os.getcwd(), 'repository', f'{today}', 'air_reports')
repo_folder_water = os.path.join(os.getcwd(), 'repository', f'{today}', 'water_reports')

os.makedirs(repo_folder_air, exist_ok = True)
os.makedirs(repo_folder_water, exist_ok = True)

# Locate all reviewed files to repository if targetted, and collect errors if files cannot be opened
i = 0
for file in list(docx_files):
    # print(file)
    if i % 200 == 0:
        print(round(((i/len(docx_files))*100) , 1)) # to exhibit the percentage of progress
    try:
        doc = Document(file)
        text = read_text_docx(file)
        reviewed_file_list.append(file)
    
        if 'field air tightness' in text:
            air_file_list.append(file)
                        
            if file[0:1] == 'T':
                os.makedirs(os.path.join(repo_folder_air, 't_drive'), exist_ok = True)
                new_file = os.path.join(repo_folder_air, 't_drive', f'{get_file_name(file)}')
            else:
                os.makedirs(os.path.join(repo_folder_air, 'e_drive'), exist_ok = True)
                new_file = os.path.join(repo_folder_air, 'e_drive', f'{get_file_name(file)}')
            doc.save(new_file)

        elif 'field water tightness' in text:
            water_file_list.append(file)
            
            if file[0:1] == 'T':
                os.makedirs(os.path.join(repo_folder_water, 't_drive'), exist_ok = True)
                new_file = os.path.join(repo_folder_water, 't_drive', f'{get_file_name(file)}')
            else: 
                os.makedirs(os.path.join(repo_folder_water, 'e_drive'), exist_ok = True)
                new_file = os.path.join(repo_folder_water, 'e_drive', f'{get_file_name(file)}')
            doc.save(new_file)
    
    except Exception as e:
        fail_list.append((file, str(e)))
        # if not os.access(file, os.R_OK):
        #    raise PermissionError(f"The file {file} is not readable.")
        #    permission_error_list.append(file)

        continue
    i += 1


# Convert lists created to pd.Series and DataFrames
reviewed_file_list_sr = pd.Series(reviewed_file_list)
fail_list_df = pd.DataFrame(fail_list, columns = ['File', 'Error'])
# per_err_sr = pd.Series(permission_error_list)

air_file_list_sr = pd.Series(air_file_list)
water_file_list_sr = pd.Series(water_file_list)

Summary = {'all docx files': len(docx_files),
           'all reviewed files': len(reviewed_file_list),
           'files failed to be reviewed': len(fail_list),
           'files with permission errors': len(permission_error_list),
           'all air test files': len(air_file_list),
           'all water test files': len(water_file_list)}

df = pd.DataFrame({'Description': Summary.keys(), 'No of Files': Summary.values()})

# Save listes created in an excel file
with pd.ExcelWriter(os.path.join(processed_folder, f'file_list_reviewed_docx_{today}.xlsx')) as writer2:
    reviewed_file_list_sr.to_excel(writer2, sheet_name = 'reviewed_files', index = False, header = False)
    fail_list_df.to_excel(writer2, sheet_name = 'failed_files', index = False, header = False)
    air_file_list_sr.to_excel(writer2, sheet_name = 'air_test_files', index = False, header = False)
    water_file_list_sr.to_excel(writer2, sheet_name = 'water_test_files', index = False, header = False)
    df.to_excel(writer2, sheet_name = 'overall_stats', index = False, header = True)
    
### NOTE 2: It will take approximately 2:30 - 3:30 h to run Step 2
