''' 
Create a script that scans a folder containing multiple files, calculates the number of rows in each file, and then exports the results to an Excel spreadsheet.
'''

import os
from pathlib import Path
import io
import openpyxl

counter = 0
# extract current directory
user_home = str(Path.home())

# current path / path that contains the folder with the files
documents_path = os.path.join(user_home, 'Documents', 'recordcount')
# Create excel workbook and add a sheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
# Set the header row
worksheet.append(['File Name', 'Line Count'])

# iterate through the folder
for files in os.listdir(documents_path):
    # join the current path with the file name to access it later
    filename = os.path.join(documents_path,files)
    # define the encoding of the file, doesn't work without it
    # iterate through files
    with io.open(filename,encoding="utf8") as f:
        # i = row count, _ is the content of the file which is ignored
        for i, _ in enumerate(f):
            pass
    # output
    # uncomment if you want results to be displayed on the terminal
    # print('file name :',filename,'number of rows is : ',i)
    worksheet.append([filename,i])

# Save the workbook to a file
excel_file = os.path.join(user_home, 'Documents', 'line_counts.xlsx')
workbook.save(excel_file)
print(f'Line counts exported to {excel_file}')
