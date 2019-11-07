import os
import pandas as pd
import openpyxl
import pyodbc

path = ''
os.chdir(path)

included_extensions = ['MSR.xlsm']
file_names = [fn for fn in os.listdir(path) \
                             if any(fn.endswith(ext) for ext in included_extensions)]

'Read them in and delete the first row for all frames except the first'
excelFiles = [pd.read_excel(name, 'MSR Template', skiprows = 6) for name in file_names]
excelFiles[1:] = [dfs[1:] for dfs in excelFiles[1:]]

"Concat the MSR's and write the data to a new excel sheet"
df = pd.concat(excelFiles)
resultFile = 'Master MSR List.xlsx'
resultSheet = 'MSR Data'
df.to_excel(resultFile, sheet_name = resultSheet, header=False, index=False)

'Retrieve the workbook and worksheet objects.'
workbook  = openpyxl.load_workbook('Master MSR List.xlsx')
worksheet = workbook.active
'Set the bullets column format to bullet points.'
for i in range(1, 101):
    cell = worksheet.cell(row = i, column = 4)
    cell.number_format = '•  @'
workbook.save('Master MSR List.xlsx')
workbook.close()

