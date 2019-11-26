import os
import sys
import time

import pandas as pd
import schedule

import openpyxl
import win32com.client

path = 'C:/Users/Austin Keller/Desktop/MSR Consolidation/TEST'
os.chdir(path)

def consolidate():
    'DOWNLOAD MSR FILES INTO A LOCAL FOLDER PATH'
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI') #Opens Microsoft Outlook

    email = outlook.Folders[0]  # Email
    folder = email.Folders[1]  # Inbox
    subFolder = folder.Folders[2]  # MSR Folder
    print('Email: ' + str(email) + ', Folder: ' + str(folder) + ', Subfolder: ' + str(subFolder))

    subFolderMessages = subFolder.Items  # Individual MSR Emails
    message = subFolderMessages.GetFirst() 

    for i in range(0, subFolder.Items.Count):
        subFolderItemAttachments = message.Attachments
        nbrOfAttachmentInMessage = subFolderItemAttachments.Count
        x = 1
        while nbrOfAttachmentInMessage == x:
            attachment = subFolderItemAttachments.item(x)
            #Saves attachment to location
            attachment.SaveAsFile(str(path) + '/'+ str(attachment))
            break
        message = subFolderMessages.GetNext()



    'CONSOLIDATE THE MSR FILES'
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

    print('The consolidation is complete.')
    sys.exit()
    return
    
consolidate()

#Automatically run the Python during a given time
#schedule.every(1).minute.do(consolidate)
#schedule.every().day.at('08:15').do(consolidate)
##while True:
##    schedule.run_pending()
##    time.sleep(10)


