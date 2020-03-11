import os
import sys
import time

import pandas as pd
import schedule

import openpyxl
import win32com.client

path = 'insert file path here'
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

        if nbrOfAttachmentInMessage > 0:
            for j in range(0, nbrOfAttachmentInMessage):
                attachment = subFolderItemAttachments.item(j + 1)
                if str(attachment)[-8:] == 'MSR.xlsm':
                    #Saves attachment to location
                    attachment.SaveAsFile(str(path) + '/'+ str(attachment))
        message = subFolderMessages.GetNext()



    'CONSOLIDATE THE MSR FILES'
    included_extensions = ['MSR.xlsm']
    file_names = [fn for fn in os.listdir(path) \
                                 if any(fn.endswith(ext) for ext in included_extensions)]

    'Read them in and delete the first row for all frames except the first'
    excelFiles = [pd.read_excel(name, 'MSR Template', skiprows = 7) for name in file_names]
    excelFiles[1:] = [dfs[1:] for dfs in excelFiles[1:]]

    "Concat the MSR's and write the data to a new excel sheet"
    df = pd.concat(excelFiles)
    resultFile = 'Master MSR List.xlsx'
    resultSheet = 'MSR Data'
    df.to_excel(resultFile, sheet_name = resultSheet, header=False, index=False)

    'Retrieve the workbook and worksheet objects.'
    workbook  = openpyxl.load_workbook('Master MSR List.xlsx')
    worksheet = workbook.active

    'Delete empty column "C"'
    worksheet.delete_cols(3,1)
    
    'Set the bullets column format to bullet points.'
    for i in range(1, 101):
        cell = worksheet.cell(row = i+1, column = 6)
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


