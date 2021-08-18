#!/usr/bin/python
# -*- coding: utf-8 -*-
# Open .PST file through Open Outlook Data File

# Complete

import win32com.client
import re
import os
import xlwings as xw

# Open saved workbook

wb_path = ''
wb = xw.Book(wb_path)

# Creating Outlook object

Outlook = win32com.client.Dispatch('Outlook.Application'
                                   ).GetNamespace('MAPI')

# To get all emails from Inbox. -> Outlook.GetDefaultFolder(6) Gives primary Inbox in Current Namespace

Outlook.GetDefaultFolder(6).Folders[1].Items

# Better way to access Outlook Data File instead of going current Namespace folders

Outlook.Folders[1].Folders[1].Items

# Access Own Folder 0: Own, 1: Outlook Data File, 2: Outlook Data File ---> You can open many outlook files form File-> Open& Export-> Open Outlook Data File

Outlook.Folders[0].Name

# To iterate over First Data File Opened

row = 2
sht = wb.sheets[0]
for email in Outlook.Folders[1].Folders[1].Items:
    if 'Job application' in email.Subject or 'jcd' \
        in email.Body.lower() or 'jcd' in email.Subject.lower() \
        or 'resume' in email.Body.lower() or 'resume' \
        in email.Subject.lower():

#         print(email.Subject)

        if not 're: ' in email.Subject.lower():
            name = subject = body = positionApplied = receivedTime = \
                workEx = currentEmployer = existingDesignation = \
                currentSalary = location = keySkills = \
                highestEducation = fileName = emailID = contactNumber = \
                attachmentCount = ''

    #     if ("jcd" in email.Body.lower()) or ("Job application" in email.Subject):
    #         print(email.Subject, email.Body, type(email.ReceivedTime), email.Attachments.Count, email.Attachments.Item(1).FileName)

            subject = email.Subject
            body = email.Body.replace('\t', '').replace('\r', ''
                    ).replace('\n', '')
            receivedTime = email.ReceivedTime
            if 'jcd' in email.Body.lower() or 'jcd' \
                in email.Subject.lower() or 'resume' \
                in email.Body.lower() or 'resume' \
                in email.Subject.lower():
                contactNumber = re.findall('[0-9 ]{10,15}', body)
                if len(contactNumber):
                    contactNumber = contactNumber[0]
                    contactNumber = str(contactNumber[:5]) + ' ' \
                        + str(contactNumber[5:])
                else:
                    contactNumber = ''
            else:
                contactNumber = ''
            attachmentCount = email.Attachments.Count
            fileName = email.Attachments.Item(1).FileName
            if not ('jcd' in email.Body.lower() or 'jcd'
                    in email.Subject.lower() or 'resume'
                    in email.Body.lower() or 'resume'
                    in email.Subject.lower()):
                positionApplied = \
                    re.search('Naukri.com for\s"([A-Za-z ]+)".\sI would'
                              , body).group(1)
                try:
                    name = \
                        re.search("Name:([A-Za-z ]+)\sWork Experience",
                                  body).group(1)
                except:
                    pass
                currentSalary = re.search('Salary:([0-9.]+)',
                        body).group(1)

#                 Gives on behalf address

                emailID = email.Sender.Address
                workEx = re.search("Work Experience:(.+)\sSalary:",
                                   body).group(1)
                location = \
                    re.search("Current Location:(.+)\sCurrent Employer:"
                              , body).group(1)
                try:
                    currentEmployer = \
                        re.search("Current Employer:(.+)\sDesignation:"
                                  , body).group(1)
                    existingDesignation = \
                        re.search("Designation:(.+)\sKey Skills:",
                                  body).group(1)
                except:
                    pass
            else:
                pass
            try:
                keySkills = \
                    re.search("Key Skills:(.+)\sHighest Education:",
                              body).group(1)
            except:
                keySkills = ''
            try:
                highestEducation = \
                    re.search('Highest Education:(.+)For more',
                              body).group(1)
            except:
                highestEducation = ''

            emailID = email.Sender.Address

            # To save attachment for emails in path

            pathToSave = ''
            if email.Attachments.Count:
                path = pathToSave + str(row - 1) + '_' \
                    + email.Attachments.Item(1).FileName
                email.Attachments.Item(1).SaveAsFile(path)

                # Not to replace same name file 1_ABC

                resumeFileName = str(row - 1) + '_' \
                    + email.Attachments.Item(1).FileName
                try:
                    sht.range('M'
                              + str(row)).add_hyperlink(address=path,
                            text_to_display=resumeFileName)
                except:
                    sht.range('M' + str(row)).value = resumeFileName
            else:
                resumeFileName = 'No Attachment'

            data = [
                name,
                subject,
                body,
                positionApplied,
                receivedTime,
                workEx,
                currentEmployer,
                existingDesignation,
                currentSalary,
                location,
                keySkills,
                highestEducation,
                resumeFileName,
                emailID,
                contactNumber,
                attachmentCount,
                ]

    #         data = [name, subject, body]

            sht.range('A' + str(row)).value = data
            row += 1
        else:

#             print(row)

            pass

print 'Done'
