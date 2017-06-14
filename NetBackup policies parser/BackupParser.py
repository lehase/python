#!/usr/bin/python
# -*- coding: utf-8 -*-

import re
import xlwt

InputFile = open('policies.txt', 'r')
ExcelFile = 'policies.xls'
WorkBook = xlwt.Workbook(encoding='utf-8')
WorkSheet = WorkBook.add_sheet('Policies')

i = 0

WorkSheet.write(0, i, 'Policy Name')
i+=1
WorkSheet.write(0, i, 'Policy Type')
i+=1
WorkSheet.write(0, i, 'Active')
i+=1
WorkSheet.write(0, i, 'Client Compress')
i+=1
WorkSheet.write(0, i, 'HW/OS/Client')
i+=1
WorkSheet.write(0, i, 'Include')
i+=1
WorkSheet.write(0, i, 'Schedule_Type')
i+=1
WorkSheet.write(0, i, 'Schedule_Frequency')
i+=1
WorkSheet.write(0, i, 'Schedule_Retention Level')
i+=1
WorkSheet.write(0, i, 'Schedule_Daily Windows')







# WorkSheet.write(0, 0, 'Policy Name')
# WorkSheet.write(0, 1, 'Policy Type')
# WorkSheet.write(0, 2, 'Active')
# WorkSheet.write(0, 3, 'Client Compress')
# WorkSheet.write(0, 4, 'Follow NFS Mounts')
# WorkSheet.write(0, 5, 'Cross Mount Points')
# WorkSheet.write(0, 6, 'Collect TIR info')
# WorkSheet.write(0, 7, 'Block Incremental')
# WorkSheet.write(0, 8, 'Mult. Data Streams')
# WorkSheet.write(0, 9, 'Client Encrypt')
# WorkSheet.write(0, 10, 'Checkpoint')
# WorkSheet.write(0, 11, 'Policy Priority')
# WorkSheet.write(0, 12, 'Max Jobs/Policy')
# WorkSheet.write(0, 13, 'Disaster Recovery')
# WorkSheet.write(0, 14, 'Collect BMR info')
# WorkSheet.write(0, 15, 'Residence')
# WorkSheet.write(0, 16, 'Volume Pool')
# WorkSheet.write(0, 17, 'Server Group')
# WorkSheet.write(0, 18, 'Keyword')
# WorkSheet.write(0, 19, 'Data Classification')
# WorkSheet.write(0, 20, 'Residence is Storage Lifecycle Policy')
# WorkSheet.write(0, 21, 'Application Discovery')
# WorkSheet.write(0, 22, 'Discovery Lifetime')
# WorkSheet.write(0, 23, 'ASC Application and attributes')
# WorkSheet.write(0, 24, 'Granular Restore Info')
# WorkSheet.write(0, 25, 'Ignore Client Direct')
# WorkSheet.write(0, 26, 'Use Accelerator')
# WorkSheet.write(0, 27, 'HW/OS/Client')
# WorkSheet.write(0, 28, 'Include')
# WorkSheet.write(0, 29, 'Schedule_Type')
# WorkSheet.write(0, 30, 'Schedule_Frequency')
# WorkSheet.write(0, 31, 'Schedule_Synthetic')
# WorkSheet.write(0, 32, 'Schedule_Checksum Change Detection')
# WorkSheet.write(0, 33, 'Schedule_PFI Recovery')
# WorkSheet.write(0, 34, 'Schedule_Maximum MPX')
# WorkSheet.write(0, 35, 'Schedule_Retention Level')
# WorkSheet.write(0, 36, 'Schedule_Number Copies')
# WorkSheet.write(0, 37, 'Schedule_Fail on Error')
# WorkSheet.write(0, 38, 'Schedule_Residence')
# WorkSheet.write(0, 39, 'Schedule_Volume Pool')
# WorkSheet.write(0, 40, 'Schedule_Server Group')
# WorkSheet.write(0, 41, 'Schedule_Residence is Storage Lifecycle Policy')
# WorkSheet.write(0, 42, 'Schedule_Daily Windows')

Data = []
Dict = {}

DataList = InputFile.readlines()





List_Schedule = []
i = 0
id = 1
while i < len(DataList):
    EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
    SplitLine = EditLine.split(':', 1)
    if SplitLine[0] == 'Policy Name':
        if List_Schedule:
            Dict['Schedule'] = List_Schedule
        if len(Dict) > 7:
            Data.append(Dict)
        Dict = {}
        List_Schedule = []
        Dict['ID'] = id
        id += 1
        Dict['Policy Name'] = SplitLine[1]
        i += 1
    elif SplitLine[0] == 'Policy Type':
        Dict['Policy Type'] = SplitLine[1]
        i += 1
    elif SplitLine[0] == 'Active':
        Dict['Active'] = SplitLine[1]
        i += 1
    elif SplitLine[0] == 'Effective date':
        Dict['Effective date'] = SplitLine[1]
        i += 1
    elif SplitLine[0] == 'Client Compress':
        Dict['Client Compress'] = SplitLine[1]
        i += 1
    
    elif SplitLine[0] == 'HW/OS/Client':
        Dict['HW/OS/Client'] = SplitLine[1]
        i += 1
    elif SplitLine[0] == 'Include':
        # print '______________' + str(i+1) + '______________' + str(id)
        SP = False
        Incl = str()
        while not re.match('^$', EditLine):
            EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
            SplitLine = EditLine.split(':', 1)
            # print str(EditLine) + ' ===> ' + str(i+1)
            if len(SplitLine) == 2 and SP == False:
                Incl += str(SplitLine[1]) + '\n'
                SP = True
            elif not re.match('^$', EditLine):
                Incl += str(EditLine) + '\n'
            if not re.match('^$', EditLine):
                i += 1
            # print 'i+=1'
        Incl = Incl[:-2]
        Dict['Include'] = Incl
    elif SplitLine[0] == 'Schedule':
        # print 'Schedule: ' + str(i+1)
        Dict_Schedule = {}
        i += 1
        while not re.match('^$', EditLine):
            EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
            SplitLine = EditLine.split(':', 1)
            if SplitLine[0] == 'Type':
                # print 'Type: ' + str(i+1)
                Dict_Schedule['Schedule_Type'] = SplitLine[1]
            elif SplitLine[0] == 'Frequency':
                Dict_Schedule['Schedule_Frequency'] = SplitLine[1]
            
            elif SplitLine[0] == 'Retention Level':
                Dict_Schedule['Schedule_Retention Level'] = SplitLine[1]

            elif SplitLine[0] == 'Daily Windows':
                i += 1
                Daily_Windows = str()
                while not re.match('^$', EditLine):
                    EditLine = (re.sub(r'\s\s\s\s\s\s+', '', DataList[i])).replace('\n', '')
                    Daily_Windows += str(EditLine) + '\n'
                    if not re.match('^$', EditLine):
                        i += 1
                Daily_Windows = Daily_Windows[:-2]
                # print 'i+=1'
                Dict_Schedule['Schedule_Daily Windows'] = Daily_Windows
            # print Daily_Windows
            if not re.match('^$', EditLine):
                i += 1
        List_Schedule.append(Dict_Schedule)
    else:
        i += 1

num = 1
for index in Data:
    for i in index['Schedule']:
        j=0
        WorkSheet.write(num, j, index['Policy Name'])
        j+=1
        WorkSheet.write(num, j, index['Policy Type'])
        j += 1
        WorkSheet.write(num, j, index['Active'])
        j += 1
        try:
            WorkSheet.write(num, j, index['Client Compress'])
            j += 1
        except:
            WorkSheet.write(num, j, 'None')
            j += 1
        try:
            WorkSheet.write(num, j, index['HW/OS/Client'])
            j += 1
        except:
            WorkSheet.write(num, j, 'None')
            j += 1
        try:
            WorkSheet.write(num, j, index['Include'])
            j += 1
        except:
            WorkSheet.write(num, j, 'None')
            j += 1
        try:
            WorkSheet.write(num, j, i['Schedule_Type'])
            j += 1
        except:
            WorkSheet.write(num, j, 'None')
            j += 1
        try:
            WorkSheet.write(num, j, i['Schedule_Frequency'])
            j += 1
        except:
            WorkSheet.write(num, j, 'None')
            j += 1
        
        
        try:
            # test = i['Schedule_Daily Windows'].encode('UTF-8')
            WorkSheet.write(num, 42, i['Schedule_Daily Windows'])
            j += 1
        except:
            WorkSheet.write(num, 42, 'None')
            j += 1
            print '1'
        num += 1

# WorkSheet.write(10, 10, test)
WorkBook.save(ExcelFile)
