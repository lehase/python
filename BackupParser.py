#!/usr/bin/python
# -*- coding: utf-8 -*-

import re
import xlwt



InputFile = open('C:\Scripts\policies.txt', 'r')
ExcelFile = 'C:/Scripts/policies.xls'
WorkBook = xlwt.Workbook(encoding='utf-8')
WorkSheet = WorkBook.add_sheet('Policies')
WorkSheet.write(0, 0, 'Policy Name')
WorkSheet.write(0, 1, 'Policy Type')
WorkSheet.write(0, 2, 'Active')
WorkSheet.write(0, 3, 'Client Compress')
WorkSheet.write(0, 4, 'Follow NFS Mounts')
WorkSheet.write(0, 5, 'Cross Mount Points')
WorkSheet.write(0, 6, 'Collect TIR info')
WorkSheet.write(0, 7, 'Block Incremental')
WorkSheet.write(0, 8, 'Mult. Data Streams')
WorkSheet.write(0, 9, 'Client Encrypt')
WorkSheet.write(0, 10, 'Checkpoint')
WorkSheet.write(0, 11, 'Policy Priority')
WorkSheet.write(0, 12, 'Max Jobs/Policy')
WorkSheet.write(0, 13, 'Disaster Recovery')
WorkSheet.write(0, 14, 'Collect BMR info')
WorkSheet.write(0, 15, 'Residence')
WorkSheet.write(0, 16, 'Volume Pool')
WorkSheet.write(0, 17, 'Server Group')
WorkSheet.write(0, 18, 'Keyword')
WorkSheet.write(0, 19, 'Data Classification')
WorkSheet.write(0, 20, 'Residence is Storage Lifecycle Policy')
WorkSheet.write(0, 21, 'Application Discovery')
WorkSheet.write(0, 22, 'Discovery Lifetime')
WorkSheet.write(0, 23, 'ASC Application and attributes')
WorkSheet.write(0, 24, 'Granular Restore Info')
WorkSheet.write(0, 25, 'Ignore Client Direct')
WorkSheet.write(0, 26, 'Use Accelerator')
WorkSheet.write(0, 27, 'HW/OS/Client')
WorkSheet.write(0, 28, 'Include')
WorkSheet.write(0, 29, 'Schedule_Type')
WorkSheet.write(0, 30, 'Schedule_Frequency')
WorkSheet.write(0, 31, 'Schedule_Synthetic')
WorkSheet.write(0, 32, 'Schedule_Checksum Change Detection')
WorkSheet.write(0, 33, 'Schedule_PFI Recovery')
WorkSheet.write(0, 34, 'Schedule_Maximum MPX')
WorkSheet.write(0, 35, 'Schedule_Retention Level')
WorkSheet.write(0, 36, 'Schedule_Number Copies')
WorkSheet.write(0, 37, 'Schedule_Fail on Error')
WorkSheet.write(0, 38, 'Schedule_Residence')
WorkSheet.write(0, 39, 'Schedule_Volume Pool')
WorkSheet.write(0, 40, 'Schedule_Server Group')
WorkSheet.write(0, 41, 'Schedule_Residence is Storage Lifecycle Policy')
WorkSheet.write(0, 42, 'Schedule_Daily Windows')

Data = []
Dict = {}

DataList = InputFile.readlines()

List_Schedule  = []
i = 0
id = 1
while i < len(DataList):
	EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
	SplitLine = EditLine.split(':', 1)
	if SplitLine[0] == 'Policy Name':
		if List_Schedule:
			Dict['Schedule'] = List_Schedule
		if len(Dict) > 10:
			Data.append(Dict)
		Dict = {}
		List_Schedule  = []
		Dict['ID'] = id
		id += 1
		Dict['Policy Name'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Policy Type':
		Dict['Policy Type'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Active':
		Dict['Active'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Effective date':
		Dict['Effective date'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Client Compress':
		Dict['Client Compress'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Follow NFS Mounts':
		Dict['Follow NFS Mounts'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Cross Mount Points':
		Dict['Cross Mount Points'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Collect TIR info':
		Dict['Collect TIR info'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Block Incremental':
		Dict['Block Incremental'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Mult. Data Streams':
		Dict['Mult. Data Streams'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Client Encrypt':
		Dict['Client Encrypt'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Checkpoint':
		Dict['Checkpoint'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Policy Priority':
		Dict['Policy Priority'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Max Jobs/Policy':
		Dict['Max Jobs/Policy'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Disaster Recovery':
		Dict['Disaster Recovery'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Collect BMR info':
		Dict['Collect BMR info'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Residence':
		Dict['Residence'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Volume Pool':
		Dict['Volume Pool'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Server Group':
		Dict['Server Group'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Keyword':
		Dict['Keyword'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Data Classification':
		Dict['Data Classification'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Residence is Storage Lifecycle Policy':
		Dict['Residence is Storage Lifecycle Policy'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Application Discovery':
		Dict['Application Discovery'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Discovery Lifetime':
		Dict['Discovery Lifetime'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'ASC Application and attributes':
		Dict['ASC Application and attributes'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Granular Restore Info':
		Dict['Granular Restore Info'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Ignore Client Direct':
		Dict['Ignore Client Direct'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Use Accelerator':
		Dict['Use Accelerator'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'HW/OS/Client':
		Dict['HW/OS/Client'] = SplitLine[1]
		i+=1
	elif SplitLine[0] == 'Include':
		# print '______________' + str(i+1) + '______________' + str(id)
		SP = False
		Incl = str()
		while not re.match('^$', EditLine):
			EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
			SplitLine = EditLine.split(':', 1)
			# print str(EditLine) + ' ===> ' + str(i+1)
			if len(SplitLine) == 2 and SP == False:
				Incl += str(SplitLine[1])+'\n'
				SP = True
			elif not re.match('^$', EditLine):
				Incl += str(EditLine)+'\n'
			if not re.match('^$', EditLine):
				i+=1
				# print 'i++'
		Incl = Incl[:-2]
		Dict['Include'] = Incl
	elif SplitLine[0] == 'Schedule':
		# print 'Schedule: ' + str(i+1)
		Dict_Schedule = {}
		i+=1
		while not re.match('^$', EditLine):
			EditLine = (re.sub(r'\s\s+', '', DataList[i])).replace('\n', '')
			SplitLine = EditLine.split(':', 1)
			if SplitLine[0] == 'Type':
				# print 'Type: ' + str(i+1)
				Dict_Schedule['Schedule_Type'] = SplitLine[1]
			elif SplitLine[0] == 'Frequency':
				Dict_Schedule['Schedule_Frequency'] = SplitLine[1]
			elif SplitLine[0] == 'Synthetic':
				Dict_Schedule['Schedule_Synthetic'] = SplitLine[1]
			elif SplitLine[0] == 'Checksum Change Detection':
				Dict_Schedule['Schedule_Checksum Change Detection'] = SplitLine[1]
			elif SplitLine[0] == 'PFI Recovery':
				Dict_Schedule['Schedule_PFI Recovery'] = SplitLine[1]
			elif SplitLine[0] == 'Maximum MPX':
				Dict_Schedule['Schedule_Maximum MPX'] = SplitLine[1]
			elif SplitLine[0] == 'Retention Level':
				Dict_Schedule['Schedule_Retention Level'] = SplitLine[1]
			elif SplitLine[0] == 'Number Copies':
				Dict_Schedule['Schedule_Number Copies'] = SplitLine[1]
			elif SplitLine[0] == 'Fail on Error':
				Dict_Schedule['Schedule_Fail on Error'] = SplitLine[1]
			elif SplitLine[0] == 'Residence':
				Dict_Schedule['Schedule_Residence'] = SplitLine[1]
			elif SplitLine[0] == 'Volume Pool':
				Dict_Schedule['Schedule_Volume Pool'] = SplitLine[1]
			elif SplitLine[0] == 'Server Group':
				Dict_Schedule['Schedule_Server Group'] = SplitLine[1]
			elif SplitLine[0] == 'Residence is Storage Lifecycle Policy':
				Dict_Schedule['Schedule_Residence is Storage Lifecycle Policy'] = SplitLine[1]
			elif SplitLine[0] == 'Daily Windows':
				i+=1
				Daily_Windows = str()
				while not re.match('^$', EditLine):
					EditLine = (re.sub(r'\s\s\s\s\s\s+', '', DataList[i])).replace('\n', '')
					Daily_Windows += str(EditLine)+'\n'
					if not re.match('^$', EditLine):
						i+=1
				Daily_Windows = Daily_Windows[:-2]
						# print 'i++'
				Dict_Schedule['Schedule_Daily Windows'] = Daily_Windows
				# print Daily_Windows
			if not re.match('^$', EditLine):
				i+=1
		List_Schedule.append(Dict_Schedule)
	else:
		i+=1

num = 1
for index in Data:
	for i in index['Schedule']:
		WorkSheet.write(num, 0, index['Policy Name'])
		WorkSheet.write(num, 1, index['Policy Type'])
		WorkSheet.write(num, 2, index['Active'])
		try:
			WorkSheet.write(num, 3, index['Client Compress'])
		except:
			WorkSheet.write(num, 3, 'None')
		try:
			WorkSheet.write(num, 4, index['Follow NFS Mounts'])
		except:
			WorkSheet.write(num, 4, 'None')
		try:
			WorkSheet.write(num, 5, index['Cross Mount Points'])
		except:
			WorkSheet.write(num, 5, 'None')
		try:
			WorkSheet.write(num, 6, index['Collect TIR info'])
		except:
			WorkSheet.write(num, 6, 'None')
		try:
			WorkSheet.write(num, 7, index['Block Incremental'])
		except:
			WorkSheet.write(num, 7, 'None')
		try:
			WorkSheet.write(num, 8, index['Mult. Data Streams'])
		except:
			WorkSheet.write(num, 8, 'None')
		try:
			WorkSheet.write(num, 9, index['Client Encrypt'])
		except:
			WorkSheet.write(num, 9, 'None')
		try:
			WorkSheet.write(num, 10, index['Checkpoint'])
		except:
			WorkSheet.write(num, 10, 'None')
		try:
			WorkSheet.write(num, 11, index['Policy Priority'])
		except:
			WorkSheet.write(num, 11, 'None')
		try:
			WorkSheet.write(num, 12, index['Max Jobs/Policy'])
		except:
			WorkSheet.write(num, 12, 'None')
		try:
			WorkSheet.write(num, 13, index['Disaster Recovery'])
		except:
			WorkSheet.write(num, 13, 'None')
		try:
			WorkSheet.write(num, 14, index['Collect BMR info'])
		except:
			WorkSheet.write(num, 14, 'None')
		try:
			WorkSheet.write(num, 15, index['Residence'])
		except:
			WorkSheet.write(num, 15, 'None')
		try:
			WorkSheet.write(num, 16, index['Volume Pool'])
		except:
			WorkSheet.write(num, 16, 'None')
		try:
			WorkSheet.write(num, 17, index['Server Group'])
		except:
			WorkSheet.write(num, 17, 'None')
		try:
			WorkSheet.write(num, 18, index['Keyword'])
		except:
			WorkSheet.write(num, 18, 'None')
		try:
			WorkSheet.write(num, 19, index['Data Classification'])
		except:
			WorkSheet.write(num, 19, 'None')
		try:
			WorkSheet.write(num, 20, index['Residence is Storage Lifecycle Policy'])
		except:
			WorkSheet.write(num, 20, 'None')
		try:
			WorkSheet.write(num, 21, index['Application Discovery'])
		except:
			WorkSheet.write(num, 21, 'None')
		try:
			WorkSheet.write(num, 22, index['Discovery Lifetime'])
		except:
			WorkSheet.write(num, 22, 'None')
		try:
			WorkSheet.write(num, 23, index['ASC Application and attributes'])
		except:
			WorkSheet.write(num, 23, 'None')
		try:
			WorkSheet.write(num, 24, index['Granular Restore Info'])
		except:
			WorkSheet.write(num, 24, 'None')
		try:
			WorkSheet.write(num, 25, index['Ignore Client Direct'])
		except:
			WorkSheet.write(num, 25, 'None')
		try:
			WorkSheet.write(num, 26, index['Use Accelerator'])
		except:
			WorkSheet.write(num, 26, 'None')
		try:
			WorkSheet.write(num, 27, index['HW/OS/Client'])
		except:
			WorkSheet.write(num, 27, 'None')
		try:
			WorkSheet.write(num, 28, index['Include'])
		except:
			WorkSheet.write(num, 28, 'None')
		try:
			WorkSheet.write(num, 29, i['Schedule_Type'])
		except:
			WorkSheet.write(num, 29, 'None')
		try:
			WorkSheet.write(num, 30, i['Schedule_Frequency'])
		except:
			WorkSheet.write(num, 30, 'None')
		try:
			WorkSheet.write(num, 31, i['Schedule_Synthetic'])
		except:
			WorkSheet.write(num, 31, 'None')
		try:
			WorkSheet.write(num, 32, i['Schedule_Checksum Change Detection'])
		except:
			WorkSheet.write(num, 32, 'None')
		try:
			WorkSheet.write(num, 33, i['Schedule_PFI Recovery'])
		except:
			WorkSheet.write(num, 33, 'None')
		try:
			WorkSheet.write(num, 34, i['Schedule_Maximum MPX'])
		except:
			WorkSheet.write(num, 34, 'None')
		try:
			WorkSheet.write(num, 35, i['Schedule_Retention Level'])
		except:
			WorkSheet.write(num, 35, 'None')
		try:
			WorkSheet.write(num, 36, i['Schedule_Number Copies'])
		except:
			WorkSheet.write(num, 36, 'None')
		try:
			WorkSheet.write(num, 37, i['Schedule_Fail on Error'])
		except:
			WorkSheet.write(num, 37, 'None')
		try:
			WorkSheet.write(num, 38, i['Schedule_Residence'])
		except:
			WorkSheet.write(num, 38, 'None')
		try:
			WorkSheet.write(num, 39, i['Schedule_Volume Pool'])
		except:
			WorkSheet.write(num, 39, 'None')
		try:
			WorkSheet.write(num, 40, i['Schedule_Server Group'])
		except:
			WorkSheet.write(num, 40, 'None')
		try:
			WorkSheet.write(num, 41, i['Schedule_Residence is Storage Lifecycle Policy'])
		except:
			WorkSheet.write(num, 41, 'None')
		try:
			# test = i['Schedule_Daily Windows'].encode('UTF-8')
			WorkSheet.write(num, 42, i['Schedule_Daily Windows'])
		except:
			WorkSheet.write(num, 42, 'None')
			print '1'
		num += 1



# WorkSheet.write(10, 10, test)
WorkBook.save(ExcelFile)