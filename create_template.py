#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Requires xlsxwriter
import xlsxwriter

workbook = xlsxwriter.Workbook('RSMP_Template_SignalExchangeList.xlsx')

# Formatting
t32b_c = workbook.add_format({'font_name':'Arial', 'font_size':'32',
    'bold':True, 'align':'center'})
t18b = workbook.add_format({'font_name':'Arial', 'font_size':'18',
    'bold':True})
t18b_c = workbook.add_format({'font_name':'Arial', 'font_size':'18',
    'bold':True, 'align':'center'})
t18_i = workbook.add_format({'font_name':'Arial', 'font_size':'18',
    'italic':True})
t18_c_i = workbook.add_format({'font_name':'Arial', 'font_size':'18',
    'align':'center', 'italic':True})
t9b_r = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':True, 'align':'right'})
t9b_l = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':True, 'align':'left'})
t9b_l_i_box = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':True, 'align':'left', 'italic':True,'border':1, 'bottom':2})
t9b_c_i_box = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':True, 'align':'center', 'italic':True,'border':1})
t9b_c_box = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':True, 'align':'center', 'border':1})
t9 = workbook.add_format({'font_name':'Arial', 'font_size':'9'})
t9_c = workbook.add_format({'font_name':'Arial', 'font_size':'9', 
    'bold':False, 'align':'center'})
t9_c_box = workbook.add_format({'font_name':'Arial', 'font_size':'9', 
    'bold':False, 'align':'center', 'border':1})
t9_box = workbook.add_format({'font_name':'Arial', 'font_size':'9',
    'bold':False, 'border':1})

# Version
worksheet = workbook.add_worksheet('Version')
worksheet.write('B2', 'Signal Exchange List', t18b_c)
worksheet.write('B4', 'Plant id', t32b_c)
worksheet.write('B6', 'Plant name', t18b_c)
worksheet.write('B8', 'Work documentation', t18b_c)
worksheet.write('A10', 'Constructor:', t9b_r)
worksheet.write('B10', '', t9_box)
worksheet.write('A12', 'Reviewed:', t9b_r)
worksheet.write('B12', '', t9_box)
worksheet.write('B13', '', t9_box)
worksheet.write('A15', 'Approved:', t9b_r)
worksheet.write('B15', '', t9_box)
worksheet.write('B16', '', t9_box)
worksheet.write('A18', 'Created date:', t9b_r)
worksheet.write('B18', 'yyyy-mm-dd', t9_c_box)
worksheet.write('A20', 'SXL revision:', t9b_r)
worksheet.write('B20', 'Revision number', t9b_c_box)
worksheet.write('C20', 'Revision date', t9b_c_box)
worksheet.write('B21', '1.0', t9_c_box)
worksheet.write('C21', 'yyyy-mm-dd', t9_c_box)
worksheet.write('B22', '', t9_box)
worksheet.write('C22', '', t9_box)
worksheet.write('B23', '', t9_box)
worksheet.write('C23', '', t9_box)
worksheet.write('B24', '', t9_box)
worksheet.write('C24', '', t9_box)
worksheet.write('A26', 'RSMP version:', t9b_r)
worksheet.write('B26', '3.1.1', t9_c_box)

# Write comments
worksheet.write_comment('B2',
  "This tab contains info about the plant itself and some revision history.\n" +
  "\nThis tab should be exported to RSMP Simulators(s) as a CSV file, " +
  "preferred name is 'Version.CSV'.")
worksheet.write_comment('B20',
   "The last revision number is always sent in the first RSMP packet, " +
   "the communication partners should use this information to validate " +
   "they actually are communicating using same Signal Exchange List (SXL) " +
   "version.\n\nRSMP Simulator(s) will use the last row hence must be ordered " +
   "with increasing revision numbers. Version format is however not crucial.\n")
worksheet.write_comment('B26',
    "RSMP version is not used by the RSMP Simulator(s), but is here for " +
    "informational purposes.")

# Adjust widths
worksheet.set_column(0, 0, 21.5)
worksheet.set_column(1, 1, 30.63)
worksheet.set_column(2, 2, 18.63)

# Object types
worksheet = workbook.add_worksheet('Object types')
worksheet.write('A1', 'Object types', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('A5', 'Grouped object types', t9b_l)
worksheet.write('A6', 'ObjectType', t9b_l_i_box)
worksheet.write('B6', 'Description/comment', t9b_l_i_box)
for col in range(0,2):
    for row in range(6,14):
        worksheet.write(row, col, "", t9_box)
worksheet.write('A17', 'Single object types', t9b_l)
worksheet.write('A18', 'ObjectType', t9b_l_i_box)
worksheet.write('B18', 'Description/comment', t9b_l_i_box)
for col in range(0,2):
    for row in range(18,26):
        worksheet.write(row, col, "", t9_box)

# Write comments
worksheet.write_comment('A1',
  "This tab should not be exported to the RSMP Simulator(s), " +
  "they do not use the information anyway. " +
  "ObjectTypes are automatically extracted from Object tab(s)")

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 38.75)
worksheet.set_column(2, 2, 41.25)

# Objects
worksheet = workbook.add_worksheet('Objects')
worksheet.write('A1', 'Site objects', t18b)
worksheet.write('A2', 'Siteid:', t9b_r)
worksheet.write('B2', 'siteid', t18_c_i)
worksheet.write('C2', 'description', t18_i)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('A5', 'Grouped objects', t9b_l)
worksheet.write('A6', 'ObjectType', t9b_l_i_box)
worksheet.write('B6', 'Object', t9b_l_i_box)
worksheet.write('C6', 'componentId', t9b_l_i_box)
worksheet.write('D6', 'NTSObjectId', t9b_l_i_box)
worksheet.write('E6', 'externalNtsId', t9b_l_i_box)
worksheet.write('F6', 'Description', t9b_l_i_box)
for col in range(0,6):
    for row in range(6,21):
        worksheet.write(row, col, "", t9_box)

worksheet.write('A23', 'Single objects', t9b_l)
worksheet.write('A24', 'ObjectType', t9b_l_i_box)
worksheet.write('B24', 'Object', t9b_l_i_box)
worksheet.write('C24', 'componentId', t9b_l_i_box)
worksheet.write('D24', 'NTSObjectId', t9b_l_i_box)
worksheet.write('E24', 'externalNtsId', t9b_l_i_box)
worksheet.write('F24', 'Description', t9b_l_i_box)
for col in range(0,6):
    for row in range(24,53):
        worksheet.write(row, col, "", t9_box)

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 34.5)
worksheet.set_column(2, 2, 27.25)
worksheet.set_column(3, 3, 27.25)
worksheet.set_column(4, 4, 57.5)

# Aggregated status
worksheet = workbook.add_worksheet('Aggregated status')
worksheet.write('A1', 'Aggregated status per grouped object', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('C5', 'Obs! Leading ''-'' should not exist in protocol level', t9)

title = [
    'ObjectType',
    'State',
    'functionalPosition',
    'functionalState',
    'Description'
]

col = 0
for item in (title):
    worksheet.write(5, col, item, t9b_l_i_box)
    col += 1

title = ['State- Bit nr (12345678)', 'Description', 'Comment']
col = 0
for item in (title):
    worksheet.write(15, col, item, t9b_l_i_box)
    col += 1

for col in range(0,5):
    for row in range(6,13):
        worksheet.write(row, col, "", t9_box)

row = 16
bits = (
    ['1', 'Local mode'],
    ['2', 'No Communications'],
    ['3', 'High Priority Fault'],
    ['4', 'Medium Priority Fault'],
    ['5', 'Low Priority Fault'],
    ['6', 'Connected / Normal - In Use'],
    ['7', 'Connected / Normal - Idle'],
    ['8', 'Not Connected']
)

for bit, description in (bits):
    worksheet.write(row, 0, bit, t9_c_box)
    worksheet.write(row, 1, description, t9_box)
    worksheet.write(row, 2, "", t9_box)
    row += 1

# Adjust widths
worksheet.set_column(0, 0, 31.13)
worksheet.set_column(1, 1, 40.5)
worksheet.set_column(2, 2, 26.38)
worksheet.set_column(3, 3, 31.38)
worksheet.set_column(4, 4, 32.13)

# Alarms
worksheet = workbook.add_worksheet('Alarms')
worksheet.write('A1', 'Alarms per object type', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('I4', 'Obs! Leading ''-'' should not exist in protocol level', t9)

title = [
    'ObjectType',
    'Object (optional)',
    'alarmCodeId',
    'Description',
    'externalAlarmCodeId',
    'externalNtSAlarmCodeId',
    'Priority',
    'Category'
]

col = 0
for item in (title):
    worksheet.write(5, col, item, t9b_l_i_box)
    for row in range(6,26):
        worksheet.write(row, col, "", t9_box)
    col += 1

col = 8
num_return_values=2
return_value = [
    'Name',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,num_return_values):
    worksheet.merge_range(4, col, 4, col+3, 'return value', t9b_c_i_box)
    for item in (return_value):
        worksheet.write(5, col, item, t9b_l_i_box)
        worksheet.set_column(col, col, 10.13)
        for row in range(6,26):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 32.13)
worksheet.set_column(2, 2, 16.63)
worksheet.set_column(3, 3, 32.13)
worksheet.set_column(4, 4, 32.13)
worksheet.set_column(5, 5, 32.13)

# Status
worksheet = workbook.add_worksheet('Status')
worksheet.write('A1', 'Status per object type', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('E4', 'Obs! Leading ''-'' should not exist in protocol level', t9)

title = [
    'ObjectType',
    'Object (optional)',
    'statusCodeId',
    'Description',
]

col = 0
for item in (title):
    worksheet.write(5, col, item, t9b_l_i_box)
    for row in range(6,52):
        worksheet.write(row, col, "", t9_box)
    col += 1

col = 4
num_return_values=2
return_value = [
    'Name',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,num_return_values):
    worksheet.merge_range(4, col, 4, col+3, 'return value', t9b_c_i_box)
    for item in (return_value):
        worksheet.write(5, col, item, t9b_l_i_box)
        worksheet.set_column(col, col, 10.13)
        for row in range(6,52):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 32.13)
worksheet.set_column(2, 2, 16.13)
worksheet.set_column(3, 3, 32.13)
col = 4
for num in range(0,num_return_values):
    worksheet.set_column(col+(4*num), col+(4*num)+2, 8)
    worksheet.set_column(col+(4*num)+3, col+(4*num)+3, 16.13)

# Commands
worksheet = workbook.add_worksheet('Commands')
worksheet.write('A1', 'Commands per object type', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('E4', 'Obs! Leading ''-'' should not exist in protocol level', t9)

worksheet.write('A5', 'Functional position')

title = [
    'ObjectType',
    'Object (optional)',
    'commandCodeId',
    'Description',
]

col = 0
for item in (title):
    worksheet.write(5, col, item, t9b_l_i_box)
    for row in range(6,9):
        worksheet.write(row, col, "", t9_box)
    col += 1

col = 4
num_arguments=2
argument = [
    'Name',
    'Command',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,num_arguments):
    worksheet.merge_range(4, col, 4, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(5, col, item, t9b_l_i_box)
        for row in range(6,9):
            worksheet.write(row, col, "", t9_box)
        col += 1


worksheet.write('A11', 'Functional state')
col = 0
for item in (title):
    worksheet.write(11, col, item, t9b_l_i_box)
    for row in range(12,15):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,num_arguments):
    worksheet.merge_range(10, col, 10, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(11, col, item, t9b_l_i_box)
        for row in range(12,15):
            worksheet.write(row, col, "", t9_box)
        col += 1

worksheet.write('A17', 'Manouver')
col = 0
for item in (title):
    worksheet.write(17, col, item, t9b_l_i_box)
    for row in range(18,21):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,num_arguments):
    worksheet.merge_range(16, col, 16, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(17, col, item, t9b_l_i_box)
        for row in range(18,21):
            worksheet.write(row, col, "", t9_box)
        col += 1

worksheet.write('A23', 'Parameter')
col = 0
for item in (title):
    worksheet.write(23, col, item, t9b_l_i_box)
    for row in range(24,27):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,num_arguments):
    worksheet.merge_range(22, col, 22, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(23, col, item, t9b_l_i_box)
        for row in range(24,27):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 32.13)
worksheet.set_column(2, 2, 16.13)
worksheet.set_column(3, 3, 32.13)
col = 4
for num in range(0,num_arguments):
    worksheet.set_column(col+(5*num), col+(5*num)+3, 8)
    worksheet.set_column(col+(5*num)+4, col+(5*num)+4, 16.13)

workbook.close()
