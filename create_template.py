#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlsxwriter
import argparse

parser = argparse.ArgumentParser(description='Create SXL template')
parser.add_argument('--no-grouped-objects', default=15, type=int,
    help='Number of grouped objects')
parser.add_argument('--no-single-objects', default=29, type=int,
    help='Number of single objects')
parser.add_argument('--no-command-functional-position', default=3, type=int,
    help='Number of commands of type functional position')
parser.add_argument('--no-command-functional-state', default=3, type=int,
    help='Number of commands of type functional state')
parser.add_argument('--no-command-maneuver', default=3, type=int,
    help='Number of commands of type maneuver')
parser.add_argument('--no-command-parameters', default=3, type=int,
    help='Number of commands of type parameter')
parser.add_argument('--alarm-rvs', default=2, type=int,
    help='Number of alarm return values')
parser.add_argument('--status-rvs', default=2, type=int,
    help='Number of status return values')
parser.add_argument('--command-args', default=2, type=int,
    help='Number of command arguments')
parser.add_argument('--output',
    default='RSMP_Template_SignalExchangeList.xlsx',
    help='Output filename')
args = parser.parse_args()

workbook = xlsxwriter.Workbook(args.output)

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

go_row=6 # start row
for col in range(0,6):
    for row in range(go_row,go_row+args.no_grouped_objects):
        worksheet.write(row, col, "", t9_box)

so_row=go_row+args.no_grouped_objects # start row

worksheet.write(so_row, 0, 'Single objects', t9b_l)
worksheet.write(so_row+1, 0, 'ObjectType', t9b_l_i_box)
worksheet.write(so_row+1, 1, 'Object', t9b_l_i_box)
worksheet.write(so_row+1, 2, 'componentId', t9b_l_i_box)
worksheet.write(so_row+1, 3, 'NTSObjectId', t9b_l_i_box)
worksheet.write(so_row+1, 4, 'externalNtsId', t9b_l_i_box)
worksheet.write(so_row+1, 5, 'Description', t9b_l_i_box)

for col in range(0,6):
    for row in range(so_row+2,so_row+2+args.no_single_objects):
        worksheet.write(row, col, "", t9_box)

# Write comments
worksheet.write_comment('A1',
  "This tab contains info about all grouped objects and their single objects."+
  "This tab should be exported to RSMP Simulators(s) as a CSV file, \n"+
  "preferred name ''SiteId.CSV'', ex ''AB_26507_881.CSV''."+
  "If there are multiple SiteId's for one plant (ex congestion tax) \n"+
  "just add more Objects tabs and export them as multiple CSV files.")
worksheet.write_comment('B2',
  "SiteId(s) are always sent in the first RSMP packet.\n"+
  "The communication partners could/should use this information to validate\n"+
  "they actually are communicating with the correct plant.")

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

worksheet.write('A7', 'Plant', t9_box)
worksheet.write('B7', 'See state-bit definitions below', t9_box)

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

# Write comments
worksheet.write_comment('A1',
  "This tab should not be exported to the RSMP Simulator(s), " +
  "they do not use the information anyway. " +
  "Aggregated Status is automatically created for each Grouped Object.")

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
return_value = [
    'Name',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,args.alarm_rvs):
    worksheet.merge_range(4, col, 4, col+3, 'return value', t9b_c_i_box)
    for item in (return_value):
        worksheet.write(5, col, item, t9b_l_i_box)
        worksheet.set_column(col, col, 10.13)
        for row in range(6,26):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Write comments
worksheet.write_comment('A1',
  "This tab contains info about all alarms.\n"+
  "This tab should be exported to RSMP Simulators(s) as a CSV file,\n"+
  "preferred name ''Alarms.CSV''\n\n"+
  "Simulator(s) will create alarms using ObjectType."+
  "If alarm exist at a specific object only,"+
  "specify it using the optional Object column."+
  "Multiple return values for each AlarmCodeId could be specified,"+
  "just add more return value column groups.")



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
return_value = [
    'Name',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,args.status_rvs):
    worksheet.merge_range(4, col, 4, col+3, 'return value', t9b_c_i_box)
    for item in (return_value):
        worksheet.write(5, col, item, t9b_l_i_box)
        worksheet.set_column(col, col, 10.13)
        for row in range(6,52):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Write comments
worksheet.write_comment('A1',
  "This tab contains info about status values, ex Speed, Time, NOx etc."+
  "Type could be string, boolean, integer, float, base64 etc."+
  "This tab should be exported to RSMP Simulators(s) as a CSV file,"+
  "preferred name ''Status.CSV''"+
  "Simulator(s) will create status values using ObjectType."+
  "If status exist at a specific object only,"+
  "specify it using the optional Object column."+
  "Multiple return values for each StatusCodeId could be specified,"+
  "just add more return value column groups.")

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 32.13)
worksheet.set_column(2, 2, 16.13)
worksheet.set_column(3, 3, 32.13)
col = 4
for num in range(0,args.status_rvs):
    worksheet.set_column(col+(4*num), col+(4*num)+2, 8)
    worksheet.set_column(col+(4*num)+3, col+(4*num)+3, 16.13)

# Commands
worksheet = workbook.add_worksheet('Commands')
worksheet.write('A1', 'Commands per object type', t18b)
worksheet.write('A3', 'Revision date:', t9b_r)
worksheet.write('B3', 'yyyy-mm-dd', t9_c)
worksheet.write('E4', 'Obs! Leading ''-'' should not exist in protocol level', t9)

worksheet.write('A5', 'Functional position', t9b_l)

title = [
    'ObjectType',
    'Object (optional)',
    'commandCodeId',
    'Description',
]

col = 0
fp_row = 6 # start row
for item in (title):
    worksheet.write(5, col, item, t9b_l_i_box)
    for row in range(fp_row,fp_row+args.no_command_functional_position):
        worksheet.write(row, col, "", t9_box)
    col += 1

col = 4
argument = [
    'Name',
    'Command',
    'Type',
    'Value',
    'Comment'
]
for num in range(0,args.command_args):
    worksheet.merge_range(fp_row-2, col, fp_row-2, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(fp_row-1, col, item, t9b_l_i_box)
        for row in range(fp_row,fp_row+args.no_command_functional_position):
            worksheet.write(row, col, "", t9_box)
        col += 1

fs_row = fp_row + args.no_command_functional_position + 3

worksheet.write(fs_row-2, 0, 'Functional state', t9b_l)
col = 0
for item in (title):
    worksheet.write(fs_row-1, col, item, t9b_l_i_box)
    for row in range(fs_row,fs_row+args.no_command_functional_state):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,args.command_args):
    worksheet.merge_range(fs_row-2, col, fs_row-2, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(fs_row-1, col, item, t9b_l_i_box)
        for row in range(fs_row,fs_row+args.no_command_functional_state):
            worksheet.write(row, col, "", t9_box)
        col += 1

m_row = fs_row + args.no_command_functional_state + 3

worksheet.write(m_row-2, 0, 'Manouver', t9b_l)
col = 0
for item in (title):
    worksheet.write(m_row-1, col, item, t9b_l_i_box)
    for row in range(m_row,m_row+args.no_command_maneuver):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,args.command_args):
    worksheet.merge_range(m_row-2, col, m_row-2, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(m_row-1, col, item, t9b_l_i_box)
        for row in range(m_row,m_row+args.no_command_maneuver):
            worksheet.write(row, col, "", t9_box)
        col += 1

p_row = m_row + args.no_command_maneuver + 3

worksheet.write(p_row-2, 0, 'Parameter', t9b_l)
col = 0
for item in (title):
    worksheet.write(p_row-1, col, item, t9b_l_i_box)
    for row in range(p_row,p_row+args.no_command_parameters):
        worksheet.write(row, col, "", t9_box)
    col += 1
col = 4
for num in range(0,args.command_args):
    worksheet.merge_range(p_row-2, col, p_row-2, col+4, 'argument', t9b_c_i_box)
    for item in (argument):
        worksheet.write(p_row-1, col, item, t9b_l_i_box)
        for row in range(p_row,p_row+args.no_command_parameters):
            worksheet.write(row, col, "", t9_box)
        col += 1

# Write comments
worksheet.write_comment('A1',
  "This tab contains info about commands,"+
  "ex Barriers, Fans, Information sign etc."+
  "Type could be string, boolean, integer, float, base64 etc."+
  "This tab should be exported to RSMP Simulators(s) as a CSV file,"+
  "preferred name ''Commands.CSV''"+
  "Simulator(s) will create commands using ObjectType."+
  "If command exist at a specific object only,"+
  "specify it using the optional Object column."+
  "Multiple command arguments for each CommandCodeId could be specified,"+
  "just add more argument column groups.")

# Adjust widths
worksheet.set_column(0, 0, 32.13)
worksheet.set_column(1, 1, 32.13)
worksheet.set_column(2, 2, 16.13)
worksheet.set_column(3, 3, 32.13)
col = 4
for num in range(0,args.command_args):
    worksheet.set_column(col+(5*num), col+(5*num)+3, 8)
    worksheet.set_column(col+(5*num)+4, col+(5*num)+4, 16.13)

workbook.close()
