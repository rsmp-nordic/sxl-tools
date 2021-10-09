#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Requires xlsxwriter
import xlsxwriter

workbook = xlsxwriter.Workbook('SXL-template.xlsx')

# Version
worksheet = workbook.add_worksheet('Version')
worksheet.write('B2, 'Signal Exchange List')
worksheet.write('B4, 'Plant id')
worksheet.write('B6, 'Plant name')
worksheet.write('B8, 'Work documentation')
worksheet.write('A10, 'Constructor:')
worksheet.write('A12, 'Reviewed:')
worksheet.write('A15, 'Approved:')
worksheet.write('A18, 'Created date:')
worksheet.write('A20, 'SXL revision:')
worksheet.write('B20, 'Revision number')
worksheet.write('C20, 'Revision date')
worksheet.write('B21, '1.0')
worksheet.write('C21, 'yyyy-mm-dd')
worksheet.write('A26, 'RSMP version:')
worksheet.write('B26, '3.1.1')

# Object types
worksheet = workbook.add_worksheet('Object types')
worksheet.write('A1, 'Object types')
worksheet.write('A3, 'Revision date:')
worksheet.write('B3, 'yyyy-mm-dd')
worksheet.write('A5, 'Grouped object types')
worksheet.write('A6, 'ObjectType')
worksheet.write('A6, 'Description/comment')
worksheet.write('A17, 'Single object types')
worksheet.write('A18, 'ObjectType')
worksheet.write('B18, 'Description/comment')

# Objects
worksheet = workbook.add_worksheet('Objects')
worksheet.write('A1, 'Site objects')
worksheet.write('A2, 'Siteid:')
worksheet.write('B2, 'siteid')
worksheet.write('B3, 'description')
worksheet.write('A3, 'Revision date:')
worksheet.write('B3, 'yyyy-mm-dd')
worksheet.write('A5, 'Grouped objects')
worksheet.write('A6, 'ObjectType')
worksheet.write('B6, 'Object')
worksheet.write('C6, 'componentId')
worksheet.write('D6, 'NTSObjectId')
worksheet.write('E6, 'externalNtsId')
worksheet.write('F6, 'Description')
worksheet.write('A23, 'Single objects')
worksheet.write('A24, 'ObjectType')
worksheet.write('B24, 'Object')
worksheet.write('C24, 'componentId')
worksheet.write('D24, 'NTSObjectId')
worksheet.write('E24, 'externalNtsId')
worksheet.write('F24, 'Description')

# Aggregated status
worksheet = workbook.add_worksheet('Aggregated status')
worksheet.write('A1, 'Aggregated status per grouped object')
worksheet.write('A3, 'Revision date:')
worksheet.write('B3, 'yyyy-mm-dd')
worksheet.write('C5, 'Obs! Leading '-' should not exist in protocol level')
worksheet.write('A6, 'ObjectType')
worksheet.write('B6, 'State')
worksheet.write('C6, 'funcationPosition')
worksheet.write('D6, 'funcationState')

workbook.close()
