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


workbook.close()
