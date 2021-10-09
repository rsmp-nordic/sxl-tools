#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Requires xlsxwriter
import xlsxwriter

workbook = xlsxwriter.Workbook('SXL-template.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Test')

workbook.close()
