#!/usr/bin/env python
# coding: utf-8

# In[8]:


import xlwings as xw
import os
from googletrans import Translator
from pathlib import Path

# Translator
translator = Translator()
translator = Translator(service_urls=['translate.googleapis.com'])
    
# Opening the Excel file
paths = Path(input('Choose your file path:'))

language = {"English":'en', "Slovak":'sk'}
ch_lang = input('Choose your language, English or Slovak?')

wb = xw.Book(paths)
sheet = wb.sheets[5]

# Unhiding hidden row 3
sheet.api.Rows(3).Hidden = False  # Unhide row 3

last_column = sheet.api.Cells(3, sheet.api.Columns.Count).End(-4159).Column  # Find the last column with data
sheet.api.Range(sheet.api.Cells(3, 1), sheet.api.Cells(3, last_column)).AutoFilter(1)  # Set filter on row 3

# Iterating through cells in column A
for cell in sheet.range('A4:A10000'):  # Define the range where comments are
    if cell.value is not None:
        if cell.api.Comment:  # Check if there is a comment in the cell
            comment_text = cell.api.Comment.Text()  # Read the comment text
            # Translating the comment
            translated_comment = translator.translate(comment_text, dest=language[ch_lang])  # 'en' for English, can be changed
            # Paste the translated comment in column B
            sheet.range(f'B{cell.row}').value = translated_comment.text
        else:
            # Paste "no comment" in column B if there is no comment
            sheet.range(f'B{cell.row}').value = "no comment"
    else:
        break

# Freezing column G
wb = xw.books.active  # Active workbook
active_window = wb.app.api.ActiveWindow  # Active window
active_window.FreezePanes = False  # Disable freezing
active_window.SplitColumn = 7  # Set the split to column H (which freezes column G)
active_window.SplitRow = 0  # Do not freeze any rows
active_window.FreezePanes = True  # Enable freezing

# wb.save()
# wb.close()

