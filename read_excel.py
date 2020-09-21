# Module Name: read_excel.py
# Author: Shaydon Bodemar
# Date: 21 September 2020
# Overview: This module will be primarily responsible for inputting the data from the excel sheet for all model parameters.

from openpyxl import load_workbook

# constructor for 
class ReadConfig:
    def __init__(self, excel_filename):
        wb = load_workbook(excel_filename, read_only=True)
        ws = wb['Chip Roadmap']

        for col in ws.columns:
            for cell in col:
                print(cell.value)

        wb.close()
