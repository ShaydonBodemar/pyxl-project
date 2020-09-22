# Module Name: read_excel.py
# Author: Shaydon Bodemar
# Date: 21 September 2020
# Overview: This module will be primarily responsible for inputting the data from the excel sheet for all model parameters.

from openpyxl import load_workbook

class ReadConfig:
    def __init__(self, excel_filename):
        workbook = load_workbook(excel_filename, read_only=True)
        chip_profiles = workbook['Chip Roadmap']

        # Reads in the data labels for the first row to be used later as keys in dict
        self.metadata_fields = []
        for cell in chip_profiles[1]:
            self.metadata_fields.append(cell.value)

        # Reads in all data fields to create a full nested dictionary representing each chip's metadata
        self.chip_profiles = {}
        for row in chip_profiles.iter_rows(min_row=2):
            chip_name = row[0].value
            self.chip_profiles[chip_name] = {}
            for cell in range(1,len(row)):
                self.chip_profiles[chip_name][self.metadata_fields[cell]] = row[cell].value


        workbook.close()

    def ReadChipMetaData:
        # TODO: reading of the metadata for each chip

    def ReadChipTierData:
        # TODO: read in TO-based data for every tier of chips