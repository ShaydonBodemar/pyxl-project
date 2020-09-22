# Module Name: read_excel.py
# Author: Shaydon Bodemar
# Date: 21 September 2020
# Overview: This module will be primarily responsible for inputting the data from the excel sheet for all model parameters.

from openpyxl import load_workbook

class ReadConfig:
    """
    @brief Constructor for this module
    @param excel_filename Used to determine the file that will be interacted with.
    @note Only opens the file to be used and initializes the member variables.
    """
    def __init__(self, excel_filename):
        self._workbook = load_workbook(excel_filename, read_only=True)
        self._metadata_fields = []
        self._chip_profiles = {}
    

    """
    @brief Destructor for this module
    @note Closes workbook used for reading sheets in this module.
    """
    def __del__(self):
        self._workbook.close()


    """
    @brief Reads the metadata for each chip (scaler, TO date, etc)
    @param worksheet_name Name of the worksheet to be used for reading in appropriate data
    @note Wipes any metadata currently stored
    """
    def ReadChipMetaData(self, worksheet_name):
        chip_profiles = self._workbook[worksheet_name]

        # Reads in the data labels for the first row to be used later as keys in dict
        self._metadata_fields = []
        for cell in chip_profiles[1]:
            self._metadata_fields.append(cell.value)

        # Reads in all data fields to create a full nested dictionary representing each chip's metadata
        self._chip_profiles = {}
        for row in chip_profiles.iter_rows(min_row=2):
            chip_name = row[0].value
            self._chip_profiles[chip_name] = {}
            for cell in range(1,len(row)):
                self._chip_profiles[chip_name][self._metadata_fields[cell]] = row[cell].value
        # print(self._chip_profiles)


    """
    @brief Reads the tool usage profiles for each chip tier
    @param worksheet_name Name of the worksheet to be used for reading in appropriate data
    """
    def ReadChipTierData(self, worksheet_name):
        # TODO: read in TO-based data for every tier of chips
        print(1)