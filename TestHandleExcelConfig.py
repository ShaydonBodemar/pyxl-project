# Module Name: TestHandleExcelConfig.py
# Author: Shaydon Bodemar
# Date: 22 September 2020
# Overview: This module will be responsible for testing and running ReadUsageData for demonstration purposes.


import HandleExcelConfig


def MakePlot():
    mp = HandleExcelConfig.ReadUsageData('Forecast_Model.xlsx')
    mp.ReadChipMetaData('Chip Roadmap')
    mp.ReadChipTierData('Profiles')
    mp.PlotData()


if __name__ == '__main__':
    MakePlot()