import pandas as pd
import numpy as np 
import os

class ExlSheet:

   def __init__(self, excelFileName, excelSheetName):
     self.excelFileName = excelFileName
     self.excelSheetName = excelSheetName
     self.pathExcelName = os.path.relpath(f'spreadsheets/{self.excelFileName}.xlsx')

   def getSheetInstance(self):
     sinst = pd.read_excel(self.pathExcelName, sheet_name=self.excelSheetName)
     return sinst


