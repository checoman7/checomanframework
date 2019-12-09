import pandas as pd
import xlwings as xw
import os

class ExlBinarySheet():

 def __init__(self, excelFileName, excelSheetName):
    self.excelFileName = excelFileName
    self.excelSheetName = excelSheetName
    self.pathExcelName = os.path.relpath(f'spreadsheets/{excelFileName}.xlsb')
    self.book = xw.Book(self.pathExcelName)
    self.sheet = self.book.sheets(self.excelSheetName)

 def getSheet(self, datarange):
    df = self.sheet.range(datarange).options(pd.DataFrame, expand='table').value
    self.book.app.kill()
    #xw.apps.active.kill()
    return df

 