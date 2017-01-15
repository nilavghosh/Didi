import numpy as np
import xlwings as xw

def world():
    wb = xw.Workbook.caller()
    # wb.sheets[0].range('A1').value = 'Hello World!'
    xw
    xw.Range('A1').value = 100