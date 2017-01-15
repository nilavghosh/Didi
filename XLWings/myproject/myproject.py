import xlwings as xw
import numpy as np

def myfunction():
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = 'Hellos World!'
    sht = wb.sheets[0]

    # sht.range('C23:AF28').value = leaderboard
    sht.range('C23:AF28').value = sht.range('C5:AF10').value


