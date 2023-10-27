from openpyxl import *
from openpyxl.utils import *

def copyFormat(ws, src_row, src_col, dest_row, dest_col):
    src_cell = ws.cell(row=src_row, column=src_col)
    dest_cell = ws.cell(row=dest_row, column=dest_col)
    
    dest_cell.font = src_cell.font.copy()
    dest_cell.border = src_cell.border.copy()
    dest_cell.fill = src_cell.fill.copy()
    dest_cell.number_format = src_cell.number_format
    dest_cell.protection = src_cell.protection.copy()
    dest_cell.alignment = src_cell.alignment.copy()

def nextMonthInterest(endingBal):

    nextMonthInterestVal = round(endingBal * (0.10 / 12), 2)

    return nextMonthInterestVal