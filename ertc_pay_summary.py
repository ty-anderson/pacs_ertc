import os
from glob import glob
import xlwings as xw
import pandas as pd

main_file = xw.Book(r"C:\Users\tyler.anderson\Desktop\210712 - 2021 Q1 ERTC Summary.xlsx")
main_ws = main_file.sheets("PR Raw Data")


for file in glob(r"C:\Users\tyler.anderson\Documents\Finance\ETRC\211118 - 2021 Q3 ERTC\FINISHED TEMPLATES\*"):
    print(file)
    main_last_row = main_ws.range('A' + str(main_ws.cells.last_cell.row)).end('up').row
    wb = xw.Book(file)
    ws = wb.sheets("CAREs Act Data Payroll Template")
    last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    rng = ws.range(f"A2:AY{last_row}").options(ndim=2).value

    # PASTE IN DATA
    main_ws.range(f"A{main_last_row+1}:AY{main_last_row + last_row}").value = rng
    wb.close()
