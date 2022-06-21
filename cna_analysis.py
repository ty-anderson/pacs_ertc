from glob import glob
import pandas as pd
import xlwings as xw


# path = fr"C:\Users\tyler.anderson\Desktop\PACS Finance\ETRC\210308 - 2020 ERTC\FINISHED TEMPLATES\PPE * SM *.xlsx"
path = fr"C:\Users\tyler.anderson\Desktop\Clack Payroll Info\2022\*"

main_df = pd.DataFrame()
wb = xw.Book()

x=1
for file in glob(path):
    sht = xw.Book(file)
    ws = sht.sheets('Sheet1')
    l_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row

    wb.sheets[0].range(f'A{x}').value = ws.range(f'A1:AG{l_row}').value
    # ws.range("A1:BS196").copy(original.sheets(adj_name).range("A1:BS196"))
    # wb.sheets(0).range(f"A{x}").value = reporting_ws.range("A1:BS196").value
    x = x + l_row - 1
    sht.close()
    print('hold')


