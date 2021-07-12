import os
import xlwings as xw
import pandas as pd
import glob


def copyPayrollRawData():
    """Walks through all files in the path and makes a list of them"""
    summary = r"C:\Users\tyler.anderson\Documents\Finance\210525 - 2021 Q1 ERTC\210712 - 2021 Q1 ERTC Summary.xlsx"
    folder = glob.glob(r"C:\Users\tyler.anderson\Documents\Finance\210525 - 2021 Q1 ERTC\FINISHED TEMPLATES\*")
    sum_df = pd.DataFrame(index="Client ID")
    for i, file in enumerate(folder):
        df = pd.read_excel(file, sheet_name="CAREs Act Data Payroll Template", index_col="Client I")
        # last_row = len(df)
        # wb = xw.Book(file)
        # sht = wb.sheets('CAREs Act Data Payroll Template')
        # rg = sht.range("A2:AY" + str(last_row))
        # vals = rg.value
        sum_df = sum_df.append(df)
        # xw.apps.active.quit()
        print(f"File {str(i)}")
    sum_df.to_csv()


if __name__ == '__main__':
    copyPayrollRawData()
    print("Done")