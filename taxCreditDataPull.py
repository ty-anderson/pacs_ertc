import os
import xlwings as xw
from xlwings.constants import AutoFillType
import pandas as pd
import shutil
import glob

"""PR TAX CREDIT DATA PULL.  PULL PAYROLL REPORT FROM SHARED DRIVE, PUT INTO TEMPLATE"""


def getPayrollAllocationFiles():
    """Walks through all files in the path and copys to the end path (step 1)"""
    lookfolder = r"P:\PACS\Finance\Labor Data\2021"
    endfolder = r"C:\Users\tyler.anderson\Documents\Finance\210712 - 2021 Q2 ERTC\PAYROLL ALLOCATION REPORTS"
    for dirpath, dirnames, filenames in os.walk(lookfolder):
        for filename in [f for f in filenames if f.endswith(".xlsx")]:
            file = os.path.join(dirpath, filename)
            if 'Payroll Allocation Report v' in file:
                shutil.copy(file, endfolder + "/" + filename)


def copyIntoExcelTemplate():
    file_list = glob.glob(r"C:\Users\tyler.anderson\Documents\Finance\210712 - 2021 Q2 ERTC\PAYROLL ALLOCATION REPORTS\*.xlsx")
    counter = 122
    for file in file_list:
        filepath, filename = os.path.split(file)
        daterange = filename.split("(")
        daterange = daterange[1].split(")")
        wb = xw.Book(file, update_links=False, read_only=True)
        payroll_sht = wb.sheets[0]
        payroll_data = payroll_sht.range('A6:PX15000').value
        df = pd.DataFrame(payroll_data)
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
        client_names = df["Client ID"].to_list()
        for idx, val in enumerate(client_names):  # GET LAST ROW NUMBER TO AUTOFILL TO
            if val is None:
                break
        temp_wb = xw.Book(r"C:\Users\tyler.anderson\Documents\Finance\210712 - 2021 Q2 ERTC\210311 PACS Template.xlsx", update_links=True, read_only=False)
        ws = temp_wb.sheets['Payroll Allocation Report v3']
        ws.range("A2:PX15001").clear_contents()
        ws.range("A2:PX15001").value = payroll_data
        main_ws = temp_wb.sheets['CAREs Act Data Payroll Template']
        main_ws.range('A2:AY2').api.AutoFill(main_ws.range("A2:AY" + str(idx + 1)).api, AutoFillType.xlFillDefault)
        new_name = r"C:\Users\tyler.anderson\Documents\Finance\210712 - 2021 Q2 ERTC\FINISHED TEMPLATES" + "\\Payroll " + daterange[0] + " (" + str(counter) + ").xlsx"
        temp_wb.save(new_name)
        xw.apps.active.quit()
        counter += 1
        print(counter)


if __name__ == '__main__':
    # getPayrollAllocationFiles()
    copyIntoExcelTemplate()
