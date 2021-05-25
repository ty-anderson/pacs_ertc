import os
import xlwings as xw
from xlwings.constants import AutoFillType
import pandas as pd
# import win32com
# import shutil


# print(win32com.__gen_path__)
#
# try:
#     shutil.rmtree(win32com.__gen_path__[:-4])
# except:
#     pass

"""PR TAX CREDIT PT3.  COPY FILES TO TEMPLATE AND SAVE TEMPLATE"""


def getFileList():
    """Walks through all files in the path and makes a list of them"""
    folder = r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\PAYROLL ALLOCATION REPORTS\ALL REPORTS"
    file_list = []
    counter = 0
    for dirpath, dirnames, filenames in os.walk(folder):
        for filename in [f for f in filenames if f.endswith(".csv")]:
            file = os.path.join(dirpath, filename)
            if 'payroll allocation' in file:
                file_list.append(file)
                print(file)
                counter += 1
    print(counter)
    """PULL INFORMATION FROM BUDGET FILES"""
    getExcelData(file_list, counter)


def getExcelData(file_list, counter):
    for file in file_list:
        filepath, filename = os.path.split(file)
        wb = xw.Book(file, update_links=False, read_only=True)
        payroll_sht = wb.sheets[0]
        payroll_data = payroll_sht.range('A1:PX15000').value
        df = pd.DataFrame(payroll_data)
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
        client_names = df["Client ID"].to_list()
        for idx, val in enumerate(client_names):
            if val is None:
                break
        temp_wb = xw.Book(r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\210311 PACS Template.xlsx", update_links=True, read_only=False)
        ws = temp_wb.sheets['Payroll Allocation Report v3']
        ws.range("A2:PX15001").clear_contents()
        ws.range("A2:PX15001").value = payroll_data
        main_ws = temp_wb.sheets['CAREs Act Data Payroll Template']
        main_ws.range('A2:AY2').api.AutoFill(main_ws.range("A2:AY" + str(idx + 1)).api, AutoFillType.xlFillDefault)
        # main_ws.range("A2:AY" + str(idx + 1)).value = main_ws.range("A2:AY" + str(idx + 1)).value  # SAVE AS VALUES
        # for sht in temp_wb.sheets:
        #     if sht.name != "CAREs Act Data Payroll Template":
        #         sht.delete()
        filename = filename.upper()
        new_name = r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\FINISHED TEMPLATES" + "\\" + filename[:13] + " (" + str(counter) + ").xlsx"
        # new_name = r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\FINISHED TEMPLATES\test.xlsx"
        # temp_wb.save()
        temp_wb.save(new_name)
        xw.apps.active.quit()
        counter -= 1
        print(counter)


if __name__ == '__main__':
    getFileList()
