import os
import xlwings as xw
import pandas as pd

"""PR TAX CREDIT PT2.  PULL PAYROLL REPORT FROM FILE CREATED FROM PT1, REMOVE SM OR BW FACILITIES BASED ON ID # AND SAVE"""

# ADD A TEST COMMENT

# CREATE REFERENCE LIST OF WHICH BUILDINGS ARE SM OR BW
payroll_lookup = pd.read_excel(r"P:\PACS\Finance\General Info\Finance Misc\Facility List.xlsx", sheet_name="PR Ref", usecols=["ID", "Schedule"])
print(payroll_lookup)
count = 0
for i in payroll_lookup['Schedule']:
    if i == 'Semi-Monthly':
        payroll_lookup = payroll_lookup.drop(index=count)
    count += 1
sm_list = payroll_lookup["ID"].to_list()


def getExcelDataBW(file_list, counter):
    """Will pull name and net income from budget files and write to csv file on desktop"""
    for file in file_list:
        filepath, filename = os.path.split(file)
        filename = filename[:14] + "payroll allocation report"
        wb = xw.Book(file, update_links=False, read_only=True)
        payroll_sht = wb.sheets[0]
        payroll_data = payroll_sht.range('A1:PX15000').value
        df = pd.DataFrame(payroll_data)
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
        client_names = df["Client ID"].to_list()
        # LOOP THROUGH FACILITY LIST COLUMN TO REMOVE SEMI MONTHLY'S
        for idx, val in enumerate(client_names, start=1):
            val = str(val)
            if val is None:
                break
            for j in sm_list:
                if j.lower() in val.lower():
                    df = df.drop(idx)
                    print("drop " + val)
                    break
        new_name = r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\PAYROLL ALLOCATION REPORTS\ALL REPORTS" + "\\" + filename + " - count " + str(counter) + ".csv"
        df.to_csv(new_name, index=False, header=True)
        xw.apps.active.quit()
        counter -= 1
        print(counter)


def getFileList():
    """Walks through all files in the path and makes a list of them"""
    folder = r"C:\Users\tyler.anderson\Documents\Finance\210308 - 2020 Tax Credit - Payroll\PAYROLL ALLOCATION REPORTS\BW2"
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
    getExcelDataBW(file_list, counter)


if __name__ == '__main__':
    getFileList()
