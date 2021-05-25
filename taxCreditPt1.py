import os
import xlwings as xw
import pandas as pd

"""PR TAX CREDIT PT1.  PULL PAYROLL REPORT FROM FILE IN FOLDER, REMOVE SM OR BW FACILITIES AND SAVE"""

payroll_lookup = pd.read_excel(r"P:\PACS\Finance\General Info\Finance Misc\Facility List.xlsx", usecols=["Common Name", "Payroll"])
print(payroll_lookup)
count = 0
for i in payroll_lookup['Payroll']:
    if i == 'Semi-Monthly':
        payroll_lookup = payroll_lookup.drop(index=count)
    count += 1
sm_list = payroll_lookup["Common Name"].to_list()


def getExcelDataBW(file_list, counter):
    """Pull reports and filter based on Bi-weekly or Semi-Monthly"""
    for file in file_list:
        filepath, filename = os.path.split(file)
        filename = filename[:10] + " SM payroll allocation report"
        wb = xw.Book(file, update_links=False, read_only=True)
        try:
            payroll_sht = wb.sheets['Data - Raw Data File']
            payroll_data = payroll_sht.range('E6:PX15000').value
        except:
            payroll_sht = wb.sheets['Sheet1']
            payroll_data = payroll_sht.range('A4:PX15000').value
        df = pd.DataFrame(payroll_data)
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
        client_names = df["Client Name"].to_list()
        # LOOP THROUGH FACILITY LIST COLUMN TO REMOVE SEMI MONTHLY'S
        for idx, val in enumerate(client_names, start=1):
            if val is None:
                break
            for j in sm_list:
                if j.lower() in val.lower():
                    df = df.drop(idx)
                    print("drop " + str(val))
                    break
        new_name = r"C:\Users\tyler.anderson\Desktop\Payroll Reports Finished" + "\\" + filename + " - count " + str(counter) + ".csv"
        df.to_csv(new_name, index=False, header=True)
        xw.apps.active.quit()
        counter -= 1
        print(counter)


def getFileList():
    """Walks through all files in the path and makes a list of them"""
    folder = r"C:\Users\tyler.anderson\Desktop\2020 Payroll\Semi-Monthly"
    file_list = []
    counter = 0
    for dirpath, dirnames, filenames in os.walk(folder):
        for filename in [f for f in filenames if f.endswith(".xlsx")]:
            file = os.path.join(dirpath, filename)
            if 'Payroll Allocation' in file:
                file_list.append(file)
                print(file)
                counter += 1
    print(counter)
    """CHANGE A VALUE IN THE BUDGET FILES"""
    # updateExcelFiles(file_list)   # ADJUST SECOND PARAMETER FOR RENAME
    """PULL INFORMATION FROM BUDGET FILES"""
    getExcelDataBW(file_list, counter)
    """RENAME THE BUDGET FILES"""


if __name__ == '__main__':
    getFileList()
