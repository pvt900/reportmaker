
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from collections import defaultdict
from UntimeReportTool import ManualReporter
import collections
import datetime
import time
import os
import os.path
import win32com.client


class TransferSheet:
    def __init__(self):
        '''
        Initializes all Variables for use within the ClassProgram
        '''
        Tk().withdraw()
        self.utbook = load_workbook(filename='Untimed Report2019-06-26.xlsx')
        self.macbook = load_workbook(
            filename='UploadUntimed.xlsm', keep_vba=True)
    def run_report_tool(self):
        Report = ManualReporter()
        Report.open_sapfile()
        Report.open_tracker()
        Report.trim_report()
        Report.check_rows()
        Report.CreateReport()
        Report.payables()
        Report.worktracker_scanner()
        Report.part_programs()
        Report.count_programs()
        Report.final_report()
        Report.save_workbook()

    def transfer_report(self):
        '''
        This creates the sheets for the two workbooks, clears the contents of the Upload Document
        Then once cleared it transfers the content of the UT. Report.
        '''

        utsheet = self.utbook.worksheets[2]
        macsheet = self.macbook.worksheets[0]
        new_mac = self.macbook.create_sheet("Untimed Datas", 1)
        new_mac = self.macbook.worksheets[1]

        index_mac = {}

        rows_to_add = []
        print(bool(index_mac))
        for mac_rows in macsheet.iter_rows():
            key = (mac_rows[3].value, mac_rows[2].value)
            index_mac[key] = mac_rows
        for ut_rows in utsheet.iter_rows():
            key = (ut_rows[3].value, ut_rows[2].value)
            if index_mac.get(key) == None:
                new_mac.append([cell.value for cell in ut_rows])

        std = self.macbook.worksheets[0]
        self.macbook.remove(std)

    def save_sheet(self):
        '''
        Saves the Work Book.
        '''
        self.macbook.save(filename='UploadUntimed.xlsm')
    '''
    NOTES FOR FUTURE DEVELOPMENT
    ````````````````````````````
        If InfoPath/Designer does not pan out and Flow/PowerApps continue to be unavailable. 
        Look into creating a comparison between the last uploaded set of parts and the current pull from the Report
        Disregarding items that are within. Anything new will stay and anything removed will stay. Before Uploading examine the
        current status of said parts and then move forward from there. probably two checks one to check if row is in macbook that feeds into new parts 
        and one to check if row from macbook is in new pull (check for Ts)

    '''

    def run_macro(self):
        '''
        Runs the Macro from the excel sheet. 
        Current Deprecated. Needs Revision to Work. Doesn't run Macro. 
        '''
        # Launch Excel and Open Wrkbook
        # xl=win32com.client.Dispatch("Excel.Application")
        # xl.Workbooks.Open(r"C:\\Users\\UFJUDFM\Desktop\\Python_Project\\UploadUntimed.xlsm") #opens workbook in readonly mode.

        # Run Macro
        # xl.Application.Run("UploadUntimed.xlsm!Module2.AddNew_SP")

        # Save Document and Quit.
        # xl.Application.Save()
        # xl.Application.Quit()

        # Cleanup the com reference.
        #del xl
        if os.path.exists("UploadUntimed.xlsm"):
            xl = win32com.client.Dispatch('Excel.Application')
            xl.Workbooks.Open(Filename="UploadUntimed.xlsm", ReadOnly=1)
            xl.Application.run("AddNew_SP")
            xl.Application.Quit()
            del xl


def main():
    x = TransferSheet()
    print("Script Initialized")
    x.run_report_tool()
    print("Untimed Ran")
    currentReport = "Untimed Report"+str(datetime.date.today())+".xlsx."
    x.utbook = load_workbook(filename= currentReport)
    x.transfer_report()
    print("Data Transferred!")
    x.save_sheet()
    print("Saved! & Ready For Macro!")
    # x.run_macro()
    # print("Done!")

if __name__ == "__main__":
    main()
