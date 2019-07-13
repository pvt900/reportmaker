## TransferSheet.py Created by Robert Farmer
## Created to Transfer the Untimed Data from the Untimed Report
## To a Workbook containing a Macro that will upload new In-Product parts to the Ready Board.
##
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
        self.utbook = None
        self.macbook = load_workbook(
            filename='UploadUntimed.xlsm', keep_vba=True)
    def run_report_tool(self):
        '''
        This Function calls the Untimed Report Tool and Runs the Report tool
        '''
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

        utsheet = self.utbook.worksheets[2] #Sheet w/ Untimed Report 
        macsheet = self.macbook.worksheets[0] # Macro Workbook Sheets 
        new_mac = self.macbook.create_sheet("Untimed Datas", 1) #New Sheet in the Macro book
        new_mac = self.macbook.worksheets[1] #Selects the New Sheet.

        index_mac = {} # Initializes the Dictionary
        #Estimates Only List is to Exclude the Future Product Parts from The Ready Board.
        estimates_only = ["F9","F9A","F9B","F11A","7RMY20", "F9X"]
        #print(bool(index_mac))
        for mac_rows in macsheet.iter_rows(): # Creates Key-Value Dictionary for Iteration
            key = (mac_rows[3].value, mac_rows[2].value)
            index_mac[key] = mac_rows # The Row w/ a Key of a Part/Op is added to the Dictionary.
        for ut_rows in utsheet.iter_rows(): # Iterates Through the Rows of the Report
            key = (ut_rows[3].value, ut_rows[2].value) # Part/Op is the Key
            if index_mac.get(key) == None: # if the current Key from Row X doesn't exist
                if ut_rows[12].value not in estimates_only: # And if the Program isn't in the Estimates List
                    new_mac.append([cell.value for cell in ut_rows]) # Add it to the Sheet

        std = self.macbook.worksheets[0] # Sets the Old Macro Sheet as std
        self.macbook.remove(std) # Deletes the old Sheet so the New Populated Sheet is the only One

    def save_sheet(self):
        '''
        Saves the Work Book.
        '''
        self.macbook.save(filename='UploadUntimed.xlsm')

    def run_macro(self):
        '''
        Runs the Macro from the excel sheet. 
        Current Deprecated. Needs Revision to Work. Doesn't run Macro. 
        Do not Run This function it won't work. Run the Macro from within the Excel Workbook Manually
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
    currentReport = ("Untimed Report"+str(datetime.date.today()))
    x.utbook = load_workbook(filename= str(currentReport)+".xlsx")
    print("Book Loaded!")
    x.transfer_report()
    print("Data Transferred!")
    x.save_sheet()
    print("Saved! & Ready For Macro!")
    # x.run_macro()
    # print("Done!")

if __name__ == "__main__":
    main()
