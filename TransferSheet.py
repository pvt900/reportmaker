
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from collections import defaultdict
import collections 
import datetime
import time
import os, os.path
import win32com.client

class TransferSheet:
    def __init__(self):
        '''
        Initializes all Variables for use within the ClassProgram
        '''
        Tk().withdraw()
        self.utbook = load_workbook(filename = 'Untimed Report2019-06-17.xlsx')
        self.macbook = load_workbook(filename = 'UploadUntimed.xlsm', keep_vba= True)

    def transfer_report(self):
        utsheet = self.utbook.worksheets[2]
        macsheet = self.macbook.worksheets[0]

        for data_rows in macsheet.iter_rows():
            for data_cell in data_rows:
                data_cell.value = None

        for row in utsheet.iter_rows():
            macsheet.append([cell.value for cell in row])

    def save_sheet(self):
        self.macbook.save(filename= 'UploadUntimed.xlsm')

    def run_macro(self):

        #Launch Excel and Open Wrkbook
        xl=win32com.client.Dispatch("Excel.Application")  
       # xl.Workbooks.Open(Filename="C:\Users\UFJUDFM\Desktop\Python_Project\UploadUntimed.xlsm") #opens workbook in readonly mode. 

        #Run Macro
        xl.Application.Run("UploadUntimed.xlsm!Module2.AddNew_SP") 

        #Save Document and Quit.
        xl.Application.Save()
        xl.Application.Quit() 

        #Cleanup the com reference. 
        del xl

def main():
    x = TransferSheet()
    x.transfer_report()
    x.save_sheet()
    #x.run_macro()
    print("Done!")
if __name__ == "__main__": main()