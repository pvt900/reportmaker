
from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from collections import defaultdict
import collections 
import datetime
import time


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
        for row in utsheet.iter_rows():
            macsheet.append([cell.value for cell in row])
    def save_sheet(self):
        self.macbook.save(filename= 'UploadUntimed.xlsm')

def main():
    x = TransferSheet()
    x.transfer_report()
    x.save_sheet()
    print("Done!")
if __name__ == "__main__": main()