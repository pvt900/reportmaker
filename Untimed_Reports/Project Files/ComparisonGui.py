from tkinter import *
from CRG import ManualReporter

class App:
    def __init__(self, master):
        self.workhandler = ManualReporter()
        frame = Frame(master)
        frame.pack()
        self.openO = Button(frame, text="Open Old Filtered Report", command=self.load_old)
        self.openO.grid(column=0, row=0)
        self.openN = Button(frame, text="Open Lastest SAP Report", command=self.load_new)
        self.openN.grid(column=2, row=0)

        self.calculate1 = Button(frame,text="Run Report & Save", command=self.calc)
        self.calculate1.grid(column=2, row=10)
        
        self.up= Label(frame, text="Total Untimed Parts: NaN")
        self.xt = Label(frame, text="Total Total X's to T's: NaN")
        self.p = Label(frame,  text="Payables: NaN")
        self.np = Label(frame, text="Non Payables: NaN")
        self.up.grid(column=1, row=3)
        self.xt.grid(column=1, row=4)
        self.p.grid(column=1, row=5)
        self.np.grid(column=1, row=6)
    def load_old(self):
        self.openO.configure(text="Old File Loaded")
        self.workhandler.open_sapfile()
    def load_new(self):
        self.openN.configure(text="New File Loaded")
        self.workhandler.open_tracker()
    def calc(self):
        self.workhandler.trim_report()
        self.workhandler.check_rows()
        self.workhandler.CreateReport()
        self.workhandler.payables()
        XT,Pay,NPay,Total = self.workhandler.save_workbook()
        
        txt1 = "Total Untimed Parts: " + str(Total)
        txt2 = "Total X's to T's: " + str(XT)
        txt3 = "Payables : " + str(Pay)
        txt4 = "Non Payables: " + str(NPay)
        self.up.configure(text=txt1)
        self.xt.configure(text=txt2)
        self.p.configure(text=txt3)
        self.np.configure(text=txt4)
      

root = Tk()

root.title("Untimed Report Data Handler")
app = App(root)
root.mainloop()
