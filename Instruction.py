from tkinter import *

def openInstruction():
    InsWin = Toplevel()
    InsWin.title("Instructions")
    InsWin.state("zoomed")
    
    DESDescLabel = Label(InsWin, text="DES: Design Quantity (Please Specify)").grid(row=0, column=0, padx=(30,10), pady=(30,0), sticky=W)
    SPADescLabel = Label(InsWin, text="SPA: Spare Quantity, for certain comsumables (Please Specify)").grid(row=1, column=0, padx=(30,10), sticky=W)
    OHDescLabel = Label(InsWin, text="OH: On Hand Quantity, number of parts currently in company stock (Please Specify)").grid(row=2, column=0, padx=(30,10), sticky=W)
    REQDescLabel = Label(InsWin, text="REQ: Required Quantity, number of parts needed for this assembly, DES time Design Qty. plus SPA minus OH").grid(row=3, column=0, padx=(30,10), sticky=W)
    PCHDescLabel = Label(InsWin, text="PCH: Purchased Quantity, number of parts purchased, refer to Purchase Order").grid(row=4, column=0, padx=(30,10), sticky=W)
    BALDescLabel = Label(InsWin, text="BAL: Balance Quantity, number of parts yet to be purchased, completed assembly should have BAL = 0").grid(row=5, column=0, padx=(30,10), sticky=W)
    RCVDescLabel = Label(InsWin, text="RCV: Received Quantity, number of parts purchased and received, currently in stock or in used (Please Specify)").grid(row=6, column=0, padx=(30,10), sticky=W)
    OSDescLabel = Label(InsWin, text="OS: Outstanding Quantity, number of parts purchased but yet to be delivered, PCH minus RCV").grid(row=7, column=0, padx=(30,10), sticky=W)
    