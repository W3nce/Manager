from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox 
from tkcalendar import *
from datetime import datetime
from PIL import ImageTk, Image
from mysql import *
import mysql.connector
import csv
import re 
from fpdf import FPDF
import ConnConfig


def openRFQ(tabNote):
    
    RFQFrame = Frame(tabNote)
    tabNote.add(RFQFrame, text="Request for Quotation")
    RFQFrame.columnconfigure(0, weight=1)
    tabNote.select(RFQFrame)
   
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background="silver",
                    rowheight=20,
                    fieldbackground="light grey")
    
    style.map("Treeview")
# =============================================================================
#     
#     RFQWin = Tk()
#     RFQWin.iconbitmap("MWA_Icon.ico")
#     RFQWin.title('Request for Quotation')
#     RFQWin.state("zoomed")
#     RFQWin.rowconfigure(0, weight=1)
#     RFQWin.columnconfigure(0, weight=1)
#     RFQWin.resizable(height = None, width = None) 
#     
#     def CloseRFQ():
#         global output,XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED
#         if messagebox.askokcancel("Exit RFQ", "Do you want to clost this tab?",parent = root):
#             root.destroy()
#             output = email,driver,0
#             
#     root.protocol("WM_DELETE_WINDOW", Close)
#     
#     RFQWin.mainloop()
# =============================================================================
