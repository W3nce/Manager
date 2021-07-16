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
import ConnConfig,CountryRef
from CBTreeView import CbTreeview
#from CheckBoxTreeView import CheckboxTreeview


import os
import pathlib
import win32com.client as client
from AutoCombo import AutocompleteCombobox 

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

connRFQ = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password = logininfo[2])



def openRFQ(tabNote):
    global MachineList,MakerList,img
    
    #photo = Image.open('mag3.png')
    photo = PhotoImage(file = r"mag3.png")
    img = photo.subsample(4, 4)
    
    curRFQ = connRFQ.cursor()
    curRFQ.execute("""
                   SELECT `PROJECT_CLS_ID`,`PROJECT_NAME` FROM `index_pro_master`.`project_info`
                   """)
                   
    ProjNameDict = {projid[0]:projid[1] for projid in curRFQ.fetchall()}
    ProjList = list(ProjNameDict)
    
    curRFQ.execute("""
               SELECT EMPLOYEE_ID,EMPLOYEE_NAME FROM index_emp_master.emp_data where EMPLOYEE_NAME <> "ADMIN 1"
               """)
                       

    EmployeeDict = {Employee [0]: Employee[1] for Employee in curRFQ.fetchall()}
    
    curRFQ.close()
    
    
    MachineList = []
    MakerList = []
        
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
    
    def _QueryMachNo():
        global MachineList,MachineDict
        curRFQ = connRFQ.cursor()
        
        curRFQ.execute(f"""
                       SELECT MACH_ID,MACH_NAME,ORDER_QTY FROM {ProjectCombo.get()}.mach_index
                       """)
                       
        MachineDict = {Machid[0]:(Machid[1],Machid[2]) for Machid in curRFQ.fetchall()}
        MachineList = list(MachineDict)
        
        if MachineList:
            pass
        else:
            MachineCombo.set('No Machine for this Machine')
            return
        
        MachineCombo['value'] = MachineList
        MachineCombo.current(0) if MachineCombo['value'] else MachineCombo.set('Please create a Machine')
        curRFQ.close()
        
        ProjectNameLabel['text'] = ProjNameDict[ProjectCombo.get()]
        

    
    def _QueryMaker():
        global MakerList,MachineDict,AsmDict
        curRFQ = connRFQ.cursor()
        curRFQ.execute(f"""
                       SELECT `ASSEM_FULL`,`DES_QTY` FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}`
                       """)
                       
        AsmDict = {Asm[0]:Asm[1] for Asm in curRFQ.fetchall()}
        AsmList = list(AsmDict)

        if AsmList:
            pass
        else:
            MakerCombo.set('No Assembly for this Machine')
            return

        
        MakerList = []
        for Asm in AsmList:
            try: 
                MakerList.remove('No Maker for Selected Parts')
            except ValueError:
                pass
            curRFQ.execute(f"""
                           SELECT `Maker` FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}_{Asm}`
                           """)
            for Maker in curRFQ.fetchall():
                if Maker[0]:
                    if Maker[0] not in MakerList:
                        MakerList.append(Maker[0])
                elif 'No Maker for Selected Parts' not in MakerList:
                    MakerList.append('No Maker for Selected Parts')
                else:
                    pass
                           
            
        curRFQ.close()
        MakerCombo.set_completion_list(completion_list = MakerList)
        MakerCombo.current(0) if MakerCombo['value'] else MakerCombo.set('No parts in Unit Assembly')
        
        MacQtyBox.configure(state="normal")
        MacQtyBox.delete(0,END)
        MacQtyBox.insert(END,MachineDict[MachineCombo.get()][1])
        MacQtyBox.configure(state="readonly")
        
        MachineNameLabel['text'] = MachineDict[MachineCombo.get()][0]

    def QueryMachNo(event) :
        _QueryMachNo()   
        _QueryMaker()

    def QueryMaker(event):
        _QueryMaker()
        



    def CloseRFQ():
        if messagebox.askokcancel("Exit RFQ", "Do you want to clost this tab?",parent = RFQFrame):
            tabNote.forget(tabNoteg.select())
            RFQFrame.destroy()
            
    def SelectVendor():
        _SelectVendor(VendorBox,VendorAddressBox)     
        
    
    def ShowPurchaserName(event):
        _ShowPurchaserName()
        
    def _ShowPurchaserName():
                
        PurchaserNameBox.configure(state="normal")
        PurchaserNameBox.delete(0,END)
        PurchaserNameBox.insert(END,EmployeeDict[PurchaserCombo.get()])
        PurchaserNameBox.configure(state="readonly")
        
        
    def queryTreeUnit(Maker = None):
        
        if Maker == "No Maker for Selected Parts":
            _Maker = ''
        elif Maker:
            _Maker = Maker
        else:
            pass
        
        curRFQ = connRFQ.cursor()
        curRFQ.execute(f"""
                       SELECT `ASSEM_FULL`,`DES_QTY` FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}`
                       """)
                       
        AsmDict = {Asm[0]:Asm[1] for Asm in curRFQ.fetchall()}
        AsmList = list(AsmDict)

        
        
        for Asm in AsmList:
            if Maker:
                
                curRFQ.execute(f"""
                               SELECT oid, 
                               PartNum, Description, CLS,
                               Maker, Spec, V,  DES, SPA,
                               OH, REQ, PCH,  
                               UnitCost, Currency, TotalUnitCost
                               FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}_{Asm}`
                               where Maker = "{_Maker}"
                               """)
            else:
                
                curRFQ.execute(f"""
                               SELECT oid, 
                               PartNum, Description, CLS,
                               Maker, Spec, V,  DES, SPA,
                               OH, REQ, PCH,  
                               UnitCost, Currency, TotalUnitCost
                               FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}_{Asm}`
                               """)
              
            
                        
            PartsList = curRFQ.fetchall()
            for rec in PartsList:
                UnitTreeView.insert(parent="", index=END, iid=rec[0], 
                                values=(rec[1], rec[2], rec[3],
                                        rec[4], rec[5], rec[6], rec[7], rec[8],
                                        rec[9], AsmDict[Asm] , rec[10], rec[11], rec[12], rec[13], 
                                        rec[14],''))

            curRFQ.close()
            
    def clearTreeUnit():
        
        UnitTreeView.delete(*UnitTreeView.get_children())
            
    Frame1 = LabelFrame(RFQFrame)
    Frame1.columnconfigure(2, weight=1) 
    Frame1.grid(row = 0, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    ProjectLabel = Label(Frame1, text = 'Project')
    ProjectLabel.grid(row = 0, column = 0, padx = 10,pady =4)
    
    ProjectCombo = AutocompleteCombobox(Frame1, width=30, height = 6)         
    ProjectCombo.set_completion_list(completion_list = ProjList)
    ProjectCombo.current(0) if ProjectCombo['value'] else ProjectCombo.set('Please Create a Project')
    ProjectCombo.grid(row = 0, column = 1, padx = 10,pady =4,sticky = W)
    ProjectCombo.bind("<<ComboboxSelected>>", QueryMachNo)
    
    ProjectNameLabel = Label(Frame1, text = '')
    ProjectNameLabel.grid(row = 0, column = 2,pady =4,sticky  = W)
    
    MachineLabel = Label(Frame1, text = 'Machine')
    MachineLabel.grid(row = 1, column = 0, padx = 10,pady =4)
    
    MachineCombo = ttk.Combobox(Frame1, width=30, height = 6,value = MachineList)
    MachineCombo.grid(row = 1, column = 1, padx = 10,pady =4,sticky = W)
    MachineCombo.bind("<<ComboboxSelected>>", QueryMaker)
    
    MachineNameLabel = Label(Frame1, text = '')
    MachineNameLabel.grid(row = 1, column = 2,pady =4,sticky  = W)
    
    MakerLabel = Label(Frame1, text = 'Maker')
    MakerLabel.grid(row = 2, column = 0, padx = 10,pady =4)
    
    SelectMakerPartsFrame = Frame(Frame1)
    SelectMakerPartsFrame.grid(row = 2, column = 1,columnspan = 2, padx = 10,pady =4,sticky = EW)

    MakerCombo = AutocompleteCombobox(SelectMakerPartsFrame, width=30, height = 6)         
    MakerCombo.set_completion_list(completion_list = MakerList)
    MakerCombo.grid(row = 0, column = 0)
    
    SelectMakerPartsButton = Button (SelectMakerPartsFrame, width=14, text = 'Select Parts', command = lambda : queryTreeUnit(MakerCombo.get()))  
    SelectMakerPartsButton.grid(row = 0, column = 1,padx = 9)
    
    VendorLabel = Label(Frame1, text = 'Vendor')
    VendorLabel.grid(row = 3, column = 0, padx = 10,pady =4)
    
    SelectVendFrame = LabelFrame(Frame1)
    SelectVendFrame.grid(row = 3, column = 1, padx = (10,0),pady =4,sticky = EW)
    
    VendorBox = Entry(SelectVendFrame,width=33,state = "readonly")
    VendorBox.grid(row = 0,column = 0)
    

    SelectVendButton = Button(SelectVendFrame,command = SelectVendor, image = img)
    SelectVendButton.grid(row = 0,column = 1, padx = (8,0))
    
    VendorNameLabel = Label(Frame1, text = '')
    VendorNameLabel.grid(row = 3, column = 2,pady =4,sticky  = W)
    
    VendorAddressBox = Entry(Frame1)
    VendorAddressBox.grid(row = 4, column = 1, columnspan = 2, padx = 10,pady =4,sticky = EW)
    VendorAddressBox.insert(END,'Show Address')
    VendorAddressBox.configure(state="readonly")
    
    RFQRefLabel = Label(Frame1, text = 'RFQ Ref')
    RFQRefLabel.grid(row = 0, column = 4, padx = (10,4),pady =4)
    
    RFQRefBox = Entry(Frame1,width=28,state = "readonly")
    RFQRefBox.grid(row = 0, column = 5, padx = 4,pady =4,sticky = W)   
   
    StatusLabel = Label(Frame1, text = 'Status')
    StatusLabel.grid(row = 1, column = 4, padx = (10,4),pady =4)
    
    StatusBox = Entry(Frame1,width=28,state = "readonly")
    StatusBox.grid(row = 1, column = 5, padx = 4,pady =4,sticky = W)
    
    IssuedLabel = Label(Frame1, text = 'Issued')
    IssuedLabel.grid(row = 2, column = 4, padx = (10,4),pady =4)
    
    IssuedBox = Entry(Frame1,width=28,state = "readonly")
    IssuedBox.grid(row = 2, column = 5, padx = 4,pady =4,sticky = W)   
        
    MacQtyLabel = Label(Frame1, text = 'Mach Qty')
    MacQtyLabel.grid(row = 3, column = 4, padx = (10,4),pady =4)
    
    MacQtyBox = Entry(Frame1,width=28,state = "readonly")
    MacQtyBox.grid(row = 3, column = 5, padx = 4,pady =4,sticky = W)
         
    Frame2 = LabelFrame(RFQFrame)
    Frame2.columnconfigure(1, weight=1) 
    Frame2.grid(row = 1, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    Frame21 = Frame(Frame2)
    Frame21.grid(row = 0, column = 0,sticky = W)
    
    ReplyDueLabel = Label(Frame21,text = 'Reply Due')
    ReplyDueLabel.grid(row = 0, column = 0,padx = (10,4),pady =4,sticky = W)
    
    ReplyDueCalEnt = Entry(Frame21,state = 'readonly',width = 14)
    ReplyDueCalEnt.grid(row = 0, column = 1,padx = 4,pady =4,sticky = W)
    
    ReplyDueCalButton = Button(Frame21, text="Cal", font=("Arial", 8), command = lambda : CalWin(ReplyDueCalEnt))
    ReplyDueCalButton.grid(row = 0, column = 2,padx = 4,pady =4,sticky = W)
    
    DeliverByLabel = Label(Frame21,text = 'Deliver By')
    DeliverByLabel.grid(row = 0, column = 3,padx = (10,4),pady =4,sticky = W)
    
    DeliverBCalEnt = Entry(Frame21,state = 'readonly',width = 14)
    DeliverBCalEnt.grid(row = 0, column = 4,padx = 4,pady =4,sticky = W)
    
    DeliverBCalButton = Button(Frame21, text="Cal", font=("Arial", 8), command = lambda : CalWin(DeliverBCalEnt))
    DeliverBCalButton.grid(row = 0, column = 5,padx = 4,pady =4,sticky = W)
    
    CompletedCheckbutton  = Checkbutton(Frame21,text = 'Completed')
    CompletedCheckbutton.grid(row = 0, column = 6,padx = (10,4),pady =4,sticky = W)

    Frame22 = Frame(Frame2)
    Frame22.grid(row = 0, column = 1,sticky = E)

    CurrencyLabel = Label(Frame22,text = 'Currency')
    CurrencyLabel.grid(row = 0, column = 0,padx = (10,4),pady =4,sticky = W)
    
    CurrencyCombo = ttk.Combobox(Frame22,width=6, value=CountryRef.getCcyLst(), state="readonly")
    CurrencyCombo.grid(row = 0, column = 1,padx = 4,pady =4,sticky = W)
    CurrencyCombo.current(0)

    PurchaserLabel = Label(Frame22,text = 'Purchaser')
    PurchaserLabel.grid(row = 0, column = 2,padx = (10,4),pady =4,sticky = W)

    PurchaserCombo = ttk.Combobox(Frame22,width=8, value=list(EmployeeDict), state="readonly")
    PurchaserCombo.grid(row = 0, column = 3,padx = 4,pady =4,sticky = W)
    PurchaserCombo.bind("<<ComboboxSelected>>", ShowPurchaserName)
    PurchaserCombo.current(0)
    
    PurchaserNameBox = Entry(Frame22,width=28,state = "readonly")
    PurchaserNameBox.grid(row = 0, column = 4,padx = 4,pady =4,sticky = W)
    
    Frame3 = LabelFrame(RFQFrame)
    Frame3.columnconfigure(1, weight=1)
    Frame3.grid(row = 2, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    UnitTreeScroll = Scrollbar(Frame3)
    UnitTreeScroll.pack(side=RIGHT, fill=Y)
    
    UnitTreeView = CbTreeview(Frame3, yscrollcommand=UnitTreeScroll.set, 
                                selectmode="browse")
    UnitTreeScroll.config(command=UnitTreeView.yview)
    
    UnitTreeView.pack(padx=2, pady=2,fill="x", expand=True)
    
    
    
    
    UnitTreeView["columns"] = ("Part No", "Description", "CLS",  
                                "Maker", "Spec","V", "DES", "SPA", 
                                "OH","UA", "REQ", "PCH", 
                                "UnitCost", "Curr","Total","PO")
    
    UnitTreeView.column("#0", width=0, stretch=NO)
    # UnitTreeView.column("#0", width=50, anchor=W,stretch=NO)
    UnitTreeView.column("Part No", anchor=W, width=80)
    UnitTreeView.column("Description", anchor=W, width=140)
    UnitTreeView.column("CLS", anchor=CENTER, width=50)
    UnitTreeView.column("Maker", anchor=W, width=90)
    UnitTreeView.column("Spec", anchor=W, width=200)
    UnitTreeView.column("V", anchor=W, width=30)
    UnitTreeView.column("DES", anchor=E, width=35)
    UnitTreeView.column("SPA", anchor=E, width=35)
    UnitTreeView.column("OH", anchor=E, width=35)
    UnitTreeView.column("UA", anchor=E, width=35)
    UnitTreeView.column("REQ", anchor=E, width=35)
    UnitTreeView.column("PCH", anchor=E, width=35)
    UnitTreeView.column("UnitCost", anchor=E, width=80)
    UnitTreeView.column("Curr", anchor=W, width=45)
    UnitTreeView.column("Total", anchor=E, width=80)
    UnitTreeView.column("PO", anchor=CENTER, width=35,stretch=NO)
    
    
    UnitTreeView.heading("#0", text="Index", anchor=W)
    UnitTreeView.heading("Part No", text="Part", anchor=W)
    UnitTreeView.heading("Description", text="Description", anchor=W)
    UnitTreeView.heading("CLS", text="CLS", anchor=CENTER)
    UnitTreeView.heading("Maker", text="Maker", anchor=W)
    UnitTreeView.heading("Spec", text="Maker Spec", anchor=W) 
    UnitTreeView.heading("V", text="Ver", anchor=W)
    UnitTreeView.heading("DES", text="DES", anchor=E)
    UnitTreeView.heading("SPA", text="SPA", anchor=E)
    UnitTreeView.heading("OH", text="OH", anchor=E)
    UnitTreeView.heading("UA", text="UA", anchor=E)
    UnitTreeView.heading("REQ", text="REQ", anchor=E)
    UnitTreeView.heading("PCH", text="PCH", anchor=E)
    UnitTreeView.heading("UnitCost", text="Unit Cost", anchor=E)
    UnitTreeView.heading("Curr", text="Curr", anchor=W)
    UnitTreeView.heading("Total", text="Total Cost", anchor=E)
    UnitTreeView.heading("PO", text="PO", anchor=CENTER)
    
    
    BottomTreeFrame = Frame(Frame3)
    BottomTreeFrame.pack(side= BOTTOM, fill=BOTH,expand = True)
    BottomTreeFrame.columnconfigure(0, weight=1)
    
    BlankBox = Entry(BottomTreeFrame,state = "readonly" )
    BlankBox.grid(row = 0, column = 0,pady =4,sticky = EW)
    
    TotalSGDLabel = Label(BottomTreeFrame,text = 'Total SGD',font = ('Arial 10 bold'))
    TotalSGDLabel.grid(row = 0, column = 1,pady =4,padx = 10,sticky = E)
    
    TotalSGDBox = Entry(BottomTreeFrame,width=24,state = "readonly",font = ('Arial 10 bold'),justify=RIGHT)
    TotalSGDBox.grid(row = 0, column = 2, padx = 4,pady =4,sticky = W) 
    
    ClearPartsButton = Button(BottomTreeFrame,text = 'Clear Parts', width=14,command = clearTreeUnit)
    ClearPartsButton.grid(row = 0, column = 3, padx = 4,pady =4,sticky = W) 
    
    Frame4 = Frame(RFQFrame)
    Frame4.columnconfigure(2, weight=1)
    Frame4.grid(row = 3, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    CreateRFQButton = Button(Frame4,text = 'Create RFQ',width = 16)
    CreateRFQButton.grid(row = 0, column = 0)
    
    IssueRFQButton = Button(Frame4,text = 'Issue RFQ' ,width = 16)
    IssueRFQButton.grid(row = 0, column = 1,padx = 8, pady = 6)

    EditPartsButton = Button(Frame4,text = 'Save' ,width = 16)
    EditPartsButton.grid(row = 0, column = 3,padx = 8, pady = 6)

    SavePartsButton = Button(Frame4,text = 'Edit Parts' ,width = 16)
    SavePartsButton.grid(row = 0, column = 4,padx = 8, pady = 6)
    
    CloseButton = Button(Frame4,text = 'Close Tab' ,width = 16)
    CloseButton.grid(row = 0, column = 5,padx = (8,0), pady = 6)

    

    _QueryMachNo()

    _QueryMaker()

    _ShowPurchaserName()
    





def _SelectVendor(VendorBox,VendorAddressBox):
    global img
    connVend = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database = "INDEX_VEND_MASTER")
    
    curVend = connVend.cursor()
    curVend.execute(f"""SELECT * FROM VENDOR_LIST
                    """)
    VendList = curVend.fetchall()
    
    VendWin = Toplevel()    
    VendWin.title("Vendor Manager")
    VendWin.geometry("320x360")

    SearchFrame = Frame(VendWin)
    SearchFrame.grid(row = 0, column = 0, columnspan = 3, sticky = W, padx=5, pady=5,)
    
    TreeFrame = Frame(VendWin)
    TreeFrame.grid(row = 1, column = 0, columnspan = 3, sticky = W, padx=(15,5), pady=5,)
    
    VendScroll = Scrollbar(TreeFrame)
    VendScroll.pack(side=RIGHT, fill=Y)
    
    VendorSelectionTree = ttk.Treeview(TreeFrame,yscrollcommand=VendScroll.set, 
                                    selectmode="browse")

    VendScroll.config(command=VendorSelectionTree.yview)
    
    VendorSelectionTree.pack(padx=5, pady=5, ipadx=5, ipady=5)
    
    VendorSelectionTree['columns'] = ("Vendor",
                                      "Class",
                                      "Status")
    VendorSelectionTree.column("#0",width=0, stretch=NO, minwidth = 0)
    VendorSelectionTree.column("Vendor",width =140,minwidth = 140)
    VendorSelectionTree.column("Class",width =40,minwidth = 40)
    VendorSelectionTree.column("Status",width =60,minwidth =60)
    
    VendorSelectionTree.heading("#0",text = "Index", anchor = CENTER)
    VendorSelectionTree.heading("Vendor",text = "Vendor", anchor = W)   
    VendorSelectionTree.heading("Class",text = "Class", anchor = W)
    VendorSelectionTree.heading("Status",text = "Status", anchor = W)
                        
    VendorSelectionTree.delete(*VendorSelectionTree.get_children())
    
    statLst = ["Inactive", "Active"]
    for Vend in VendList:
        VendorSelectionTree.insert(parent="", index=END, iid=Vend[0], 
                                    text=Vend[0], values=(Vend[1], Vend[2], 
                                                          statLst[Vend[18]]))

    global curVEND
    curVEND = None
    def selectItem(event):
        global curVEND
        curVEND =VendorSelectionTree.item(VendorSelectionTree.focus())['values']

        
    VendorSelectionTree.bind('<ButtonRelease-1>', selectItem) 
    
    
    SearchLabel = Label(SearchFrame, text = 'Search by Vendor :')
    SearchLabel.grid(row=0, column = 0,padx=5, pady=5,)
    
    SearchEntry = Entry(SearchFrame)
    SearchEntry.grid(row=0, column = 1,padx=5, pady=5,)
    
    def Search():
        newlst = []
        for i in VendList:
            if SearchEntry.get().lower() in i[0].lower():
                newlst.append(i)
        VendorSelectionTree.delete(*VendorSelectionTree.get_children())
        icount = 1
        for Vend in newlst:
            VendorSelectionTree.insert('','end', text = str(icount),values = tuple([Vend[i] for i in range(len(Vend))]))
            icount +=1
            
    

    SearchButton = Button(SearchFrame,command = Search, image = img)
    SearchButton.grid(row=0, column = 2)

    def SelectedVend():
        try:
            VendorBox.config(state="normal")
            VendorBox.delete(0, END)  
            VendorBox.insert(END,curVEND[0])
            VendorBox.config(state="readonly")
            
            VendorAddressBox.config(state="normal")
            VendorAddressBox.delete(0, END)  
            VendAddress = curVEND[8]+','+curVEND[9]+','+curVEND[6]+','+curVEND[5]+','+curVEND[7]+','+curVEND[4]
            VendorAddressBox.insert(END,VendAddress)
            VendorAddressBox.config(state="readonly")
            
            VendorNameLabel['text'] = curVEND[3]
            
            
            
            VendWin.destroy()
        except:
            messagebox.showwarning("No Vendor Selected", "Please Select a Vendor", 
                                   parent=VendWin)
    
    def ClearVend():
        VendorBox.config(state="normal")
        VendorBox.delete(0, END)
        VendorBox.config(state="readonly")
        
        VendorAddressBox.config(state="normal")
        VendorAddressBox.delete(0, END)
        VendorAddressBox.config(state="readonly")
        VendWin.destroy()
        
        VendorNameLabel['text'] = ''
    
    def BackVend():
        VendWin.destroy()

    SelButton = Button(VendWin, text ="Select", command = SelectedVend, width = 10)
    SelButton.grid(row = 2, column = 0)
    
    ClearButton = Button(VendWin, text ="Clear", command = ClearVend, width = 10)
    ClearButton.grid(row = 2, column = 1)
    
    ExitButton = Button(VendWin, text ="Exit", command = BackVend, width = 10)
    ExitButton.grid(row = 2, column = 2)


def CalWin(Entry):
    calWin = Toplevel()
    calWin.title("Select the Date")
    calWin.columnconfigure( 1,weight = 1)
    
    cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
    cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
    
    def confirmDate():
        val = cal.get_date()
        Entry.config(state="normal")
        Entry.delete(0, END)
        Entry.insert(0, val)
        Entry.config(state="readonly")
        calWin.destroy()
    
    def emptyDate():
        Entry.config(state="normal")
        Entry.delete(0, END)
        Entry.config(state="readonly")
        calWin.destroy()

    buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
    buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
    buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
    buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
    buttonClose = Button(calWin, text="Close", command=calWin.destroy)
    buttonClose.grid(row=1, column=2, padx=5, pady=5)


if __name__ == "__main__":
    
    root = Tk()
    root.title("Request for Quotation") 
    root.state("zoomed")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    global tabNote
    tabNote = ttk.Notebook(root)
    tabNote.grid(row=0, column=0, sticky="NSEW") 
    openRFQ(tabNote)
    root.mainloop()
    
    
    
    
    