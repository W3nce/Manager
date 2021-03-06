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

from 

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

def openPurchase():
    RepWin = Toplevel()
    RepWin.title("Generate Purchase Order")
    RepWin.state("zoomed")
    RepWin.columnconfigure(0, weight=1)
    RepWin.rowconfigure(0, weight=1)

    
    tabNoteRep = ttk.Notebook(RepWin)
    tabNoteRep.grid(row=0, column=0, sticky="NSEW")
    
    frameRep = Frame(tabNoteRep)
    tabNoteRep.add(frameRep, text="Purchase Order Selection")
    
    frameRep.columnconfigure(0, weight=1)

    connInit = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2])

    SetupCommand = ["""CREATE SCHEMA IF NOT EXISTS `PUR_ORDER_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci""",
                    
                    """
                    CREATE TABLE IF NOT EXISTS `PUR_ORDER_MASTER`.`PUR_ORDER_LIST`
                    (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                     `PurOrderNum` VARCHAR(20),
                     `PaymentTerm` VARCHAR(255),
                     `OrderDate` DATE,
                        `VendorRemark` VARCHAR(100),
                     `DateEntry` DATETIME NOT NULL)
                    
                    ENGINE = InnoDB
                    DEFAULT CHARACTER SET = utf8mb4
                    COLLATE = utf8mb4_0900_ai_ci""",
                    
                    """CREATE SCHEMA IF NOT EXISTS `COMPANY_INFO` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci""",
                    
                    """
                    CREATE TABLE IF NOT EXISTS `COMPANY_INFO`.`COMPANY_MWA`
                    (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                     `ComName` VARCHAR(100),
                     `ComRegNum` VARCHAR(30),
                     `GSTRegNum` VARCHAR(30),
                     `Address` VARCHAR(100),
                     `PosCode` VARCHAR(50),
                     `CenterA` VARCHAR(50),
                     `CenterB` VARCHAR(50),
                     `ContactNum` VARCHAR(20),
                     `Email` VARCHAR(50))
                    
                    ENGINE = InnoDB
                    DEFAULT CHARACTER SET = utf8mb4
                    COLLATE = utf8mb4_0900_ai_ci"""
                    ]

    curInit = connInit.cursor()
    for com in SetupCommand:
        curInit.execute(com)
        connInit.commit()
    
    curInit.close()



    connMain = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database= "INDEX_PRO_MASTER")
    
    connCom = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database= "COMPANY_INFO")
    
    connPur = mysql.connector.connect(host = logininfo[0],
                                      user = logininfo[1], 
                                      password =logininfo[2],
                                      database= "PUR_ORDER_MASTER")
    
    connVend = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database= "INDEX_VEND_MASTER")

    RepTabTitleLabel = Label(frameRep, text="Purchase Order List", font=("Arial", 12))
    RepTabTitleLabel.grid(row=0, column=0, padx=0, pady=0, ipadx=0, ipady=5, sticky=W+E)
    
    OrderTreeFrame = Frame(frameRep)
    OrderTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky="EW")
    # OrderTreeFrame.pack(fill="x", expand=True)
    
    OrderTreeScroll = Scrollbar(OrderTreeFrame)
    OrderTreeScroll.pack(side=RIGHT, fill=Y)
    
    OrderTreeView = ttk.Treeview(OrderTreeFrame, yscrollcommand=OrderTreeScroll.set, selectmode="browse")
    OrderTreeScroll.config(command=OrderTreeView.yview)
    # OrderTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
    OrderTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
    
    OrderTreeView['columns'] = ("PurOrderNum", 
                                "PaymentTerm", 
                                "OrderDate", 
                                "VendorRemark",
                                "DateEntry")
    
    OrderTreeView.column("#0", anchor = CENTER, width =50, minwidth = 0)
    # OrderTreeView.column("#0",  width=0, stretch=NO)
    OrderTreeView.column("PurOrderNum", anchor = CENTER, width = 200, minwidth = 50)
    OrderTreeView.column("PaymentTerm", anchor = CENTER, width= 200, minwidth = 50)
    OrderTreeView.column("OrderDate", anchor = CENTER, width = 150, minwidth = 50)
    OrderTreeView.column("VendorRemark", anchor = W, width = 450, minwidth = 50)
    OrderTreeView.column("DateEntry", anchor = CENTER, width = 150, minwidth = 50)
    
    OrderTreeView.heading("#0", text = "Index")
    OrderTreeView.heading("PurOrderNum", text = "Purchase Order Number")
    OrderTreeView.heading("PaymentTerm", text = "Payment Term")
    OrderTreeView.heading("OrderDate", text = "Order Date")
    OrderTreeView.heading("VendorRemark", text = "Vendor Remark")
    OrderTreeView.heading("DateEntry", text = "Date of Entry")

    def genOrderNum():
        curPur = connPur.cursor()
        curPur.execute("""SELECT MAX(oid) FROM PUR_ORDER_LIST """)
        maxOID = curPur.fetchall()[0][0]
        
        if maxOID == None:
            nextInt = 1
        
        else:
            curPur.execute(f"""SELECT * FROM PUR_ORDER_LIST WHERE oid = {maxOID} """)
            result = curPur.fetchall()
            
            currentNum = result[0][1]
            
            latestYear = result[0][5].year
            latestMonth = result[0][5].month
            latestDay = result[0][5].day
            
            currentYear = datetime.now().year
            currentMonth = datetime.now().month
            currentDay = datetime.now().day
        
            if latestYear == currentYear and latestMonth == currentMonth and latestDay == currentDay:
                currentInt = int(str(currentNum)[-2:])
                nextInt = currentInt + 1
            else:
                nextInt = 1
        
        connPur.commit()
        curPur.close()
        
        yearDigit = str((datetime.now().year)%100)
        monthDigit = str(datetime.now().month).rjust(2,"0")
        dayDigit = str(datetime.now().day).rjust(2,"0")
        numTwoDigit = str(nextInt).rjust(2,"0")
        newPurOrderID = f"MWAPO{yearDigit}{monthDigit}{dayDigit}{numTwoDigit}"
        
        PurOrderNumBox.delete(0, END)
        PurOrderNumBox.insert(0, newPurOrderID)
    
    def payLstGen():
        payWin = Toplevel()
        payWin.title("Select Payment Method")
        payWin.geometry("210x300")
        
        sclFrame = Frame(payWin)
        sclFrame.grid(row=0, column=0, columnspan=3, padx=20, pady=10, ipadx=5, ipady=0, sticky=W+E)
                
        scl = Scrollbar(sclFrame, orient=VERTICAL)
        scl.pack(side=RIGHT, fill=Y)
        
        payLstBox = Listbox(sclFrame, width=20, selectmode=SINGLE, yscrollcommand=scl.set)
        scl.config(command=payLstBox.yview)
        payLstBox.pack(ipady=20)
        
        payMethodLst = ["COD", "EOM", "CND", "CBS", "CIA", "CWO",
                        "Net 7", "Net 10", "Net 30", "Net 60", "Net 90", 
                        "Others (Please Specify)"]
        
        for method in payMethodLst:
            payLstBox.insert(END, method)
        
        PaymentTermBox.delete(0, END)
        PaymentTermBox.config(state="readonly")
        
        def confirmPay():
            paySelect = payLstBox.get(ANCHOR)
            if paySelect == "":
                messagebox.showwarning("No Method Selected", "Please Select a Method", 
                                       parent=payWin)
            elif paySelect == "Others (Please Specify)":
                PaymentTermBox.config(state="normal")
                PaymentTermBox.delete(0, END)
                payWin.destroy()
            else:
                PaymentTermBox.config(state="normal")
                PaymentTermBox.delete(0, END)
                PaymentTermBox.insert(0, paySelect)
                PaymentTermBox.config(state="readonly")
                payWin.destroy()
        
        def emptyPay():
            PaymentTermBox.config(state="normal")
            PaymentTermBox.delete(0, END)
            PaymentTermBox.config(state="readonly")
            payWin.destroy()
        
        buttonConfirmPay = Button(payWin, text="Confirm", command=confirmPay)
        buttonConfirmPay.grid(row=1, column=0, padx=5, pady=5)
        
        buttonEmptyPay = Button(payWin, text="Clear", command=emptyPay)
        buttonEmptyPay.grid(row=1, column=1, padx=5, pady=5)
        
        buttonClosePay = Button(payWin, text="Close", command=payWin.destroy)
        buttonClosePay.grid(row=1, column=2, padx=5, pady=5)
    
    def vendorLstGen():
        curVend = connVend.cursor()
        curVend.execute(f"""SELECT * FROM VENDOR_LIST """)
        VendList = curVend.fetchall()
        
        VendWin = Toplevel()    
        VendWin.title("Vendor Selection")
        VendWin.geometry("320x360")

        SearchFrame = Frame(VendWin)
        SearchFrame.grid(row = 0, column = 0, columnspan = 3,sticky = W,padx=5, pady=5,)
        
        TreeFrame = Frame(VendWin)
        TreeFrame.grid(row = 1, column = 0, columnspan = 3,sticky = W,padx=5, pady=5,)
        
        VendScroll = Scrollbar(TreeFrame)
        VendScroll.pack(side=RIGHT, fill=Y)
        
        VendorSelectionTree = ttk.Treeview(TreeFrame,yscrollcommand=VendScroll.set, 
                                        selectmode="browse")
    
        VendScroll.config(command=VendorSelectionTree.yview)
        
        VendorSelectionTree.pack(padx=(15,5), pady=5, ipadx=5, ipady=5)
        
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
        SearchLabel.grid(row=0,column = 0,padx=5, pady=5,)
        
        SearchEntry = Entry(SearchFrame)
        SearchEntry.grid(row=0,column = 1,padx=5, pady=5,)
        
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
                
        global img
        #photo = Image.open('mag3.png')
        photo = PhotoImage(file = r"mag3.png")
        img = photo.subsample(4, 4)
        SearchButton = Button(SearchFrame,command = Search, image = img)
        SearchButton.grid(row=0, column = 2)
    
        def SelectedVend():
            try:
                VendorRemarkBox.config(state="normal")
                VendorRemarkBox.delete(0, END)  
                VendorRemarkBox.insert(END,curVEND[0])
                VendorRemarkBox.config(state="readonly")
                VendWin.destroy()
            except:
                VendorRemarkBox.config(state="normal")
                VendorRemarkBox.delete(0, END)  
                VendorRemarkBox.config(state="readonly")
                messagebox.showwarning("No Vendor Selected", "Please Select a Vendor", 
                                       parent=VendWin)
        
        def ClearVend():
            VendorRemarkBox.config(state="normal")
            VendorRemarkBox.delete(0, END)
            VendorRemarkBox.config(state="readonly")
            VendWin.destroy()
        
        def BackVend():
            VendWin.destroy()
    
        SelButton = Button(VendWin, text ="Select",command = SelectedVend, width = 10)
        SelButton.grid(row = 2, column = 0)
        
        ClearButton = Button(VendWin, text ="Clear", command = ClearVend, width = 10)
        ClearButton.grid(row = 2, column = 1, sticky = W)
        
        ExitButton = Button(VendWin, text ="Exit", command = BackVend, width = 10)
        ExitButton.grid(row = 2, column = 2, sticky = W)
    
    def orderCalPro():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("240x240")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            OrderDateBox.config(state="normal")
            OrderDateBox.delete(0, END)
            OrderDateBox.insert(0, val)
            OrderDateBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            OrderDateBox.config(state="normal")
            OrderDateBox.delete(0, END)
            OrderDateBox.config(state="readonly")
            calWin.destroy()
        
        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)

    def queryTreeOrder():
        curPur = connPur.cursor()
        curPur.execute("SELECT * FROM PUR_ORDER_LIST")
        recLst = curPur.fetchall()    
        connPur.commit()
        curPur.close()
        
        curVend = connVend.cursor()
        
        if recLst == []:
            buttonLoadPur.config(state=DISABLED)
        
        else:
            buttonLoadPur.config(state=NORMAL)   
            for rec in recLst:
                curVend.execute(f"SELECT * FROM VENDOR_LIST WHERE VENDOR_NAME = '{rec[4]}'")
                vendorInfo = curVend.fetchall()
                if vendorInfo == []:
                    addressFull = ""
                else:
                    addressFull = f"{vendorInfo[0][3]} ({vendorInfo[0][2]}) {vendorInfo[0][6]} {vendorInfo[0][5]} {vendorInfo[0][7]} {vendorInfo[0][4]}"
                OrderTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                     values=(rec[1], rec[2], rec[3], addressFull, rec[5]))
        curVend.close()

    def updateOrder():
        sqlCommand = f"""UPDATE PUR_ORDER_LIST SET
        `PaymentTerm` = %s,
        `OrderDate` = %s,
        `VendorRemark` = %s
        
        WHERE `oid` = %s"""
        
        selected = OrderTreeView.focus()
        orderNumRef = PurOrderNumBox.get()
        
        def checkDateOrder(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        inputs = (PaymentTermBox.get(), checkDateOrder(OrderDateBox),
                  VendorRemarkBox.get(), selected)
        
        curPur = connPur.cursor()
        curPur.execute(sqlCommand, inputs)
        connPur.commit()
        curPur.close()
        
        clearEntryOrder()
        OrderTreeView.delete(*OrderTreeView.get_children())
        queryTreeOrder()
        
        messagebox.showinfo("Update Successful", 
                            f"You Have Updated Order {orderNumRef}", parent=RepWin) 
    
    def createOrder():
        timeNow = datetime.now()
        formatDate = timeNow.strftime('%Y-%m-%d %H:%M:%S')
        
        createComOrder = f"""INSERT INTO `PUR_ORDER_LIST` (
        PurOrderNum, PaymentTerm, OrderDate, VendorRemark, DateEntry)
        
        VALUES (%s, %s, %s, %s, %s)"""
        
        def checkDateOrder(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        valueOrder = (PurOrderNumBox.get(), PaymentTermBox.get(), 
                      checkDateOrder(OrderDateBox),
                      VendorRemarkBox.get(), formatDate)
        
        curPur = connPur.cursor()
        curPur.execute(createComOrder, valueOrder)
        connPur.commit()
        
        PurOrderRef = PurOrderNumBox.get()       
        
        curPur.execute(f""" CREATE TABLE IF NOT EXISTS `{PurOrderRef}` (
            `oid` INT AUTO_INCREMENT PRIMARY KEY,
            `PartNum` VARCHAR(50),
            `Description` VARCHAR(100),
            `Maker` VARCHAR(100),
            `Spec` VARCHAR(150),
            `REQ` VARCHAR(10),
            `Tax` VARCHAR(20),
            `Vendor` VARCHAR(100),
            `UnitCost` VARCHAR(20),
            `Currency` VARCHAR(20))
        
            ENGINE = InnoDB
            DEFAULT CHARACTER SET = utf8mb4
            COLLATE = utf8mb4_0900_ai_ci""")
            
        connPur.commit()
        curPur.close()
        clearEntryOrder()
        OrderTreeView.delete(*OrderTreeView.get_children())
        queryTreeOrder()
        
        messagebox.showinfo("Create Successful", 
                            f"You Have Created Purchase Order {PurOrderRef}", parent=RepWin) 
    
    def deleteOrder():
        selected = OrderTreeView.focus()
        sqlDelete = "DELETE FROM PUR_ORDER_LIST WHERE oid = %s"
        valDelete = (selected, )
        
        curPur = connPur.cursor()
        curPur.execute(sqlDelete, valDelete)
        connPur.commit()
        
        orderNumRef = PurOrderNumBox.get()
        
        curPur.execute(f"DROP TABLE IF EXISTS `{orderNumRef}`")
        connPur.commit()
        
        clearEntryOrder()
        OrderTreeView.delete(*OrderTreeView.get_children())
        queryTreeOrder()
        
        messagebox.showinfo("Delete Successful", 
                            f"You Have Deleted Purchase Order {orderNumRef}", parent=RepWin) 
        curPur.close()
    
    def selectOrder():
        try:
            selected = OrderTreeView.focus()        
            sqlSelect = f"SELECT * FROM PUR_ORDER_LIST WHERE oid = %s"
            valSelect = (selected, )
            
            curPur = connPur.cursor()
            curPur.execute(sqlSelect, valSelect)
            recLst = curPur.fetchall()
            connPur.commit()
            curPur.close()
    
            clearEntryOrder()
            
            PurOrderNumBox.insert(0, recLst[0][1])
            
            PaymentTermBox.config(state="normal")
            PaymentTermBox.insert(0, recLst[0][2])
            PaymentTermBox.config(state="readonly")
            
            if recLst[0][3] == None:
                OrderDateBox.config(state="normal")
                OrderDateBox.insert(0, "")
                OrderDateBox.config(state="readonly")
            else:
                OrderDateBox.config(state="normal")
                OrderDateBox.insert(0, recLst[0][3])
                OrderDateBox.config(state="readonly")
            
            VendorRemarkBox.config(state="normal")
            VendorRemarkBox.insert(0, recLst[0][4])
            VendorRemarkBox.config(state="readonly")
            
            buttonUpdatePur.config(state=NORMAL)
            buttonCreatePur.config(state=DISABLED)
            buttonDeletePur.config(state=NORMAL)
            
        except:
            clearEntryOrder()
            messagebox.showerror("Error", "Please Select a Value", parent=RepWin)
    
    def deselectOrder():
        selected = OrderTreeView.selection()
        if len(selected) > 0:
            OrderTreeView.selection_remove(selected[0])
            clearEntryOrder()
        else:
            clearEntryOrder()
    
    def clearEntryOrder():
        buttonUpdatePur.config(state=DISABLED)
        buttonDeletePur.config(state=DISABLED)
        buttonCreatePur.config(state=NORMAL)
        
        PurOrderNumBox.delete(0, END)
        
        PaymentTermBox.config(state="normal")
        PaymentTermBox.delete(0, END)
        PaymentTermBox.config(state="readonly")
        
        OrderDateBox.config(state="normal")
        OrderDateBox.delete(0, END)
        OrderDateBox.config(state="readonly")
        
        VendorRemarkBox.config(state="normal")
        VendorRemarkBox.delete(0, END)
        VendorRemarkBox.config(state="readonly")
    
    def refreshOrder():
        clearEntryOrder()
        OrderTreeView.delete(*OrderTreeView.get_children())
        queryTreeOrder()
    
    def loadOrder():
        buttonLoadPur.config(state=DISABLED)
        
        curPur = connPur.cursor()
        selected = OrderTreeView.focus()
        sqlSelect = "SELECT * FROM PUR_ORDER_LIST WHERE oid = %s"
        valSelect = (selected, )

        curPur.execute(sqlSelect, valSelect)
        purSelect = curPur.fetchall()
        
        connPur.commit()
        curPur.close() 

        PurOrderNumRef = purSelect[0][1]
        VendorRef = purSelect[0][4]
        
        framePur = Frame(tabNoteRep)
        tabNoteRep.add(framePur, text="Select Unit for Purchase Order")
        framePur.columnconfigure(0, weight=1)
        tabNoteRep.select(1)





        def FetchProIndex():
            curMain = connMain.cursor()
            curMain.execute("SELECT * FROM PROJECT_INFO")
            proIndex = curMain.fetchall()
            curMain.close()
            
            if proIndex == []:
                return ["No Project Found"]
            else:
                proIndexLst = []
                for item in proIndex:
                    proIndexLst.append(item[1])    
                return proIndexLst
        
        def ProSelect(e):
            proName = selectProBox.get()
            if proName == "No Project Found":
                selectMachBox.config(value=["No Mach Found"])
                selectAssemBox.config(value=["No Assem Found"])
                selectMachBox.current(0)
                selectAssemBox.current(0)
                
            else:
                connLoad = mysql.connector.connect(host = logininfo[0],
                                                   user = logininfo[1], 
                                                   password =logininfo[2],
                                                   database= f"{proName}")
                
                curLoad = connLoad.cursor()
                curLoad.execute("SELECT * FROM MACH_INDEX")
                machIndex = curLoad.fetchall()
                curLoad.close()
                
                if machIndex == []:
                    selectMachBox.config(value=["No Mach Found"])
                    selectAssemBox.config(value=["No Assem Found"])
                else:
                    machIndexLst = []
                    for item in machIndex:
                        machIndexLst.append(item[1])
                    selectMachBox.config(value=machIndexLst)
                selectMachBox.current(0)
                selectAssemBox.current(0)
    
        def MachSelect(e):
            proName = selectProBox.get()
            machName = selectMachBox.get()
            if machName == "No Mach Found":
                selectAssemBox.config(value=["No Assem Found"])
                selectAssemBox.current(0)
            
            else:
                connLoad = mysql.connector.connect(host = logininfo[0],
                                                   user = logininfo[1], 
                                                   password =logininfo[2],
                                                   database= f"{proName}")
                curLoad = connLoad.cursor()
                curLoad.execute(f"SELECT * FROM `{machName}`")
                assemIndex = curLoad.fetchall()
                curLoad.close()
                
                if assemIndex == []:
                    selectAssemBox.config(value=["No Assem Found"])
                else:
                    assemIndexLst = []
                    for item in assemIndex:
                        assemIndexLst.append(f"{item[1]}{item[2]}")
                    selectAssemBox.config(value=assemIndexLst)
                selectAssemBox.current(0)
                
        def AssemSelect(e):
            proName = selectProBox.get()
            machName = selectMachBox.get()
            assemName = selectAssemBox.get()
            if assemName == "No Assem Found":
                pass
            else:
                global AssemblyFullName
                AssemblyFullName = f"{proName}-{machName}-{assemName}"
                assemFullLabel.config(text=AssemblyFullName)
                fullName = f"{machName}_{assemName}"
                connLoad = mysql.connector.connect(host = logininfo[0],
                                                   user = logininfo[1], 
                                                   password =logininfo[2],
                                                   database= f"{proName}")
                
                curLoad = connLoad.cursor()
                curLoad.execute(f"SELECT * FROM `{fullName}`")
                unitIndex = curLoad.fetchall()
                curLoad.close()
                
                ValTreeView.delete(*ValTreeView.get_children())
                for rec in unitIndex:
                    ValTreeView.insert(parent="", index=END, iid=rec[0],
                                        values=(rec[1], rec[2], rec[3], rec[4], rec[5], rec[6], rec[7], 
                                                rec[8], rec[9], rec[10], rec[11], rec[12], rec[13], 
                                                rec[14], rec[15], rec[16], rec[17], rec[18]))
                    
        framePMA = Frame(framePur)
        framePMA.grid(row=0, column=0, columnspan=2, padx=10, pady=0, ipadx=10, ipady=0 , sticky=W+E)
        
        selectProLabel = Label(framePMA, text="Select Project")
        selectMachLabel = Label(framePMA, text="Select Machine")
        selectAssemLabel = Label(framePMA, text="Select Assembly")
        selectProBox = ttk.Combobox(framePMA, width=15, value=FetchProIndex(), state="readonly")
        selectMachBox = ttk.Combobox(framePMA, width=15, value=["No Mach Found"], state="readonly")
        selectAssemBox = ttk.Combobox(framePMA, width=15, value=["No Assem Found"], state="readonly")
        BOMSelectLabel = Label(framePMA, text="BOM Selected:")
        assemFullLabel = Label(framePMA, text="")
        
        selectProBox.current(0)
        if selectProBox.get() != "No Project Found":
            selectProBox.config(state="normal")
            selectProBox.delete(0, END)
            selectProBox.insert(0, "Please Select")
            selectProBox.config(state="readonly")
        selectMachBox.current(0)
        selectAssemBox.current(0)
        
        selectProBox.bind("<<ComboboxSelected>>", ProSelect)
        selectMachBox.bind("<<ComboboxSelected>>", MachSelect)
        selectAssemBox.bind("<<ComboboxSelected>>", AssemSelect)
        
        selectProLabel.grid(row=0, column=0, padx=10, pady=(10,0), sticky=E)
        selectMachLabel.grid(row=0, column=2, padx=10, pady=(10,0), sticky=E)
        selectAssemLabel.grid(row=0, column=4, padx=10, pady=(10,0), sticky=E)
        selectProBox.grid(row=0, column=1, padx=10, pady=(10,0), sticky=W)
        selectMachBox.grid(row=0, column=3, padx=10, pady=(10,0), sticky=W)
        selectAssemBox.grid(row=0, column=5, padx=(10,50), pady=(10,0), sticky=W)
        BOMSelectLabel.grid(row=0, column=6, padx=10, pady=(10,0), sticky=E)
        assemFullLabel.grid(row=0, column=7, padx=10, pady=(10,0), sticky=W)
        
        # style = ttk.Style()
        # style.theme_use("clam")
        
        # style.configure("Treeview",
        #                 background="silver",
        #                 rowheight=20,
        #                 fieldbackground="light grey")
        
        # style.map("Treeview") 
        
        ValTreeFrame = Frame(framePur)
        ValTreeFrame.grid(row=1, column=0, columnspan=2, padx=10, pady=(5,0), ipadx=10, ipady=0, sticky=W+E)
        # ValTreeFrame.pack()
        
        ValTreeScroll = Scrollbar(ValTreeFrame)
        ValTreeScroll.pack(side=RIGHT, fill=Y)
        
        ValTreeView = ttk.Treeview(ValTreeFrame, yscrollcommand=ValTreeScroll.set, 
                                    height=6, selectmode="extended")
        ValTreeScroll.config(command=ValTreeView.yview)
        # ValTreeView.grid(row=2, column=0, columnspan=10, padx=10)
        ValTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
        
        ValTreeView["columns"] = ("Part", "Description", "D", "S", "CLS", "V", 
                                    "Maker", "Spec", "DES", "SPA", "OH", "REQ", 
                                    "PCH", "BAL", "RCV", "ADJ", "Vendor", "UnitCost")
        
        # ValTreeView.column("#0", anchor=CENTER, width=50)
        ValTreeView.column("#0", width=0 ,stretch=NO)
        ValTreeView.column("Part", anchor=CENTER, width=45)
        ValTreeView.column("Description", anchor=CENTER, width=180)
        ValTreeView.column("D", anchor=CENTER, width=30)
        ValTreeView.column("S", anchor=CENTER, width=30)
        ValTreeView.column("CLS", anchor=CENTER, width=60)
        ValTreeView.column("V", anchor=CENTER, width=30)
        ValTreeView.column("Maker", anchor=CENTER, width=100)
        ValTreeView.column("Spec", anchor=CENTER, width=200)
        ValTreeView.column("DES", anchor=CENTER, width=40)
        ValTreeView.column("SPA", anchor=CENTER, width=40)
        ValTreeView.column("OH", anchor=CENTER, width=40)
        ValTreeView.column("REQ", anchor=CENTER, width=40)
        ValTreeView.column("PCH", anchor=CENTER, width=40)
        ValTreeView.column("BAL", anchor=CENTER, width=40)
        ValTreeView.column("RCV", anchor=CENTER, width=40)
        ValTreeView.column("ADJ", anchor=CENTER, width=40)
        ValTreeView.column("Vendor", anchor=CENTER, width=100)
        ValTreeView.column("UnitCost", anchor=CENTER, width=80)
    
        ValTreeView.heading("#0", text="Index", anchor=W)
        ValTreeView.heading("Part", text="Part", anchor=CENTER)
        ValTreeView.heading("Description", text="Description", anchor=CENTER)
        ValTreeView.heading("D", text="D", anchor=CENTER)
        ValTreeView.heading("S", text="S", anchor=CENTER)
        ValTreeView.heading("CLS", text="CLS", anchor=CENTER)
        ValTreeView.heading("V", text="V", anchor=CENTER)
        ValTreeView.heading("Maker", text="Maker", anchor=CENTER)
        ValTreeView.heading("Spec", text="Maker Specification", anchor=CENTER)
        ValTreeView.heading("DES", text="DES", anchor=CENTER)
        ValTreeView.heading("SPA", text="SPA", anchor=CENTER)
        ValTreeView.heading("OH", text="OH", anchor=CENTER)
        ValTreeView.heading("REQ", text="REQ", anchor=CENTER)
        ValTreeView.heading("PCH", text="PCH", anchor=CENTER)
        ValTreeView.heading("BAL", text="BAL", anchor=CENTER)
        ValTreeView.heading("RCV", text="RCV", anchor=CENTER)
        ValTreeView.heading("ADJ", text="ADJ", anchor=CENTER)
        ValTreeView.heading("Vendor", text="Vendor", anchor=CENTER)
        ValTreeView.heading("UnitCost", text="Unit Cost", anchor=CENTER)
        
        def formatUnitNum(num):
            threeDigit = str(num).rjust(3, "0")
            UnitFullName = f"{AssemblyFullName}-{threeDigit}"
            return UnitFullName
        
        def queryTreeSelect():
            curPur = connPur.cursor()
            curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
            recLst = curPur.fetchall()
            connPur.commit()
            curPur.close() 

            for rec in recLst:
                SelectTreeView.insert(parent="", index=END, iid=rec[0], 
                                      values=(rec[1], rec[2], rec[3], rec[4], 
                                              rec[5], rec[6], rec[7], rec[8]))
                
            if recLst == []:
                TaxBox.config(state="normal")
                TaxBox.delete(0, END)
                TaxBox.insert(0, "NIL")
                CurrencyBox.config(state="normal")
                CurrencyBox.delete(0, END)
                CurrencyBox.insert(0, "NIL")
                CurrencyBox.config(state="readonly")
            
            else:
                TaxBox.config(state="normal")
                TaxBox.delete(0, END)
                TaxBox.insert(0, recLst[0][6])
                CurrencyBox.config(state="normal")
                CurrencyBox.delete(0, END)
                CurrencyBox.insert(0, recLst[0][9])
                CurrencyBox.config(state="readonly")
                

        
        def selectUnit():
            selIndex = ValTreeView.selection()
            if selIndex == ():
                pass
            else:
                curPur = connPur.cursor()
                selectUnitCom = f"""INSERT INTO `{PurOrderNumRef}` (
                PartNum, Description, Maker, Spec, 
                REQ, Tax, Vendor, UnitCost, Currency)
        
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                
                for index in selIndex:
                    rec = ValTreeView.item(index, "values")
                    valSelect = (formatUnitNum(rec[0]), rec[1], rec[6], rec[7], 
                                 rec[11], "7%", rec[16], rec[17], "SGD")
                    curPur.execute(selectUnitCom, valSelect)
                    connPur.commit()    
                curPur.close()
                
            SelectTreeView.delete(*SelectTreeView.get_children())
            queryTreeSelect()
    
        def deleteUnit():
            selected = SelectTreeView.selection()
            curPur = connPur.cursor()
            for i in selected:
                curPur.execute(f"DELETE from `{PurOrderNumRef}` WHERE oid={i}")
                
            connPur.commit()
            curPur.close() 
            SelectTreeView.delete(*SelectTreeView.get_children())
            queryTreeSelect()
            
        def closeTabPur():
            framePur.destroy()
            buttonLoadPur.config(state=NORMAL)
        
        
        
        def clearUnitData():
            ReqQtyBox.delete(0, END)
            UnitCostSelectBox.delete(0, END)
            
        def selectUnitData(e):
            selected = SelectTreeView.focus()
            curPur = connPur.cursor()
            curPur.execute(f"SELECT * from `{PurOrderNumRef}` WHERE oid={selected}")
            
            unitVal = curPur.fetchall()
            clearUnitData()
            
            ReqQtyBox.insert(0, unitVal[0][5])
            UnitCostSelectBox.insert(0, unitVal[0][8])
        
        def updateUnitData():
            sqlCommand = f"""UPDATE `{PurOrderNumRef}` SET
            REQ = %s,
            UnitCost = %s
            WHERE oid = %s
            """

            selected = SelectTreeView.focus()
            
            inputs = (ReqQtyBox.get(), UnitCostSelectBox.get(), selected)
            
            response = messagebox.askokcancel("Confirmation", "Confirm Update", parent=RepWin)
            if response == True:
                curPur = connPur.cursor()
                curPur.execute(sqlCommand, inputs)
                connPur.commit()
                curPur.close()

                clearUnitData()
                SelectTreeView.delete(*SelectTreeView.get_children())
                queryTreeSelect()
                
                messagebox.showinfo("Update Successful", 
                                    f"You Have Updated This Part", parent=RepWin) 
            else:
                pass

        def updateUnitDataClick(e):
            updateUnitData()

        def defaultCommon():
            curPur = connPur.cursor()
            curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
            recLst = curPur.fetchall()
            connPur.commit()
            curPur.close() 
            
            TaxBox.config(state="normal")
            TaxBox.delete(0, END)
            TaxBox.insert(0, recLst[0][6])
            CurrencyBox.config(state="normal")
            CurrencyBox.delete(0, END)
            CurrencyBox.insert(0, recLst[0][9])
            CurrencyBox.config(state="readonly")
        
        def updateCommon():
            sqlComTax = f"""UPDATE `{PurOrderNumRef}` SET
            Tax = %s,
            Currency = %s
            """

            inputs = (TaxBox.get(), CurrencyBox.get())
            
            response = messagebox.askokcancel("Confirm Update", "This will Update ALL ROWS", parent=RepWin)
            if response == True:
                curPur = connPur.cursor()
                curPur.execute(sqlComTax, inputs)
                connPur.commit()
                curPur.close()

                clearUnitData()
                SelectTreeView.delete(*SelectTreeView.get_children())
                queryTreeSelect()
                
                messagebox.showinfo("Update Successful", 
                                    f"You Have Updated Tax and Currency", parent=RepWin) 
            else:
                pass










        def genPurOrder():
            curCom = connCom.cursor()
            curCom.execute("SELECT * FROM COMPANY_MWA")
            companyInfo = curCom.fetchall()
            
            curPur = connPur.cursor() 
            curPur.execute(f"SELECT * FROM PUR_ORDER_LIST WHERE PurOrderNum = '{PurOrderNumRef}'")
            purOrderInfo = curPur.fetchall()
            
            curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
            purOrderUnit = curPur.fetchall()
            
            curVend = connVend.cursor()
            curVend.execute(f"SELECT * FROM VENDOR_LIST WHERE VENDOR_NAME = '{VendorRef}'")
            vendorInfo = curVend.fetchall()
            
            curPur.close()
            curVend.close()
            
            def totalCostCalc(Qty, OneCost):
                QtyNum = float(Qty) if Qty else 0
                OneCostNum = float(OneCost) if OneCost else 0
                total = QtyNum * OneCostNum
                totalDigit = str("{:.2f}".format(total))
                return totalDigit
            
            unitDataLst = [["Index", "Part No.", "Description", "Maker Spec", "Quantity", 
                            "Tax", "Unit Cost", "Total Cost"]]
            for i in range(0, len(purOrderUnit)):
                unitDataLst.append([str(i+1), purOrderUnit[i][1], purOrderUnit[i][2],
                                    purOrderUnit[i][4], purOrderUnit[i][5], 
                                    purOrderUnit[i][6], 
                                    str("{:.2f}".format(float(purOrderUnit[i][8]) if purOrderUnit[i][8] else 0)) + " " + str(purOrderUnit[0][9]), 
                                    str(totalCostCalc(purOrderUnit[i][5], purOrderUnit[i][8])) + " " + str(purOrderUnit[0][9])])

            class PurOrder(FPDF):
                def footer(self):
                    self.set_y(-15) # 15 mm above from bottom
                    self.set_font("Arial", "I", 10) # 10 Font Size
                    self.cell(0, 10, f"Page {self.page_no()} / {{nb}}", align="C")
                
                def companyDetail(self):
                    self.image("MWA.jpg", x=10, y=8, h=25)
                    self.set_font("Arial", "", 20)
                    self.cell(w=0, h=25, txt="PURCHASE ORDER", align="R", border=False, ln=True)
                    
                    self.set_font("Arial", "", 8)
                    self.cell(w=60, h=3, txt=f"Company Reg. No: {companyInfo[0][2]}", align="L")
                    self.cell(w=50, h=3, txt="Purchase Order Date", align="R")
                    self.cell(w=0, h=3, txt="Vendor", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"GST Reg. No: {companyInfo[0][3]}", align="L")
                    self.cell(w=50, h=3, txt=f"{purOrderInfo[0][3]}", align="R") # Order Date
                    self.cell(w=0, h=3, txt=f"{vendorInfo[0][3]}", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"{companyInfo[0][1]}", align="L")
                    self.cell(w=50, h=3, txt="", align="R")
                    self.cell(w=0, h=3, txt=f"{vendorInfo[0][8]}", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"{companyInfo[0][4]} {companyInfo[0][5]}", align="L")
                    self.cell(w=50, h=3, txt="Purchase Order No.", align="R")
                    self.cell(w=0, h=3, txt=f"{vendorInfo[0][9]}", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"{companyInfo[0][6]}", align="L")
                    self.cell(w=50, h=3, txt=f"{purOrderInfo[0][1]}", align="R")
                    self.cell(w=0, h=3, 
                              txt=f"{vendorInfo[0][6]} {vendorInfo[0][5]} Postal {vendorInfo[0][7]}", 
                              align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"{companyInfo[0][7]}", align="L")
                    self.cell(w=50, h=3, txt="", align="R")
                    self.cell(w=0, h=3, txt=f"{vendorInfo[0][4]}", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt=f"Contact No.: {companyInfo[0][8]}", align="L")
                    self.cell(w=50, h=3, txt="Payment Terms", align="R")
                    self.cell(w=0, h=3, txt=f"POC: {vendorInfo[0][11]}", align="R", ln=True)
                            
                    self.cell(w=60, h=3, txt=f"Email: {companyInfo[0][9]}", align="L")
                    self.cell(w=50, h=3, txt=f"{purOrderInfo[0][2]}", align="R")
                    self.cell(w=0, h=3, txt=f"Contact: {vendorInfo[0][12]}", align="R", ln=True)
                    
                    self.cell(w=60, h=3, txt="", align="L")
                    self.cell(w=50, h=3, txt="", align="R")
                    self.cell(w=0, h=3, txt=f"{vendorInfo[0][13]}", align="R", ln=True)
                    
                    self.ln(10)
                
                def columnWidth(self, lstRef): # Get the Maximum Width of Each Column  
                    pageWidthMargin = self.w - 20
                    fontSize = 10
                    while True:
                        self.set_font("Arial", "", fontSize)
                        colLenLst = []
                        for i in range(len(lstRef[0])): # Select Rows 1 2 3
                            strLenLst = []
                            for j in range(len(lstRef)): # Select Columns 1 2 3
                                val = lstRef[j][i]
                                strLen = self.get_string_width(val) + 5
                                strLenLst.append(strLen)
                            colLenLst.append(max(strLenLst))
                        if sum(colLenLst) <= pageWidthMargin:
                            break
                        else:
                            fontSize = fontSize - 0.5
                    return colLenLst
                
                def printHeading(self, lst):
                    self.set_fill_color(200, 220, 255)
                    for i in range(len(lst[0])):
                        strWidth = self.columnWidth(lst)[i]
                        self.cell(strWidth, 5, lst[0][i], border=True, ln=False, fill=True, align="C")
                
                def printContent(self, lst):
                    for i in range(1, len(lst)):
                        self.ln()
                        for j in range(len(lst[0])):
                            strWidth = self.columnWidth(lst)[j]
                            self.cell(strWidth, 5, lst[i][j], border=True)
                    self.ln(10)
                    
                def printTotal(self, lst):
                    costLst = []
                    for i in range(1, len(lst)):
                        val = re.sub("[^0-9^.]", "", lst[i][7])
                        costLst.append(float(val))
                    subTotal = sum(costLst)
                    GST = subTotal * float(str(purOrderUnit[0][6]).strip("%"))/100
                    totalSGD = subTotal + GST
                    
                    subTotalStr = str("{:.2f}".format(subTotal)) + " " + str(purOrderUnit[0][9])
                    GSTStr = str("{:.2f}".format(GST)) + " " + str(purOrderUnit[0][9])
                    totalStr = str("{:.2f}".format(totalSGD)) + " " + str(purOrderUnit[0][9])
                    
                    self.set_font("Arial", "", 10)
                    self.cell(w=185, h=5, txt=f"Subtotal   {subTotalStr}", align="R", ln=True)
                    self.cell(w=185, h=5, txt=f"GST   {GSTStr}", align="R", ln=True)
                    self.cell(w=185, h=5, txt=f"Total   {totalStr}", align="R", ln=True)
                    
                def printNotes(self):
                    self.set_font("Arial", "", 8)
                    self.cell(w=0, h=5, txt="NOTES:", ln=True, align="L")
                    self.cell(w=0, h=5, txt="1. Please acknowledge receipt & delivery date on our P.O. by chop & sign and email it back to us immediately.",
                              ln=True, align="L")
                    self.cell(w=0, h=5, txt="2. Purchase Order Number must indicate on all the delivery orders and invoices",
                              ln=True, align="L")
                    
                    self.ln(30)
                    
                    self.cell(w=100, h=5, txt="________________________________________", align="L")
                    self.cell(w=100, h=5, txt="______________________________", ln=True, align="L")
                    self.cell(w=100, h=5, txt="SUPPLIER hereby confirm acceptance of this Order", 
                              align="L")
                    self.cell(w=100, h=5, txt="MOTIONWELL Automation Pte. Ltd.",
                              ln=True, align="L")
                    
                    self.cell(w=100, h=5, txt="Name, Designation, Signature and Date", 
                              align="L")
                    self.cell(w=100, h=5, txt="Authorized Signature and Date",
                              ln=True, align="L")
        
            PurOrderGen = PurOrder("P", "mm", "A4")
            PurOrderGen.set_auto_page_break(auto=True, margin=15)
            PurOrderGen.add_page()
            
            PurOrderGen.companyDetail()
            PurOrderGen.printHeading(unitDataLst)
            PurOrderGen.printContent(unitDataLst)
            
            if PurOrderGen.get_y() >= 268:
                PurOrderGen.add_page()
            
            PurOrderGen.printTotal(unitDataLst)
            PurOrderGen.ln(20)
            
            if PurOrderGen.get_y() >= 218:
                PurOrderGen.add_page()
            
            PurOrderGen.printNotes()
            
            PurOrderGen.output(f"{PurOrderNumRef}.pdf")
            
            messagebox.showinfo("Create Successful", 
                                f"You Have Generated PO {PurOrderNumRef}", parent=framePur) 
        
        
        def ExportToXero():
            XeroExportWin = Toplevel()
            XeroExportWin.geometry('400x600')
            
            
        
        
        
        
        
        
        
    
        ButtonRepFrame = LabelFrame(framePur, text="Command")
        ButtonRepFrame.grid(row=2, column=0, padx=(50,10), pady=5, ipadx=10, ipady=5, sticky=W)
        
        DataRepFrame = LabelFrame(framePur, text="Data")
        DataRepFrame.grid(row=2, column=1, padx=(10,75), pady=(5,0), ipadx=10, ipady=5, sticky=W)
        
        TaxCurrencyFrame = LabelFrame(framePur, text="Tax & Currency")
        TaxCurrencyFrame.grid(row=3, column=1, padx=(10,75), pady=(0,5), ipadx=10, ipady=5, sticky=W)
        
        buttonSelectThis = Button(ButtonRepFrame, text="Select Unit (Top)", command=selectUnit)
        buttonSelectThis.grid(row=0, column=0, padx=10, pady=5, sticky=W)
        
        buttonDeleteThis = Button(ButtonRepFrame, text="Delete Unit (Below)", command=deleteUnit)
        buttonDeleteThis.grid(row=0, column=1, padx=10, pady=5, sticky=W)
        
        buttonGenPur = Button(ButtonRepFrame, text="Generate Purchase Order", command=genPurOrder)
        buttonGenPur.grid(row=0, column=2, padx=10, pady=5, sticky=W)
        
        buttonExportXero = Button(ButtonRepFrame, text="Export to Xero", command=ExportToXero)
        buttonExportXero.grid(row=0, column=3, padx=10, pady=5, sticky=W)
    
        buttonClosePur = Button(ButtonRepFrame, text="Close Tab", command=closeTabPur)
        buttonClosePur.grid(row=0, column=4, padx=10, pady=5, sticky=W)
        
        
        
        ReqQtyLabel = Label(DataRepFrame, width=8, text="Req Qty.")
        ReqQtyBox = Entry(DataRepFrame, width=8)
        UnitCostSelectLabel = Label(DataRepFrame, width=12, text="Unit Cost")
        UnitCostSelectBox = Entry(DataRepFrame, width=12)
        
        buttonClearData = Button(DataRepFrame, text="Clear", width=8, command=clearUnitData)
        buttonUpdateData = Button(DataRepFrame, text="Update", width=8, command=updateUnitData)
        
        TaxLabel = Label(TaxCurrencyFrame, width=8, text="Tax")
        TaxBox = Entry(TaxCurrencyFrame, width=8)
        CurrencyLabel = Label(TaxCurrencyFrame, width=12, text="Currency")
        CurrencyBox = ttk.Combobox(TaxCurrencyFrame, width=10, 
                                   value=["SGD", "USD", "CNY", "HKD", "EUR",
                                          "JPY", "GBP", "AUD", "CAD", "CHF",
                                          "Others"], state="readonly")

        buttonDefaultCommon = Button(TaxCurrencyFrame, text="Default", width=8, command=defaultCommon)
        buttonUpdateCommon = Button(TaxCurrencyFrame, text="Update", width=8, command=updateCommon)
        
        def CurrencySelect(e):
            if CurrencyBox.get() == "Others":
                CurrencyBox.config(state="normal")
                CurrencyBox.delete(0, END)
            else:
                CurrencyBox.config(state="readonly")
        
        ReqQtyLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
        ReqQtyBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)        
        UnitCostSelectLabel.grid(row=0, column=2, padx=(10,10), pady=5, sticky=E)
        UnitCostSelectBox.grid(row=0, column=3, padx=10, pady=5, sticky=W)
        
        buttonClearData.grid(row=0, column=4, padx=(20,5), pady=5, sticky=W)
        buttonUpdateData.grid(row=0, column=5, padx=(5,10), pady=5, sticky=W)

        TaxLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
        TaxBox.grid(row=1, column=1, padx=10, pady=5, sticky=E)
        CurrencyLabel.grid(row=1, column=2, padx=(10,10), pady=5, sticky=E)
        CurrencyBox.grid(row=1, column=3, padx=10, pady=5, sticky=E)
        
        buttonDefaultCommon.grid(row=1, column=4, padx=(20,5), pady=5, sticky=W)
        buttonUpdateCommon.grid(row=1, column=5, padx=(5,10), pady=5, sticky=W)

        CurrencyBox.bind("<<ComboboxSelected>>", CurrencySelect)




    
        SelectTreeFrame = Frame(framePur)
        SelectTreeFrame.grid(row=4, column=0, columnspan=2, padx=10, pady=(5,0), ipadx=10, ipady=0, sticky=W+E)
        # SelectTreeFrame.pack()
        
        SelectTreeScroll = Scrollbar(SelectTreeFrame)
        SelectTreeScroll.pack(side=RIGHT, fill=Y)
        
        SelectTreeView = ttk.Treeview(SelectTreeFrame, yscrollcommand=SelectTreeScroll.set, 
                                      height=8, selectmode="extended")
        SelectTreeScroll.config(command=SelectTreeView.yview)
        # SelectTreeView.grid(row=2, column=0, columnspan=10, padx=10)
        SelectTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
        
        SelectTreeView["columns"] = ("Part", "Description", "Maker", "Spec", 
                                     "REQ", "Tax", "Vendor", "UnitCost")
        
        # SelectTreeView.column("#0", anchor=CENTER, width=50)
        SelectTreeView.column("#0", width=0 ,stretch=NO)
        SelectTreeView.column("Part", anchor=CENTER, width=135)
        SelectTreeView.column("Description", anchor=CENTER, width=260)
        SelectTreeView.column("Maker", anchor=CENTER, width=140)
        SelectTreeView.column("Spec", anchor=CENTER, width=240)
        SelectTreeView.column("REQ", anchor=CENTER, width=120)
        SelectTreeView.column("Tax", anchor=CENTER, width=60)
        SelectTreeView.column("Vendor", anchor=CENTER, width=120)
        SelectTreeView.column("UnitCost", anchor=CENTER, width=100)
    
        SelectTreeView.heading("#0", text="Index", anchor=W)
        SelectTreeView.heading("Part", text="Part", anchor=CENTER)
        SelectTreeView.heading("Description", text="Description", anchor=CENTER)
        SelectTreeView.heading("Maker", text="Maker", anchor=CENTER)
        SelectTreeView.heading("Spec", text="Maker Specification", anchor=CENTER)
        SelectTreeView.heading("REQ", text="Require Quantity", anchor=CENTER)
        SelectTreeView.heading("Tax", text="Tax", anchor=CENTER)
        SelectTreeView.heading("Vendor", text="Vendor", anchor=CENTER)
        SelectTreeView.heading("UnitCost", text="Unit Cost", anchor=CENTER)

        queryTreeSelect()

        SelectTreeView.bind('<Double-Button-1>', selectUnitData)
        SelectTreeView.bind('<Return>', updateUnitDataClick)
        # DataRepFrame.bind('<Return>', updateUnitData)
        























    OrderDataFrame = LabelFrame(frameRep, text="Record")
    OrderDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # OrderDataFrame.pack(fill="x", expand="yes", padx=20)
    
    OrderButtonFrame = LabelFrame(frameRep, text="Command")
    OrderButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # OrderButtonFrame.pack(fill="x", expand="yes", padx=20)
    
    PurOrderNumLabel = Label(OrderDataFrame, text="Purchase Order Number")
    PaymentTermLabel = Label(OrderDataFrame, text="Payment Term")
    OrderDateLabel = Label(OrderDataFrame, text="Order Date")
    VendorRemarkLabel = Label(OrderDataFrame, text="Vendor")



    PurOrderNumBox = Entry(OrderDataFrame, width=20)
    OrderNumGenButton = Button(OrderDataFrame, text="Gen", width=5, command=genOrderNum)
    PaymentTermBox = Entry(OrderDataFrame, width=20, state="readonly")
    PaymentTermButton = Button(OrderDataFrame, text="List", width=5, command=payLstGen)

    OrderDateBox = Entry(OrderDataFrame, width=20, state="readonly")
    OrderDateCalendarButton = Button(OrderDataFrame, text="Cal", width=5, 
                                     command=orderCalPro)
    VendorRemarkBox = Entry(OrderDataFrame, width=20, state="readonly")
    VendorSelectButton = Button(OrderDataFrame, text="List", width=5, command=vendorLstGen)

    PurOrderNumLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    PaymentTermLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    OrderDateLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    VendorRemarkLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)

    PurOrderNumBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    OrderNumGenButton.grid(row=0, column=2, padx=(0,10), pady=5, sticky=W)
    
    PaymentTermBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    PaymentTermButton.grid(row=1, column=2, padx=(0,10), pady=5, sticky=W)
    
    OrderDateBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    OrderDateCalendarButton.grid(row=2, column=2, padx=(0,10), pady=5, sticky=W)
    
    VendorRemarkBox.grid(row=3, column=1, columnspan=4, padx=10, pady=5, sticky=W)
    VendorSelectButton.grid(row=3, column=2, columnspan=4, padx=(0,10), pady=5, sticky=W)

    
    buttonUpdatePur = Button(OrderButtonFrame, text="Update Order", command=updateOrder, state=DISABLED)
    buttonCreatePur = Button(OrderButtonFrame, text="Create New Purchase Order", command=createOrder)
    buttonDeletePur = Button(OrderButtonFrame, text="Delete Order", command=deleteOrder, state=DISABLED)
    buttonSelectPur = Button(OrderButtonFrame, text="Select Order", command=selectOrder)
    buttonDeselectPur = Button(OrderButtonFrame, text="Deselect Order", command=deselectOrder)
    buttonLoadPur = Button(OrderButtonFrame, text="Load Order", command=loadOrder)
    buttonClearEntryPur = Button(OrderButtonFrame, text="Clear Entry", command=clearEntryOrder)
    buttonRefreshPur = Button(OrderButtonFrame, text="Refresh", command=refreshOrder)

    buttonUpdatePur.grid(row=0, column=0, padx=10, pady=10)
    buttonCreatePur.grid(row=0, column=1, padx=10, pady=10)
    buttonDeletePur.grid(row=0, column=2, padx=10, pady=10)
    buttonSelectPur.grid(row=0, column=3, padx=10, pady=10)
    buttonDeselectPur.grid(row=0, column=4, padx=10, pady=10)
    buttonLoadPur.grid(row=0, column=5, padx=10, pady=10)
    buttonClearEntryPur.grid(row=0, column=6, padx=10, pady=10)
    buttonRefreshPur.grid(row=0, column=7, padx=10, pady=10)










    def editCompanyInfo():
        comWin = Toplevel()
        comWin.title("Edit Company Information")
        comWin.geometry("450x480")
        comWin.columnconfigure(0, weight=1)
        comWin.rowconfigure(0, weight=1)
        
        CompanyDataFrame = LabelFrame(comWin, text="Record")
        CompanyDataFrame.grid(row=0, column=0, padx=20, pady=(20,5), ipadx=5, ipady=5, sticky=W+E)
        # CompanyDataFrame.pack(fill="x", expand="yes", padx=20)
        
        CompanyButtonFrame = LabelFrame(comWin, text="Command")
        CompanyButtonFrame.grid(row=1, column=0, padx=20, pady=(5,20), ipadx=5, ipady=5, sticky=W+E)
        # CompanyButtonFrame.pack(fill="x", expand="yes", padx=20)

        ComNameLabel = Label(CompanyDataFrame, text="Company Name")
        ComRegNumLabel = Label(CompanyDataFrame, text="Company Reg No.")
        GSTRegNumLabel = Label(CompanyDataFrame, text="GST Reg No.")
        AddressLabel = Label(CompanyDataFrame, text="Address")
        PosCodeLabel = Label(CompanyDataFrame, text="Postal Code")
        CenterALabel = Label(CompanyDataFrame, text="Center A Info")
        CenterBLabel = Label(CompanyDataFrame, text="Center B Info")
        ContactNumLabel = Label(CompanyDataFrame, text="Contact No.")
        EmailLabel = Label(CompanyDataFrame, text="Contact Email")

        ComNameBox = Entry(CompanyDataFrame, width=30)
        ComRegNumBox = Entry(CompanyDataFrame, width=30)
        GSTRegNumBox = Entry(CompanyDataFrame, width=30)
        AddressBox = Entry(CompanyDataFrame, width=30)
        PosCodeBox = Entry(CompanyDataFrame, width=30)
        CenterABox = Entry(CompanyDataFrame, width=30)
        CenterBBox = Entry(CompanyDataFrame, width=30)
        ContactNumBox = Entry(CompanyDataFrame, width=30)
        EmailBox = Entry(CompanyDataFrame, width=30)

        ComNameLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
        ComRegNumLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
        GSTRegNumLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
        AddressLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)
        PosCodeLabel.grid(row=4, column=0, padx=10, pady=5, sticky=E)
        CenterALabel.grid(row=5, column=0, padx=10, pady=5, sticky=E)
        CenterBLabel.grid(row=6, column=0, padx=10, pady=5, sticky=E)
        ContactNumLabel.grid(row=7, column=0, padx=10, pady=5, sticky=E)
        EmailLabel.grid(row=8, column=0, padx=10, pady=5, sticky=E)
        
        ComNameBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
        ComRegNumBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
        GSTRegNumBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
        AddressBox.grid(row=3, column=1, padx=10, pady=5, sticky=W)
        PosCodeBox.grid(row=4, column=1, padx=10, pady=5, sticky=W)
        CenterABox.grid(row=5, column=1, padx=10, pady=5, sticky=W)
        CenterBBox.grid(row=6, column=1, padx=10, pady=5, sticky=W)
        ContactNumBox.grid(row=7, column=1, padx=10, pady=5, sticky=W)
        EmailBox.grid(row=8, column=1, padx=10, pady=5, sticky=W)
        
        def queryComInfo():
            curCom = connCom.cursor()
            curCom.execute("SELECT * FROM COMPANY_MWA")
            comInfo = curCom.fetchall()
            curCom.close()
            
            if comInfo == []:
                pass
            else:
                ComNameBox.insert(0, comInfo[0][1])
                ComRegNumBox.insert(0, comInfo[0][2])
                GSTRegNumBox.insert(0, comInfo[0][3])
                AddressBox.insert(0, comInfo[0][4])
                PosCodeBox.insert(0, comInfo[0][5])
                CenterABox.insert(0, comInfo[0][6])
                CenterBBox.insert(0, comInfo[0][7])
                ContactNumBox.insert(0, comInfo[0][8])
                EmailBox.insert(0, comInfo[0][9])
        
        def updateComInfo():
            curCom = connCom.cursor()
            curCom.execute("SELECT * FROM `COMPANY_MWA` WHERE oid = 1")
            result = curCom.fetchall()
            if result == []:
                createInfoCom = """ INSERT INTO `COMPANY_MWA`
                (ComName, ComRegNum, GSTRegNum, Address, PosCode, 
                 CenterA, CenterB, ContactNum, Email)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
            
                InfoComValue = (ComNameBox.get(), ComRegNumBox.get(), GSTRegNumBox.get(), 
                                AddressBox.get(), PosCodeBox.get(), CenterABox.get(),
                                CenterBBox.get(), ContactNumBox.get(), EmailBox.get())
                
                curCom.execute(createInfoCom, InfoComValue)
                connCom.commit()
                messagebox.showinfo("Update Successful", 
                                    "You Have Successfuly Updated Company Info", 
                                    parent=comWin)
                comWin.destroy()
            
            else:
                updateInfoCom = """ UPDATE `COMPANY_MWA` SET
                `ComName` = %s,
                `ComRegNum` = %s,
                `GSTRegNum` = %s,
                `Address` = %s,
                `PosCode` = %s,
                `CenterA` = %s,
                `CenterB` = %s,
                `ContactNum` = %s,
                `Email` = %s
                
                WHERE `oid` = 1"""
                
                UpdateComValue = (ComNameBox.get(), ComRegNumBox.get(), GSTRegNumBox.get(), 
                                  AddressBox.get(), PosCodeBox.get(), CenterABox.get(),
                                  CenterBBox.get(), ContactNumBox.get(), EmailBox.get())
                
                curCom.execute(updateInfoCom, UpdateComValue)
                connCom.commit()
                messagebox.showinfo("Update Successful", 
                                    "You Have Successfuly Updated Company Info", 
                                    parent=comWin)
                comWin.destroy()
            
        def clearEntryMWA():
            ComNameBox.delete(0, END)
            ComRegNumBox.delete(0, END)
            GSTRegNumBox.delete(0, END)
            AddressBox.delete(0, END)
            PosCodeBox.delete(0, END)
            CenterABox.delete(0, END)
            CenterBBox.delete(0, END)
            ContactNumBox.delete(0, END)
            EmailBox.delete(0, END)
        
        buttonUpdateComInfo = Button(CompanyButtonFrame, 
                                     text="Update Company Information", 
                                     command=updateComInfo)
        buttonUpdateComInfo.grid(row=0, column=0, padx=10, pady=10)
        buttonClearEntryInfo = Button(CompanyButtonFrame, 
                                     text="Clear Entry", 
                                     command=clearEntryMWA)
        buttonClearEntryInfo.grid(row=0, column=1, padx=10, pady=10)

        queryComInfo()
    


    menuRep = Menu(RepWin)
    RepWin.config(menu=menuRep)
        
    fileMenu = Menu(menuRep, tearoff=0)
    menuRep.add_cascade(label="File", menu=fileMenu)
    fileMenu.add_command(label="Edit Company Info", command=editCompanyInfo)
    fileMenu.add_separator()
    fileMenu.add_command(label="Exit", command=RepWin.destroy)

    queryTreeOrder()
    
    
    
if __name__ == '__main__':
    a = Tk()
    openPurchase()
    a.mainloop()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


























