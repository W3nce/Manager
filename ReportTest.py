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
import Login

#additional modules
from xerogui.app import *
from xero_python.accounting.models import PurchaseOrder,PurchaseOrders,LineItem,LineAmountTypes,CurrencyCode
from xero_python.api_client.serializer import serialize
from xero_python import models
import time
from AutoCombo import AutocompleteCombobox

import requests
import CountryRef
import os

# Login.AUTHLVL

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

# if Login.AUTHLVL == 0:
#     pass
# else:
#     RetrieveToken()

def openPurchase():
    RepWin = Toplevel()
    RepWin.iconbitmap("MWA_Icon.ico")
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
                     `TransCcy` VARCHAR(10),
                     `TransExRate` FLOAT,
                     `ProgressStat` INT DEFAULT 0,
                     `ApproveStat` INT DEFAULT 0,
                     `IssueStat` INT DEFAULT 0,
                     `OrderStat` INT DEFAULT 0,
                     `TotalSGD` FLOAT)
                    
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
                     `Address` VARCHAR(100),
                     `CenterA` VARCHAR(80),
                     `CenterB` VARCHAR(80),
                     `Building` VARCHAR(80),
                     `PosCode` VARCHAR(50),
                     `ComRegNum` VARCHAR(30),
                     `Buyer` VARCHAR(100),
                     `ContactNum` VARCHAR(30),
                     `Email` VARCHAR(100))
                    
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

    #0 BOM Cost
    #1 Export CSV
    #2 Total Cost
    #3 Delete Project
    #4 Edit Employee
    #5 Approve PurOrder # USED TO UNLOCK COMPLETED ORDER
    #6 Generate PurOrder
    
    def AuthLevel(Auth, index):
        AuthDic = {0:[0,0,0,0,0,0,0], 1:[1,1,0,0,0,0,1],
                   2:[1,1,1,1,0,0,0], 3:[1,1,1,1,1,1,1]}
        AuthBool = AuthDic.get(Auth)[index]
        return AuthBool

    # Insert Default Company Info
# =============================================================================
#     curCom = connCom.cursor()
#     curCom.execute("SELECT * FROM COMPANY_MWA")
#     existLst = curCom.fetchall()
#     
#     if existLst == []:
#         defaultComSql = f"""INSERT INTO `COMPANY_MWA` (
#         ComName, ComRegNum, GSTRegNum, Address, PosCode, 
#         CenterA, CenterB, ContactNum, Email)
#     
#         VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
#                     
#         defaultComInfo = ("Motionwell Automation Pte. Ltd.", "201435019E", "201435019E",
#                           "20 Woodlands Link", "738733", "#09-08 (Design & Assembly Center)", 
#                           "#07-26 (Manufacturing Shop)", "98506102", "info@motionwell.com.sg")
#         
#         curCom.execute(defaultComSql, defaultComInfo)
#         connCom.commit()
#     curCom.close()
# 
# =============================================================================




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
                                "TransCcy",
                                "OrderStatus",
                                "TotalSGD")
    
    OrderTreeView.column("#0", anchor = CENTER, width =80, minwidth = 0)
    # OrderTreeView.column("#0",  width=0, stretch=NO)
    OrderTreeView.column("PurOrderNum", anchor = CENTER, width = 130, minwidth = 50)
    OrderTreeView.column("PaymentTerm", anchor = CENTER, width= 100, minwidth = 50)
    OrderTreeView.column("OrderDate", anchor = CENTER, width = 100, minwidth = 50)
    OrderTreeView.column("VendorRemark", anchor = W, width = 370, minwidth = 50)
    OrderTreeView.column("TransCcy", anchor = CENTER, width = 100, minwidth = 50)
    OrderTreeView.column("OrderStatus", anchor = CENTER, width = 200, minwidth = 50)
    OrderTreeView.column("TotalSGD", anchor = CENTER, width = 120, minwidth = 50)
    
    OrderTreeView.heading("#0", text = "Index")
    OrderTreeView.heading("PurOrderNum", text = "Purchase Order No.")
    OrderTreeView.heading("PaymentTerm", text = "Payment Term")
    OrderTreeView.heading("OrderDate", text = "Order Date")
    OrderTreeView.heading("VendorRemark", text = "Vendor Remark")
    OrderTreeView.heading("TransCcy", text = "Trans. Ccy")
    OrderTreeView.heading("OrderStatus", text = "Status")
    OrderTreeView.heading("TotalSGD", text = "Total SGD")





    def deselectOrderClick(e):
        deselectOrder()
    
    def selectOrderClick(e):
        selectOrder()
    
    def updateOrderReturn(e):
        if buttonUpdatePur["state"] == "disabled":
            messagebox.showwarning("Unable to Update", 
                                   "Please Select an Order", parent=frameRep) 
        else:
            updateOrder()
    
    def deleteOrderDel(e):
        if buttonDeletePur["state"] == "disabled":
            messagebox.showwarning("Unable to Delete", 
                                   "Please Select an Order", parent=frameRep) 
        else:
            deleteOrder()

    def loadOrderClick(e):
        if buttonLoadPur["state"] == "normal":
            loadOrder()

    def exitOrderEsc(e):
        tabNum = tabNoteRep.index("current")
        if tabNum == 0:
            respExitRep = messagebox.askokcancel("Confirmation",
                                                 "Exit Purchase Order Window Now?",
                                                 parent=frameRep)
            if respExitRep == True:
                RepWin.destroy()
            else:
                pass

    OrderTreeView.bind("<Button-3>", deselectOrderClick)
    OrderTreeView.bind("<Double-Button-1>", selectOrderClick)
    OrderTreeView.bind("<Return>", updateOrderReturn)
    OrderTreeView.bind("<Delete>", deleteOrderDel) 
    OrderTreeView.bind("<Button-2>", loadOrderClick)
    OrderTreeView.bind("<Escape>", exitOrderEsc)





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
                        "NET 7", "NET 10", "NET 30", "NET 60", "NET 90", 
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
            selVal = VendorSelectionTree.selection()
            if selVal == ():
                messagebox.showerror("Unable to Select",
                                     "Please Select a Value", parent=VendWin)
            else:
                curVEND =VendorSelectionTree.item(selVal[0])['values']
    
        VendorSelectionTree.bind('<ButtonRelease-1>', selectItem) 
        
        SearchLabel = Label(SearchFrame, text = 'Search by Vendor :')
        SearchLabel.grid(row=0,column = 0, padx=5, pady=5)
        
        SearchEntry = Entry(SearchFrame)
        SearchEntry.grid(row=0,column = 1, padx=5, pady=5)
        
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
        calWin.geometry("270x260")
        
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
        
        orderYearLst = []
        orderMonthLst = []
        for rec in recLst:
            yearVal = f"Y20{rec[1][5:7]}"
            if yearVal not in orderYearLst:
                orderYearLst.append(yearVal)
            
            monthVal = rec[1][5:9]
            if monthVal not in orderMonthLst:
                orderMonthLst.append(monthVal)
                
        for yr in orderYearLst:
            OrderTreeView.insert(parent="", index=END, iid=yr, text=yr,
                                 values=("",))
        
        monthDic = {1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun",
                    7:"Jul", 8:"Aug", 9:"Sept", 10:"Oct", 11:"Nov", 12:"Dec"}
        
        for mt in orderMonthLst:
            yrInt = mt[0:2]
            yrStr = f"Y20{yrInt}"
            mtInt = mt[2:]
            mtStr = monthDic.get(int(mtInt))
            
            OrderTreeView.insert(parent=yrStr, index=END, iid=f"M{mt}", 
                                  text=mtStr, values=("",))

        for rec in recLst:
            boolLst = ["In Progress", "Awaiting Approval", 
                       "Awaiting Issue", "Issued", "Rejected"]
            curVend.execute(f"SELECT * FROM VENDOR_LIST WHERE VENDOR_NAME = '{rec[4]}'")
            vendorInfo = curVend.fetchall()
            if vendorInfo == []:
                addressFull = ""
            else:
                addressFull = f"{vendorInfo[0][3]} ({vendorInfo[0][2]}) {vendorInfo[0][6]} {vendorInfo[0][5]} {vendorInfo[0][7]} {vendorInfo[0][4]}"
            OrderTreeView.insert(parent=f"M{rec[1][5:9]}", index=END, iid=rec[0], text="", 
                                  values=(rec[1], rec[2], rec[3], addressFull, rec[5],
                                          boolLst[rec[10]], 
                                          f"{rec[11]} SGD"))            
        curVend.close()
        
    def updateOrderCcy():
        orderNumRef = PurOrderNumBox.get()
        TransExRateUsed = TransExRateBox.get()
        
        curPur = connPur.cursor()
        curPur.execute(f"SELECT * FROM `{orderNumRef}`")
        fullLst = curPur.fetchall()
        
        SelectOIDLst = []
        SGDCostLst = []
        for val in fullLst:
            SelectOIDLst.append(val[0])
            SGDCostLst.append(val[10])
        
        def calcTrans(Cost, ExRate):
            if Cost == None or Cost == "" or Cost == "None":
                return None
            else:
                costVal = float(Cost)
                exRateVal = 1/(float(ExRate))
                return round(costVal * exRateVal, 2)
        
        TransCostLst = []
        for val in SGDCostLst:
            TransCostLst.append(calcTrans(val, TransExRateUsed))

        updateCcyCost = f"""UPDATE `{orderNumRef}` SET
        `CostTrans` = %s
        
        WHERE `oid` = %s"""
                
        for i in range(len(SelectOIDLst)):
            inputs = (TransCostLst[i], SelectOIDLst[i])
            curPur.execute(updateCcyCost, inputs)
        connPur.commit()        

    def updateOrder():
        sqlCommand = f"""UPDATE PUR_ORDER_LIST SET
        `PaymentTerm` = %s,
        `OrderDate` = %s,
        `VendorRemark` = %s,
        `TransCcy` = %s,
        `TransExRate` = %s,
        `ProgressStat` = %s,
        `ApproveStat` = %s,
        `IssueStat` = %s,
        `OrderStat` = %s
        
        WHERE `oid` = %s"""
        
        selected = OrderTreeView.selection()[0]
        orderNumRef = PurOrderNumBox.get()
        
        def checkDateOrder(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        ProgressStat = ProgressStatBox.current()
        ApproveStat = ApproveStatBox.current()
        IssueStat = IssueStatBox.current()
        
        if ApproveStat == 2:
            OrderStat = 4
        else:
            if ProgressStat == 0:
                OrderStat = 0
            else:
                if ApproveStat == 0:
                    OrderStat = 1
                else:
                    if IssueStat == 0:
                        OrderStat = 2
                    else:
                        OrderStat = 3
        
        inputs = (PaymentTermBox.get(), checkDateOrder(OrderDateBox),
                  VendorRemarkBox.get(), TransCcyBox.get(), TransExRateBox.get(),
                  ProgressStat, ApproveStat, IssueStat, OrderStat, selected)
        
        respUpdateOrder = messagebox.askokcancel("Confirmation",
                                                 "Update This Order?",
                                                 parent=frameRep)
        if respUpdateOrder == True:
            curPur = connPur.cursor()
            curPur.execute(sqlCommand, inputs)
            connPur.commit()
            curPur.close()
            
            updateOrderCcy()
            
            clearEntryOrder()
            OrderTreeView.delete(*OrderTreeView.get_children())
            queryTreeOrder()
            
            messagebox.showinfo("Update Successful", 
                                f"You Have Updated Order {orderNumRef}", parent=RepWin)
        
        else:
            pass
    
    def createOrderCom():
        # timeNow = datetime.now()
        # formatDate = timeNow.strftime("%Y-%m-%d")
        
        # OrderStat 
        # 0 In Progress
        # 1 Awaiting Approval
        # 2 Awaiting Issue
        # 3 Issued
        # 4 Rejected
        
        createComOrder = f"""INSERT INTO `PUR_ORDER_LIST` (
        PurOrderNum, PaymentTerm, OrderDate, VendorRemark, TransCcy, 
        TransExRate, ProgressStat, ApproveStat, IssueStat, OrderStat)
        
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
        
        def checkDateOrder(dateVar):
            if dateVar.get() == "":
                return Noneu
            else:
                return dateVar.get()
        
        ProgressStat = ProgressStatBox.current()
        ApproveStat = ApproveStatBox.current()
        IssueStat = IssueStatBox.current()
        
        if ApproveStat == 2:
            OrderStat = 4
        else:
            if ProgressStat == 0:
                OrderStat = 0
            else:
                if ApproveStat == 0:
                    OrderStat = 1
                else:
                    if IssueStat == 0:
                        OrderStat = 2
                    else:
                        OrderStat = 3
        
        valueOrder = (PurOrderNumBox.get(), PaymentTermBox.get(), 
                      checkDateOrder(OrderDateBox),
                      VendorRemarkBox.get(), TransCcyBox.get(),
                      TransExRateBox.get(), ProgressStat, ApproveStat, 
                      IssueStat, OrderStat)
        
        
        
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
            `Currency` VARCHAR(20),
            `CostSGD` VARCHAR(20),
            `CostTrans` VARCHAR(20))
        
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
        
    def createOrder():
        if PurOrderNumBox.get() == "":
            messagebox.showerror("Unable to Create",
                                   "Please Enter an Order Number",
                                   parent=frameRep)
        else:
            if VendorRemarkBox.get() == "":
                respCreateVen = messagebox.askokcancel("Yet to Choose Vendor",
                                                       "Create an Order without Vendor?",
                                                       parent=frameRep)
                if respCreateVen == True:
                    createOrderCom()
                else:
                    pass
            else:
                respCreateOrder = messagebox.askokcancel("Confirmation",
                                                         "Create This Order?",
                                                         parent=frameRep)
                if respCreateOrder == True:
                    createOrderCom()
                else:
                    pass
    
    def deleteOrder():
        selected = OrderTreeView.selection()[0]
        if selected[0] == "Y":
            respDelOrderYear = messagebox.askokcancel("You Have Selected a Year", 
                                                       f"This will Delete EVERYTHING under {selected}", 
                                                       parent=frameRep)
            if respDelOrderYear == True:
                selectAll = []
                monthNum = OrderTreeView.get_children(selected)
                for val in monthNum:
                    PurOrderVal = OrderTreeView.get_children(val)
                    for item in PurOrderVal:
                        selectAll.append(item)

                for index in selectAll:
                    sqlSelect = "SELECT * FROM PUR_ORDER_LIST WHERE oid = %s"
                    valSelect = (index, )
                    
                    curPur = connPur.cursor()
                    curPur.execute(sqlSelect, valSelect)
                    recVal = curPur.fetchall()
                    connPur.commit()
                    
                    orderNumRef = recVal[0][1]
                    
                    sqlDelete = "DELETE FROM PUR_ORDER_LIST WHERE oid = %s"
                    valDelete = (index,)
                    
                    curPur.execute(sqlDelete, valDelete)
                    connPur.commit()
                    
                    curPur.execute(f"DROP TABLE IF EXISTS `{orderNumRef}`")
                    connPur.commit()
                    
                    clearEntryOrder()
                    OrderTreeView.delete(*OrderTreeView.get_children())
                    queryTreeOrder()
                    
                messagebox.showinfo("Delete Successful", 
                                    f"You Have Deleted {len(selectAll)} Purchase Order", parent=RepWin) 
                curPur.close()
            else:
                pass
            
        elif selected[0] == "M":
            respDelOrderMonth = messagebox.askokcancel("You Have Selected a Month", 
                                                       f"This will Delete EVERYTHING under Month {selected[3:]}", 
                                                       parent=frameRep)
            if respDelOrderMonth == True:
                selectAll = OrderTreeView.get_children(selected)
                
                for index in selectAll:
                    sqlSelect = "SELECT * FROM PUR_ORDER_LIST WHERE oid = %s"
                    valSelect = (index, )
                    
                    curPur = connPur.cursor()
                    curPur.execute(sqlSelect, valSelect)
                    recVal = curPur.fetchall()
                    connPur.commit()
                    
                    orderNumRef = recVal[0][1]
                    
                    sqlDelete = "DELETE FROM PUR_ORDER_LIST WHERE oid = %s"
                    valDelete = (index,)
                    
                    curPur.execute(sqlDelete, valDelete)
                    connPur.commit()
                    
                    curPur.execute(f"DROP TABLE IF EXISTS `{orderNumRef}`")
                    connPur.commit()
                    
                    clearEntryOrder()
                    OrderTreeView.delete(*OrderTreeView.get_children())
                    queryTreeOrder()
                    
                messagebox.showinfo("Delete Successful", 
                                    f"You Have Deleted {len(selectAll)} Purchase Order", parent=RepWin) 
                curPur.close()
            else:
                pass
        
        else: 
            sqlDelete = "DELETE FROM PUR_ORDER_LIST WHERE oid = %s"
            valDelete = (selected, )
            
            respDelOrder = messagebox.askokcancel("Confirmation",
                                                  "Delete This Order",
                                                  parent=frameRep)
            
            if respDelOrder == True:
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
            else:
                pass
    
    def selectOrderCom():
        try:
            selected = OrderTreeView.selection()[0]     
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
            
            TransCcyBox.config(state="normal")
            TransCcyBox.delete(0, END)
            TransCcyBox.insert(0, recLst[0][5])
            TransCcyBox.config(state="readonly")
            
            TransExRateBox.config(state="normal")
            TransExRateBox.delete(0, END)
            TransExRateBox.insert(0, recLst[0][6])
            TransExRateBox.config(state="readonly")
            
            ProgressStatBox.current(recLst[0][7])
            ApproveStatBox.current(recLst[0][8])
            IssueStatBox.current(recLst[0][9])
            
            boolDic = {0:"In Progress", 1:"Awaiting Approval", 
                       2:"Awaiting Issue", 3:"Issued", 4:"Rejected"}
            
            OrderStatBox.config(state="normal")
            OrderStatBox.delete(0, END)
            OrderStatBox.insert(0, boolDic.get(recLst[0][10]))
            OrderStatBox.config(state="readonly")
            
            TotalSGDBox.config(state="normal")
            TotalSGDBox.delete(0, END)
            TotalSGDBox.insert(0, recLst[0][11])
            TotalSGDBox.config(state="readonly")


            
            buttonUpdatePur.config(state=NORMAL)
            buttonCreatePur.config(state=DISABLED)
            buttonDeletePur.config(state=NORMAL)
            OrderTreeView.config(selectmode="none")
            
        except:
            clearEntryOrder()
            messagebox.showerror("Error", "Please Check Again", parent=frameRep)
    
    def selectOrder():
        selVal = OrderTreeView.selection()
        if selVal == ():
            messagebox.showerror("Unable to Select",
                                 "Please Select a Value",
                                 parent=frameRep)
        else:
            selected = selVal[0]
            
            if selected[0] == "Y":
                respOrderYear = messagebox.askokcancel("You Have Selected a Year", 
                                                       f"This will select EVERYTHING under {selected}", 
                                                       parent=frameRep)
                if respOrderYear == True:
                    selectAll = []
                    monthNum = OrderTreeView.get_children(selected)
                    for val in monthNum:
                        PurOrderVal = OrderTreeView.get_children(val)
                        for item in PurOrderVal:
                            selectAll.append(item)
                    messagebox.showinfo("Selected", f"{len(selectAll)} Items under {selected} selected",
                                        parent=frameRep)
                    
                    buttonUpdatePur.config(state=DISABLED)
                    buttonDeletePur.config(state=NORMAL)
                    buttonCreatePur.config(state=DISABLED)
                    
                    OrderNumGenButton.config(state=DISABLED)
                    PurOrderNumBox.delete(0, END)
                    PurOrderNumBox.insert(0, f"{len(selectAll)} Items in {selected}")
                    OrderTreeView.config(selectmode="none")
                else:
                    pass
            elif selected[0] == "M":
                respOrderMonth = messagebox.askokcancel("You Have Selected a Month", 
                                                        f"This will select EVERYTHING under Month {selected[3:]}", 
                                                        parent=frameRep)
                if respOrderMonth == True:
                    selectAll = OrderTreeView.get_children(selected)
                    messagebox.showinfo("Selected", f"{len(selectAll)} Items under Month {selected[3:]} selected",
                                        parent=frameRep)
                    
                    buttonUpdatePur.config(state=DISABLED)
                    buttonDeletePur.config(state=NORMAL)
                    buttonCreatePur.config(state=DISABLED)
                    
                    OrderNumGenButton.config(state=DISABLED)
                    PurOrderNumBox.delete(0, END)
                    PurOrderNumBox.insert(0, f"{len(selectAll)} Items in Month {selected [3:]}")
                    OrderTreeView.config(selectmode="none")
                else:
                    pass
            else:
                selectOrderCom()
    
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
        
        TransCcyBox.current(0)
        TransExRateBox.config(state="normal")
        TransExRateBox.delete(0, END)
        TransExRateBox.insert(0, 1.0)
        TransExRateBox.config(state="readonly")        
        
        VendorRemarkBox.config(state="normal")
        VendorRemarkBox.delete(0, END)
        VendorRemarkBox.config(state="readonly")
        
        ProgressStatBox.current(0)
        ApproveStatBox.current(0)
        IssueStatBox.current(0)

        OrderStatBox.config(state="normal")
        OrderStatBox.delete(0, END)
        OrderStatBox.config(state="readonly")
        
        TotalSGDBox.config(state="normal")
        TotalSGDBox.delete(0, END)
        TotalSGDBox.config(state="readonly")
        
        OrderTreeView.config(selectmode="browse")
    
    def refreshOrder():
        clearEntryOrder()
        OrderTreeView.delete(*OrderTreeView.get_children())
        queryTreeOrder()
        
    def loadOrder():
        selected = OrderTreeView.selection()
        if selected == ():
            messagebox.showerror("Unable to Load",
                                 "Please Select an Order",
                                 parent=frameRep)
        
        else:
            if selected[0][0] == "Y":
                messagebox.showwarning("Unable to Load", "You Have Selected a Year",
                                        parent=frameRep)
            elif selected[0][0] == "M":
                messagebox.showwarning("Unable to Load", "You Have Selected a Month",
                                        parent=frameRep)
            else:
                loadOrderCom()
    
    def loadOrderCom():
        buttonLoadPur.config(state=DISABLED)
        
        curPur = connPur.cursor()
        selected = OrderTreeView.selection()[0]
        sqlSelect = "SELECT * FROM PUR_ORDER_LIST WHERE oid = %s"
        valSelect = (selected, )

        curPur.execute(sqlSelect, valSelect)
        purSelect = curPur.fetchall()
        
        connPur.commit()
        curPur.close() 

        PurOrderNumRef = purSelect[0][1]
        VendorRef = purSelect[0][4]
        TransCcyUsed = purSelect[0][5]
        loadOrderSelect = OrderTreeView.selection()[0]
        
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
                                                rec[14], rec[15], rec[17], rec[18], rec[20]))
                    
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
        
        ValTreeView["columns"] = ("Part", "Description", "D", "CLS", "V", 
                                    "Maker", "Spec", "DES", "SPA", "OH", "REQ", 
                                    "PCH", "BAL", "RCV", "OS", "Vendor", "UnitCost",
                                    "Currency")
        
        # ValTreeView.column("#0", anchor=CENTER, width=50)
        ValTreeView.column("#0", width=0 ,stretch=NO)
        ValTreeView.column("Part", anchor=CENTER, width=45)
        ValTreeView.column("Description", anchor=W, width=160)
        ValTreeView.column("D", anchor=CENTER, width=30)
        ValTreeView.column("CLS", anchor=CENTER, width=60)
        ValTreeView.column("V", anchor=CENTER, width=30)
        ValTreeView.column("Maker", anchor=CENTER, width=100)
        ValTreeView.column("Spec", anchor=W, width=180)
        ValTreeView.column("DES", anchor=CENTER, width=40)
        ValTreeView.column("SPA", anchor=CENTER, width=40)
        ValTreeView.column("OH", anchor=CENTER, width=40)
        ValTreeView.column("REQ", anchor=CENTER, width=40)
        ValTreeView.column("PCH", anchor=CENTER, width=40)
        ValTreeView.column("BAL", anchor=CENTER, width=40)
        ValTreeView.column("RCV", anchor=CENTER, width=40)
        ValTreeView.column("OS", anchor=CENTER, width=40)
        ValTreeView.column("Vendor", anchor=CENTER, width=100)
        ValTreeView.column("UnitCost", anchor=CENTER, width=80)
        ValTreeView.column("Currency", anchor=CENTER, width=60)
    
        ValTreeView.heading("#0", text="Index", anchor=CENTER)
        ValTreeView.heading("Part", text="Part", anchor=CENTER)
        ValTreeView.heading("Description", text="Description", anchor=CENTER)
        ValTreeView.heading("D", text="D", anchor=CENTER)
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
        ValTreeView.heading("OS", text="OS", anchor=CENTER)
        ValTreeView.heading("Vendor", text="Vendor", anchor=CENTER)
        ValTreeView.heading("UnitCost", text="Unit Cost", anchor=CENTER)
        ValTreeView.heading("Currency", text="Currency", anchor=CENTER)
        

        
        def formatUnitNum(num):
            threeDigit = str(num).rjust(3, "0")
            UnitFullName = f"{AssemblyFullName}-{threeDigit}"
            return UnitFullName
        
        def queryTreeSelect():
            curPur = connPur.cursor()
            curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
            recLst = curPur.fetchall()
            
            curPur.execute(f"SELECT * FROM `PUR_ORDER_LIST` WHERE PurOrderNum = '{PurOrderNumRef}'")
            TransCcyVal = curPur.fetchall()[0][5]
            
            connPur.commit()
            curPur.close() 

            for rec in recLst:
                SelectTreeView.insert(parent="", index=END, iid=rec[0], 
                                      values=(rec[1], rec[2], rec[3], rec[4], 
                                              rec[5], f"{rec[6]} %", rec[7], 
                                              checkCurrencyNonePur(rec[8], rec[9]),
                                              f"{rec[10]} SGD", 
                                              f"{rec[11]} {TransCcyVal}"))
        
        def CcyToSGD(Cost, Ccy):
            if Ccy == "SGD":
                ExRate = 1.00
            elif Ccy in CountryRef.getCcyLst():
                ExRate = CountryRef.getExRate(Ccy)
            else:
                ExRate = 0.00
            
            if Cost == None or Cost == "" or Cost == "None":
                return 0.00
            else:
                costVal = float(Cost)
                return round(costVal * ExRate, 2)
            
        def SGDToCcy(Cost, Ccy):
            if Ccy == "SGD":
                ExRate = 1.00
            elif Ccy in CountryRef.getCcyLst():
                ExRate = 1/(CountryRef.getExRate(Ccy))
            else:
                ExRate = 0.00
            
            if Cost == None or Cost == "" or Cost == "None":
                return 0.00
            else:
                costVal = float(Cost)
                return round(costVal * ExRate, 2)
        
        def checkCostPur(cost):
            if cost == "" or cost == "None" or cost == None:
                return None
            else:
                return float(cost)
        
        def checkCurrencyNonePur(num, ccy):
            if num == None or num == "" or num == "None":
                return num
            else:
                numCurrency = f"{num} {ccy}"
                return numCurrency


        
        def checkOrderTotal():
            curPur = connPur.cursor()
            curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
            recLst = curPur.fetchall()
            
            def checkInt(val):
                try:
                    float(val)
                    return True
                except:
                    return False
            
            def sumSingleCost(Qty, Cost):
                if checkInt(Qty) == False:
                    Qty = 0
                if checkInt(Cost) == False:
                    Cost = 0.0
                
                singleCostTotal = float(Qty) * float(Cost)
                return singleCostTotal
            
            
            
            costOrderLst = []
            for unit in recLst:
                costVal = sumSingleCost(unit[5], unit[10])
                costOrderLst.append(costVal)
            
            sumCostOrder = sum(costOrderLst)
            
            sqlOrderCost = f"""UPDATE PUR_ORDER_LIST SET
            `TotalSGD` = %s
            
            WHERE `oid` = %s"""
            
            inputs = (sumCostOrder, loadOrderSelect)
            
            curPur.execute(sqlOrderCost, inputs)
            connPur.commit()
            curPur.close()
            
            clearEntryOrder()
            OrderTreeView.delete(*OrderTreeView.get_children())
            queryTreeOrder()


        
        def selectUnitPur():
            selIndex = ValTreeView.selection()
            if selIndex == ():
                messagebox.showerror("Unable to Select",
                                     "Please Select a Part",
                                     parent=framePur)
            else:
                respSelectUnitTree = messagebox.askokcancel("Confirmation",
                                                            f"Select {len(selIndex)} Units?",
                                                            parent=framePur)
                if respSelectUnitTree == True:
                    curPur = connPur.cursor()
                    selectUnitCom = f"""INSERT INTO `{PurOrderNumRef}` (
                    PartNum, Description, Maker, Spec, 
                    REQ, Tax, Vendor, UnitCost, Currency,
                    CostSGD, CostTrans)
            
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                    
                    for index in selIndex:
                        rec = ValTreeView.item(index, "values")
                        
                        SGDVal = CcyToSGD(rec[16], rec[17])
                        TransVal = SGDToCcy(SGDVal, TransCcyUsed)
                        
                        valSelect = (formatUnitNum(rec[0]), rec[1], rec[5], rec[6], 
                                     rec[12], "7", rec[15], checkCostPur(rec[16]), 
                                     rec[17], SGDVal, TransVal)
                        
                        curPur.execute(selectUnitCom, valSelect)
                        connPur.commit()
                    curPur.close()
                    
                    checkOrderTotal()
                    SelectTreeView.delete(*SelectTreeView.get_children())
                    queryTreeSelect()
                else:
                    pass
    
        def deleteUnitPur():
            selDelete = SelectTreeView.selection()
            if selDelete == ():
                messagebox.showerror("Unable to Delete",
                                     "Please Select a Part",
                                     parent=framePur)
            else:
                respDeleteUnitTree = messagebox.askokcancel("Confirmation",
                                                            f"Delete {len(selDelete)} Units?",
                                                            parent=framePur)
                if respDeleteUnitTree == True:
                    curPur = connPur.cursor()
                    for i in selDelete:
                        curPur.execute(f"DELETE from `{PurOrderNumRef}` WHERE oid={i}")

                    connPur.commit()
                    curPur.close() 
                    
                    checkOrderTotal()
                    SelectTreeView.delete(*SelectTreeView.get_children())
                    queryTreeSelect()
                else:
                    pass
            
        def closeTabPur():
            respCloseTabPur = messagebox.askokcancel("Confirmation",
                                                     "Close This Tab?",
                                                     parent=framePur)
            if respCloseTabPur == True:
                framePur.destroy()
                buttonLoadPur.config(state=NORMAL)
            else:
                pass
        

        
        def clearUnitData():
            ReqQtyBox.delete(0, END)
            UnitCostSelectBox.delete(0, END)
            TaxBox.config(state="normal")
            TaxBox.delete(0, END)
            TaxBox.insert(0, 0)
            TaxBox.config(state="readonly")
            CurrencyBox.current(0)
            SelectTreeView.config(selectmode="extended")
            
        def selectUnitData(e):
            selVal = SelectTreeView.selection()
            if selVal == ():
                messagebox.showerror("Error",
                                     "Please Select a Part",
                                     parent=framePur)
            elif len(selVal) > 1:
                messagebox.showerror("Error",
                                     "Please Select ONLY ONE Part",
                                     parent=framePur)
            else:
                selected = selVal[0]
                curPur = connPur.cursor()
                curPur.execute(f"SELECT * from `{PurOrderNumRef}` WHERE oid={selected}")
                
                unitVal = curPur.fetchall()
                clearUnitData()
                
                TaxBox.delete(0, END)
                CurrencyBox.config(state="normal")
                CurrencyBox.delete(0, END)
                CurrencyBox.config(state="readonly")
                
                ReqQtyBox.insert(0, unitVal[0][5])
                
                if unitVal[0][8] == None or unitVal[0][8] == "None":
                    UnitCostSelectBox.insert(0, "")
                else:
                    UnitCostSelectBox.insert(0, unitVal[0][8])
                
                TaxBox.config(state="normal")
                TaxBox.delete(0, END)
                TaxBox.insert(0, unitVal[0][6])
                TaxBox.config(state="readonly")
                
                CurrencyBox.config(state="normal")
                CurrencyBox.insert(0, unitVal[0][9])
                CurrencyBox.config(state="readonly")
                
                SelectTreeView.config(selectmode="none")
            
        def deselectUnitData(e):
            selected = SelectTreeView.selection()
            if len(selected) > 0:
                for i in range(len(selected)):
                    SelectTreeView.selection_remove(selected[i])
                clearUnitData()
            else:
                clearUnitData()
                
        def deselectUnitValue(e):
            selected = ValTreeView.selection()
            if len(selected) > 0:
                for i in range(len(selected)):
                    ValTreeView.selection_remove(selected[i])
            else:
                pass
        
        def refreshUnitData():
            clearUnitData()
            SelectTreeView.delete(*SelectTreeView.get_children())
            queryTreeSelect()
        
        def updateUnitData():
            sqlCommand = f"""UPDATE `{PurOrderNumRef}` SET
            REQ = %s,
            UnitCost = %s,
            Tax = %s,
            Currency = %s,
            CostSGD = %s,
            CostTrans = %s
            
            WHERE oid = %s
            """
            
            selVal = SelectTreeView.selection()
            if len(selVal) != 1:
                messagebox.showerror("Error",
                                     "Please Select One Part ONLY",
                                     parent=framePur)
            else:
                selected = selVal[0]
                
                newTotalSGD = CcyToSGD(UnitCostSelectBox.get(), CurrencyBox.get())
                newTotalCcy = SGDToCcy(newTotalSGD, TransCcyUsed)
                
                inputs = (ReqQtyBox.get(), checkCostPur(UnitCostSelectBox.get()), 
                          TaxBox.get(), CurrencyBox.get(), 
                          newTotalSGD, newTotalCcy, selected)
                
                response = messagebox.askokcancel("Confirmation", "Confirm Update", parent=RepWin)
                if response == True:
                    curPur = connPur.cursor()
                    curPur.execute(sqlCommand, inputs)
                    connPur.commit()
                    curPur.close()
    
                    clearUnitData()
                    
                    checkOrderTotal()
                    SelectTreeView.delete(*SelectTreeView.get_children())
                    queryTreeSelect()
                    
                    messagebox.showinfo("Update Successful", 
                                        f"You Have Updated This Part", parent=RepWin) 
                else:
                    pass

        def updateUnitDataClick(e):
            updateUnitData()











        def checkIssueStat():
            curPur = connPur.cursor()

            sqlOrderStat = f"""UPDATE PUR_ORDER_LIST SET
            `IssueStat` = %s,
            `OrderStat` = %s

            WHERE `oid` = %s"""
            
            StatInput = (1, 3, loadOrderSelect)
            
            curPur.execute(sqlOrderStat, StatInput)
            connPur.commit()
            curPur.close()
            
            clearEntryOrder()
            OrderTreeView.delete(*OrderTreeView.get_children())
            queryTreeOrder()
            
        def genPurOrder():
            if AuthLevel(Login.AUTHLVL, 6) == False:
                messagebox.showerror("Insufficient Clearance",
                                     "You are NOT authorized to generate Purchase Order",
                                     parent=RepWin)
            else:
                curPur = connPur.cursor() 
                curPur.execute(f"SELECT * FROM PUR_ORDER_LIST WHERE PurOrderNum = '{PurOrderNumRef}'")
                PurInfo = curPur.fetchall()
                curPur.close()
                
                if PurInfo[0][8] == 0:
                    messagebox.showerror("Unable to Generate",
                                         "This order has YET to be Approved",
                                         parent=RepWin)
                elif PurInfo[0][8] == 2:
                    messagebox.showerror("Unable to Generate",
                                         "This order has been REJECTED",
                                         parent=RepWin)
                
                else:
                    if PurInfo[0][9] == 1:
                        respGenOrder = messagebox.askokcancel("Already Generated Before",
                                                              "Generate this order AGAIN?",
                                                              parent=RepWin)
                        if respGenOrder == True:
                            genPurOrderCom()
                        else:
                            pass
                    else:
                        genPurOrderCom()

        def genPurOrderCom():
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
            
            curCom.close()
            curPur.close()
            curVend.close()
            
            fullPartLst = []
            ProLst = []
            MachAssemLst = []
            PartLst = []
            DescLst = []
            QtyLst = []
            
            for val in purOrderUnit:
                fullPartLst.append(val[1])
                ProLst.append(val[1][0:6])
                MachAssemLst.append(f"{val[1][7:9]}_{val[1][10:13]}")
                PartLst.append(val[1][14:])
                DescLst.append(val[2])
                QtyLst.append(val[5])
            
            connStock = mysql.connector.connect(host = logininfo[0],
                                                user = logininfo[1], 
                                                password =logininfo[2],
                                                database= "STOCK_MASTER")
            
            connData = mysql.connector.connect(host = logininfo[0],
                                               user = logininfo[1], 
                                               password =logininfo[2])
            curData = connData.cursor()
            for i in range(len(fullPartLst)):
                curData.execute(f"SELECT * FROM `{ProLst[i]}`.`{MachAssemLst[i]}` WHERE PartNum = {PartLst[i]}")
                dataFetch = curData.fetchall()
                
                oidVal = dataFetch[0][0]
                BomREQ = int(dataFetch[0][11])
                BomPCH = int(dataFetch[0][12])
                BomBAL = int(dataFetch[0][13])
                BomRCV = int(dataFetch[0][14])
                                
                PurchaseQty = int(QtyLst[i]) + BomPCH
                BalanceQty = BomBAL - int(QtyLst[i])
                OutstandingQty = PurchaseQty - BomRCV
                
                # print(f"BomPCH = {BomPCH}")
                # print(f"BomBAL = {BomBAL}")
                # print(f"BomRCV = {BomRCV}")
                # print(f"PurchaseQty = {PurchaseQty}")
                # print(f"BalanceQty = {BalanceQty}")
                # print(f"OutstandingQty = {OutstandingQty}")
                
                if BalanceQty >= 0:
                    sqlUpdateBom = f"""UPDATE `{ProLst[i]}`.`{MachAssemLst[i]}` SET
                    PCH = %s,
                    BAL = %s,
                    OS = %s

                    WHERE oid = %s
                    """

                    PCHInput = (PurchaseQty, BalanceQty, OutstandingQty, oidVal)

                    curData.execute(sqlUpdateBom, PCHInput)
                    connData.commit()
                
                elif BalanceQty < 0:
                    sqlUpdateBom = f"""UPDATE `{ProLst[i]}`.`{MachAssemLst[i]}` SET
                    PCH = %s,
                    BAL = %s,
                    OS = %s

                    WHERE oid = %s
                    """
                    
                    PCHMaxQty = BomREQ
                    OSMaxQty = PCHMaxQty - BomRCV

                    PCHInput = (BomREQ, 0, OSMaxQty, oidVal)

                    curData.execute(sqlUpdateBom, PCHInput)
                    connData.commit()
                    
                    StockPartNum = fullPartLst[i]
                    StockDesc = DescLst[i]
                    StockQty = abs(BalanceQty)
                    StockOrderNum = purOrderInfo[0][1]
                    
                    timeNow = datetime.now()
                    formatDate = timeNow.strftime("%Y-%m-%d")
                    
                    createStock = f"""INSERT INTO `STOCK_LIST` (
                    PartNum, DescStock, QtyStock, PurOrderNum, PurDate, RcvDate, Remark)
                    
                    VALUES (%s, %s, %s, %s, %s, %s, %s)"""
                    
                    inputs = (StockPartNum, StockDesc, StockQty,
                              StockOrderNum, formatDate, None, "")

                    curStock = connStock.cursor()
                    curStock.execute(createStock, inputs)
                    connStock.commit()
                    curStock.close()

                else:
                    messagebox.showerror("Error",
                                         "Please Check Format",
                                         parent=RepWin)
            curData.close()
            checkIssueStat()

            def totalCostCalc(Qty, OneCost):
                QtyNum = float(Qty) if Qty else 0
                OneCostNum = float(OneCost) if OneCost else 0
                total = QtyNum * OneCostNum
                totalDigit = str("{:.2f}".format(total))
                return totalDigit
            
            unitDataLst = [["", "Part No.", "Description", "Quantity", 
                            "Tax", "Unit Cost", "Total Cost"]]
            for i in range(0, len(purOrderUnit)):
                unitDataLst.append([f"{str(i+1)}.", purOrderUnit[i][1], 
                                    f"{purOrderUnit[i][2]}, {purOrderUnit[i][4]}", 
                                    purOrderUnit[i][5], f"{purOrderUnit[i][6]}%", 
                                    str("{:.2f}".format(float(purOrderUnit[i][11]) if purOrderUnit[i][11] else 0)) + " " + str(TransCcyUsed), 
                                    str(totalCostCalc(purOrderUnit[i][5], purOrderUnit[i][11])) + " " + str(TransCcyUsed)])
            
            class PurOrder(FPDF):
                def footer(self):
                    self.set_y(-15) # 15 mm above from bottom
                    self.set_font("Arial", "I", 10) # 10 Font Size
                    self.cell(0, 10, f"Page {self.page_no()} / {{nb}}", align="C")
                
                def companyDetail(self):
                    self.image("MWA.jpg", x=98, y=8, h=30)
                    self.ln(35)
                    
                    self.set_font("Arial", "", 18)
                    self.cell(w=90, h=9, txt = "PURCHASE ORDER", align="L")
                    self.set_font("Arial", "B", 8)
                    self.cell(w=50, h=3, txt = "Purchase Order Date", align="L")
                    self.set_font("Arial", "", 8)
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][1]}", align="L", ln=True)
                    
                    self.cell(w=90, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = f"{purOrderInfo[0][3]}", align="L")
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][2]}", align="L", ln=True)
                    
                    self.cell(w=90, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][3]}", align="L", ln=True)
                    
                    self.cell(w=90, h=3, txt = f"{vendorInfo[0][3]}", align="L")
                    self.set_font("Arial", "B", 8)
                    self.cell(w=50, h=3, txt = "Purchase Order Number", align="L")
                    self.set_font("Arial", "", 8)
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][4]}", align="L", ln=True)
                    
                    self.cell(w=90, h=3, txt = f"{vendorInfo[0][8]}", align="L")
                    self.cell(w=50, h=3, txt = f"{purOrderInfo[0][1]}", align="L")
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][5]}", align="L", ln=True)
                    
                    self.cell(w=90, h=3, txt = f"{vendorInfo[0][9]}", align="L")
                    self.cell(w=50, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = f"{companyInfo[0][6]}", align="L", ln=True)
                    
                    if vendorInfo[0][4] == "Singapore":
                        self.cell(w=90, h=3, txt = f"Singapore {vendorInfo[0][7]}", align="L")
                    else:
                        self.cell(w=90, h=3, txt = f"{vendorInfo[0][7]} {vendorInfo[0][6]}, {vendorInfo[0][5]}, {vendorInfo[0][4]} ", align="L")
                    self.set_font("Arial", "B", 8)
                    self.cell(w=50, h=3, txt = "Payment Terms", align="L")
                    self.set_font("Arial", "", 8)
                    self.cell(w=50, h=3, txt = f"Co. Reg No.: {companyInfo[0][7]}", align="L", ln=True)
                    
                    if vendorInfo[0][14] == "":
                        self.cell(w=90, h=3, txt = f"Attn: {vendorInfo[0][11]}", align="L")
                    else:
                        self.cell(w=90, h=3, txt = f"Attn: {vendorInfo[0][11]} / {vendorInfo[0][14]}", align="L")
                    
                    self.cell(w=50, h=3, txt = f"{purOrderInfo[0][2]}", align="L")
                    self.cell(w=50, h=3, txt = f"Buyer: {companyInfo[0][8]}", align="L", ln=True)
                    
                    if vendorInfo[0][12] == "" and vendorInfo[0][15] == "":
                        self.cell(w=90, h=3, txt = "", align="L")
                    elif vendorInfo[0][12] != "" and vendorInfo[0][15] == "":
                        self.cell(w=90, h=3, txt = f"Tel: {vendorInfo[0][12]}", align="L")
                    else:
                        self.cell(w=90, h=3, txt = f"Tel: {vendorInfo[0][12]} / {vendorInfo[0][15]}", align="L")
                    self.cell(w=50, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = f"Contact Number: {companyInfo[0][9]}", align="L", ln=True)
                    
                    if vendorInfo[0][13] == "" and vendorInfo[0][16] == "":
                        self.cell(w=90, h=3, txt = "", align="L")
                    elif vendorInfo[0][13] != "" and vendorInfo[0][16] == "":
                        self.cell(w=90, h=3, txt = f"Email: {vendorInfo[0][13]}", align="L")
                    else:
                        self.cell(w=90, h=3, txt = f"Email: {vendorInfo[0][13]} / {vendorInfo[0][16]}", align="L")
                    self.cell(w=50, h=3, txt = "", align="L")
                    self.cell(w=50, h=3, txt = f"Email: {companyInfo[0][10]}", align="L", ln=True)
                    
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
                    GSTLst = []
                    for i in range(1, len(lst)):
                        val = float(re.sub("[^0-9^.]", "", lst[i][6]))
                        GSTVal = float(str(lst[i][4]).strip("%"))/100
                        GSTCost = val * GSTVal
                        
                        costLst.append(val)
                        GSTLst.append(GSTCost)
                        
                    subTotal = sum(costLst)
                    GSTTotal = sum(GSTLst)
                    totalSGD = subTotal + GSTTotal
                    
                    subTotalStr = str("{:.2f}".format(subTotal)) + " " + str(TransCcyUsed)
                    GSTStr = str("{:.2f}".format(GSTTotal)) + " " + str(TransCcyUsed)
                    totalStr = str("{:.2f}".format(totalSGD)) + " " + str(TransCcyUsed)
                    
                    self.set_font("Arial", "", 10)
                    self.cell(w=155, h=5, txt="Subtotal", align="R")
                    self.cell(w=30, h=5, txt=f"{subTotalStr}", align="R", ln=True)
                    self.cell(w=155, h=5, txt="Tax", align="R")
                    self.cell(w=30, h=5, txt=f"{GSTStr}", align="R", ln=True)
                    self.cell(w=155, h=5, txt="Total", align="R")
                    self.cell(w=30, h=5, txt=f"{totalStr}", align="R", ln=True)
                    
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

        

        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        #additional module
        
        def FetchAllEmployees(): # Fetch All Employees
            Fetch = mysql.connector.connect(host = logininfo[0],
                                           user = logininfo[1], 
                                           password =logininfo[2],
                                            database= "INDEX_EMP_MASTER")
            
            FetchCursor = Fetch.cursor()
            FetchCursor.execute("""SELECT * FROM EMP_DATA""")
            EmpVal = FetchCursor.fetchall()
            Fetch.commit()
            FetchCursor.close()
            if EmpVal == []:
                EmployeeNameList = ["Please Add Employees"]
                return EmployeeNameList
            else:
                EmployeeNameList = []
                for tup in EmpVal:
                    EmployeeNameList.append(tup[2])
                return EmployeeNameList[1:]
            Fetch.commit()
            FetchCursor.close()
            
            
        def ExportToXero():
            
            tkn = RetrieveToken()
            
            ContactDict = XeroSuppliersGetAll()
            ContactDict.pop(None)
            
            ThemesDict = XeroThemesGetAll()
            
            AttentionList = FetchAllEmployees()
            
            AccountCodeDict = XeroAccountCodeGetAll()
            
            TaxRatesDict = XeroTaxRatesGetAll()
            
            
            XeroExportWin = Toplevel()
            XeroExportWin.geometry('880x500')
            XeroExportWin.title("Export Purchase Order to Xero")
                          
            
            
            TopLabel = Label(XeroExportWin, text = "Confirm Details", font=("Arial", 12))
            TopLabel.grid(row=0, column=0,columnspan = 2, padx = 10, pady= 5, sticky=W)
            
            PONumberFrame = Frame(XeroExportWin)
            PONumberFrame.grid(row=0, column=3, padx = 10, pady= 5, sticky=E,columnspan = 2)
            
            PONumberLabel = Label(PONumberFrame,text = 'Current PO Number :')
            PONumberLabel.grid(row = 0 , column = 0,sticky = E )
            PONumber = Label(PONumberFrame,text = '',borderwidth = 2, relief = "sunken",width = 20,anchor = W)
            PONumber.grid(row = 0 , column = 1,sticky = W)
            
            PONumber['text'] = PurOrderNumRef
            
            def CalendarAppBase(ENTRYwidget):
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    ENTRYwidget.config(state="normal")
                    ENTRYwidget.delete(0, END)
                    ENTRYwidget.insert(0, val)
                    ENTRYwidget.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    ENTRYwidget.config(state="normal")
                    ENTRYwidget.delete(0, END)
                    ENTRYwidget.insert(0,datetime.now().strftime('%Y-%m-%d'))
                    ENTRYwidget.config(state="readonly")
                    calWin.destroy()    
                    
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
            
                buttonEmpty = Button(calWin, text="Reset Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
            
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
                
                
            def DelDateCal():
                CalendarAppBase(ShowDelDateEntry) 
                          
            def FindKey(Dict,Value):
                TupleDict = list(Dict.items())
                for pair in TupleDict:
                    if pair[1] == Value:
                        return pair[0]
              
            def TaxTally(event):
                TaxRatesKeyList = list(TaxRatesDict.keys())
                TaxRatesKey = FindKey(TaxRatesDict,AccountCodeDict[AccountCodeComb.get()][1])
                TaxRateComb.current(TaxRatesKeyList.index(TaxRatesKey))

                
            def StrToDate(StrDate):
                year, month, day = map(int, StrDate.split('-'))
                return datetime(year, month, day,12,30)

            def ExportData():
                curPur = connPur.cursor() 
                curPur.execute(f"SELECT * FROM PUR_ORDER_LIST WHERE PurOrderNum = '{PurOrderNumRef}'")
                purOrderInfo = curPur.fetchall()
                
                curPur.execute(f"SELECT * FROM `{PurOrderNumRef}`")
                purOrderUnit = curPur.fetchall()
                curPur.close()
                line_items=[]
                for i in range(0, len(purOrderUnit)):
                    PartInfo = purOrderUnit[i]
                    line_item = LineItem(description=f"{i+1}. {PartInfo[1]}\n{PartInfo[2]}",
                                         quantity=int(PartInfo[5]),
                                         unit_amount= float(PartInfo[8]),
                                         account_code = AccountCodeDict[AccountCodeComb.get()][0],
                                         tax_type = TaxRatesDict[TaxRateComb.get()],
                                         tax_amount = round(((float(PartInfo[6][:-1])/100)*float(PartInfo[8])),2) if PartInfo[8] else None
                                         )
                    
                    line_items.append(line_item)
                
                PO = PurchaseOrder(line_items = line_items,
                                   purchase_order_number = PurOrderNumRef,
                                   contact = None,
                                   date = StrToDate(ShowDateLabel['text']),
                                   delivery_date = StrToDate(ShowDelDateEntry.get()),
                                   branding_theme_id = ThemesDict[ThemeComb.get()],
                                   reference =ReferenceEntry.get(),
                                   currency_code = CurrencyCode(CurrencyCombo.get()),
                                   currency_rate = 1,
                                   line_amount_types=LineAmountTypes("NoTax" if TaxCombo.get() == 'None' else TaxCombo.get()),
                                   delivery_address = DeliveryAddressText.get("1.0",END),
                                   attention_to = AttentionCombo.get(),
                                   telephone = TelephoneEntry.get(),
                                   delivery_instructions = DeliveryInstText.get("1.0",END)
                                   )
                
                
                POS = PurchaseOrders(purchase_orders = [PO])
                POSDict = serialize(POS)
                POSDict["PurchaseOrders"][0].update({"Contact": {"ContactID": f"{ContactDict[ContactCombo.get()]}"}})
                
                #pp.pprint(POSDict)
                if messagebox.askyesno(title = 'Export Confirmation',message = f"Confirm Upload PO {PurOrderNumRef}",parent = XeroExportWin ):
                    
                    json_response = XeroPOPost(POSDict)
                    print(json_response)
                    messagebox.showinfo(title = 'Response Status',message = f"Status {json_response['PurchaseOrders'][0]['StatusAttributeString']}" ,parent = XeroExportWin)
                    XeroExportWin.destroy() 
                else: 
                    messagebox.showerror(title = 'Export Unsuccessful',message = f"PO {PurOrderNumRef} Not Uploaded" ,parent = XeroExportWin)
                   
                
            ContactLabel = Label(XeroExportWin, text = 'Choose Contact')
            ContactLabel.grid(row=1, column=0, padx = 10, pady= 3, sticky=W)
            
            ContactCombo = AutocompleteCombobox(XeroExportWin, width=20, )
                                        #value=list(ContactDict.keys()), state="readonly")
            ContactCombo.set_completion_list(completion_list = list(ContactDict.keys()))
            ContactCombo.current(0) if ContactCombo['value'] else None
            ContactCombo.grid(row=2, column=0, padx = 10, pady= 3, sticky=W)
        
            DateLabel = Label(XeroExportWin, text = 'Date')
            DateLabel.grid(row=1, column=1, padx = 10, pady= 3, sticky=W)
        
            ShowDateLabel = Label(XeroExportWin,text = datetime.now().strftime('%Y-%m-%d'), 
                                  borderwidth = 2, relief = "sunken",width = 20,anchor = W)
            ShowDateLabel.grid(row=2, column=1, padx = 10, pady= 3, sticky=W)
        
            DeliveryDateLabel = Label(XeroExportWin, text = 'Delivery Date')
            DeliveryDateLabel.grid(row=1, column=2, padx = 10, pady= 3, sticky=W)
            
            DelDateFrame = Frame(XeroExportWin)
            DelDateFrame.grid(row=2, column=2, padx = 10, pady= 3, sticky=EW)
            
            ShowDelDateEntry = Entry(DelDateFrame,width = 16)
            ShowDelDateEntry.insert(END,datetime.now().strftime('%Y-%m-%d'))
            ShowDelDateEntry.grid(row=0, column=0, padx = 2, sticky=W)
            
            DelDateCalendarButton = Button(DelDateFrame, text="Cal", width=4, 
                                             command=DelDateCal)
            
            DelDateCalendarButton.grid(row=0, column=1, sticky=W)
            
            ThemeLabel = Label(XeroExportWin, text = 'Theme')
            ThemeLabel.grid(row=1, column=3, padx = 10, pady= 3, sticky=W)
            
            ThemeComb = ttk.Combobox(XeroExportWin, 
                                        value=list(ThemesDict.keys()), state="readonly")
            
            ThemeComb.current(0) if ThemeComb['value'] else None
            
            ThemeComb.grid(row=2, column=3, padx = 10, pady= 3, sticky=EW)
            
            ReferenceLabel = Label(XeroExportWin, text = 'Reference')
            ReferenceLabel.grid(row=1, column=4, padx = 10, pady= 3, sticky=W)
        
            ReferenceEntry = Entry(XeroExportWin,width = 24)
            ReferenceEntry.grid(row=2, column=4, padx = 10, pady= 3, sticky=EW)    
            
            CurrencyLabel = Label(XeroExportWin, text = 'Currency')
            CurrencyLabel.grid(row=3, column=0, padx = 10, pady= 3, sticky=W)
            
            CurrencyCombo = ttk.Combobox(XeroExportWin, width=20, 
                                        value=CountryRef.getCcyLst(), state="readonly")
            
            CurrencyCombo.current(0) if CurrencyCombo['value'] else None
            
            CurrencyCombo.grid(row=4, column=0, padx = 10, pady= 3, sticky=W)
            
            TaxLabel = Label(XeroExportWin, text = 'Tax')
            TaxLabel.grid(row=3, column=1, padx = 10, pady= 3, sticky=W)
            
            TaxCombo = ttk.Combobox(XeroExportWin, width=20, 
                                        value=['Exclusive','Inclusive','None'], state="readonly")
            
            TaxCombo.current(0) if TaxCombo['value'] else None
            
            TaxCombo.grid(row=4, column=1, padx = 10, pady= 3, sticky=W)
            
            LongFrame = Frame(XeroExportWin)
            LongFrame.grid(row=3, column=2,columnspan = 3,rowspan = 2, sticky=EW)
                        
            AccountCodeLabel = Label(LongFrame, text = 'Account Code')
            AccountCodeLabel.grid(row=0, column=0, padx = 10, pady= 3, sticky=W)
            
            AccountCodeComb = ttk.Combobox(LongFrame,value = list(AccountCodeDict.keys()),width = 35)
            AccountCodeComb.grid(row=1, column=0, padx = 10, pady= 3, sticky=EW)
            AccountCodeComb.current(0)
            AccountCodeComb.bind("<<ComboboxSelected>>", TaxTally)
            
            TaxRateLabel = Label(LongFrame, text = 'Tax Rate')
            TaxRateLabel.grid(row=0, column=1, padx = 10, pady= 3, sticky=W)
            
            TaxRateComb = ttk.Combobox(LongFrame,value = list(TaxRatesDict.keys()),width = 38)
            TaxRateComb.grid(row=1, column=1, padx = 10, pady= 3, sticky=EW)
            TaxTally
        
            
            DeliveryAddressLabel = Label(XeroExportWin,text = 'Delivery Address')
            DeliveryAddressLabel.grid(row=5, column=0, padx = 10, pady= 3, sticky=W)
           
                                
            DeliveryAddressText = Text(XeroExportWin,width = 40,height = 8)
            DeliveryAddressText.grid(row=6, column=0, padx = 10,rowspan = 4,columnspan = 2, pady= 3, sticky=W)
            DeliveryAddressText.insert(END,
                                       """20 Woodlands Link\n#09-08(Design & Assembly Center)\n#07-26(Manufacturing Shop)\nWoodlands Industrial Estate\nSingapore 738733\nCo. Reg No.: 201435019E\nGST Reg No.: 201435019E""")
            AttentionLabel = Label(XeroExportWin, text = 'Attention')
            AttentionLabel.grid(row=5, column=2, padx = 10, pady= 3, sticky=NW)
            
            AttentionCombo = ttk.Combobox(XeroExportWin, width=20, 
                                        value=AttentionList, state="readonly")
            
            AttentionCombo.current(0) if AttentionCombo['value'] else None
            
            AttentionCombo.grid(row=6, column=2, padx = 10, pady= 3, sticky=NW)
            
            TelephoneLabel = Label(XeroExportWin, text = 'Telephone')
            TelephoneLabel.grid(row=7, column=2, padx = 10, pady= 3, sticky=NW)
        
            TelephoneEntry = Entry(XeroExportWin,width = 23)
            TelephoneEntry.grid(row=8, column=2, padx = 10, pady= 3, sticky=NW)
            
            DeliveryInstLabel = Label(XeroExportWin,text = 'Delivery Instruction')
            DeliveryInstLabel.grid(row=5, column=3, padx = 10, pady= 3,columnspan = 3, sticky=W)
            
            DeliveryInstText = Text(XeroExportWin,width = 42,height = 8)
            DeliveryInstText.grid(row=6, column=3, padx = 10,rowspan = 4,columnspan = 3, pady= 3, sticky=W)
            
            
            CommandFrame = LabelFrame(XeroExportWin,text = 'Command')
            CommandFrame.grid(ipady = 3 , ipadx = 5, row=10, column=0, padx = 10,columnspan = 6, pady= 3, sticky=W )
            
            Button(CommandFrame,text = 'Export',command = ExportData).grid(row = 0 ,column = 0, padx = 10, pady = 5)
            Button(CommandFrame,text = 'Clear Entry').grid(row = 0 ,column = 1, padx = 10, pady = 5)
        
        
        
        
        
        
        
        
        
        
    
        ButtonRepFrame = LabelFrame(framePur, text="Command")
        ButtonRepFrame.grid(row=2, column=0, padx=(50,10), pady=5, ipadx=10, ipady=5, sticky=W)
        
        DataRepFrame = LabelFrame(framePur, text="Data")
        DataRepFrame.grid(row=3, column=0, padx=(50,10), pady=(0,5), ipadx=10, ipady=5, sticky=W)
        
        # TaxCurrencyFrame = LabelFrame(framePur, text="Tax & Currency")
        # TaxCurrencyFrame.grid(row=3, column=1, padx=(10,75), pady=(0,5), ipadx=10, ipady=5, sticky=W)
        
        buttonSelectThis = Button(ButtonRepFrame, text="Select Unit (Top)", command=selectUnitPur)
        buttonSelectThis.grid(row=0, column=0, padx=10, pady=5, sticky=W)
        
        buttonDeleteThis = Button(ButtonRepFrame, text="Delete Unit (Below)", command=deleteUnitPur)
        buttonDeleteThis.grid(row=0, column=1, padx=10, pady=5, sticky=W)
        
        buttonGenPur = Button(ButtonRepFrame, text="Generate Purchase Order", command=genPurOrder)
        buttonGenPur.grid(row=0, column=3, padx=10, pady=5, sticky=W)
    
        buttonExportXero = Button(ButtonRepFrame, text="Export to Xero", command=ExportToXero)
        buttonExportXero.grid(row=0, column=4, padx=10, pady=5, sticky=W)
    
        buttonClosePur = Button(ButtonRepFrame, text="Close Tab", command=closeTabPur)
        buttonClosePur.grid(row=0, column=5, padx=10, pady=5, sticky=W)
        
        
        ReqQtyLabel = Label(DataRepFrame, text="Req Qty.")
        ReqQtyBox = Spinbox(DataRepFrame, from_=0, to=100000, width=8)
        UnitCostSelectLabel = Label(DataRepFrame, text="Unit Cost")
        UnitCostSelectBox = Spinbox(DataRepFrame, from_=0, to=100000, increment=0.01, width=8)
        

        
        TaxLabel = Label(DataRepFrame, text="Tax")
        TaxBox = Spinbox(DataRepFrame, from_=0, to=100, width=5, state="readonly")
        TaxPercentageLabel = Label(DataRepFrame, text="%")
        
        CurrencyLabel = Label(DataRepFrame, text="Currency")
        CurrencyBox = ttk.Combobox(DataRepFrame,
                                   value=CountryRef.getCcyLst(), 
                                   width=8, state="readonly")

        CurrencyBox.current(0)

        buttonClearData = Button(DataRepFrame, text="Clear", width=8, command=clearUnitData)
        buttonUpdateData = Button(DataRepFrame, text="Update", width=8, command=updateUnitData)
        buttonRefreshData = Button(DataRepFrame, text="Refresh", width=8, command=refreshUnitData)

        # buttonDefaultCommon = Button(DataRepFrame, text="Default", width=8, command=defaultCommon)
        # buttonUpdateCommon = Button(DataRepFrame, text="Update", width=8, command=updateCommon)
        
        # def CurrencySelect(e):
        #     if CurrencyBox.get() == "Others":
        #         CurrencyBox.config(state="normal")
        #         CurrencyBox.delete(0, END)
        #     else:
        #         CurrencyBox.config(state="readonly")
        
        ReqQtyLabel.grid(row=0, column=0, padx=5, pady=5, sticky=E)
        ReqQtyBox.grid(row=0, column=1, padx=5, pady=5, sticky=W)        
        UnitCostSelectLabel.grid(row=0, column=2, padx=5, pady=5, sticky=E)
        UnitCostSelectBox.grid(row=0, column=3, padx=5, pady=5, sticky=W)

        TaxLabel.grid(row=0, column=4, padx=5, pady=5, sticky=E)
        TaxBox.grid(row=0, column=5, padx=5, pady=5, sticky=E)
        TaxPercentageLabel.grid(row=0, column=6, padx=(0,5), pady=5, sticky=W)
        
        CurrencyLabel.grid(row=0, column=7, padx=5, pady=5, sticky=E)
        CurrencyBox.grid(row=0, column=8, padx=5, pady=5, sticky=E)
        
        buttonClearData.grid(row=0, column=9, padx=(20,5), pady=5, sticky=W)
        buttonUpdateData.grid(row=0, column=10, padx=5, pady=5, sticky=W)
        buttonRefreshData.grid(row=0, column=11, padx=5, pady=5, sticky=W)
        
        # buttonDefaultCommon.grid(row=1, column=4, padx=(20,5), pady=5, sticky=W)
        # buttonUpdateCommon.grid(row=1, column=5, padx=(5,10), pady=5, sticky=W)

        # CurrencyBox.bind("<<ComboboxSelected>>", CurrencySelect)



        def TaxBoxClick(e):
            TaxBox.config(state="normal")
        
        def TaxBoxNormal(e):
            TaxBox.config(state="readonly")
        
        TaxBox.bind("<Double-Button-3>", TaxBoxClick)
        TaxBox.bind("<Leave>", TaxBoxNormal)


    
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
                                     "REQ", "Tax", "Vendor", "UnitCost",
                                     "SGDCost", "TransCost")
        
        # SelectTreeView.column("#0", anchor=CENTER, width=50)
        SelectTreeView.column("#0", width=0 ,stretch=NO)
        SelectTreeView.column("Part", anchor=CENTER, width=135)
        SelectTreeView.column("Description", anchor=W, width=200)
        SelectTreeView.column("Maker", anchor=CENTER, width=140)
        SelectTreeView.column("Spec", anchor=W, width=220)
        SelectTreeView.column("REQ", anchor=CENTER, width=60)
        SelectTreeView.column("Tax", anchor=CENTER, width=60)
        SelectTreeView.column("Vendor", anchor=CENTER, width=120)
        SelectTreeView.column("UnitCost", anchor=CENTER, width=80)
        SelectTreeView.column("SGDCost", anchor=CENTER, width=80)
        SelectTreeView.column("TransCost", anchor=CENTER, width=80)
    
        SelectTreeView.heading("#0", text="Index", anchor=W)
        SelectTreeView.heading("Part", text="Part", anchor=CENTER)
        SelectTreeView.heading("Description", text="Description", anchor=CENTER)
        SelectTreeView.heading("Maker", text="Maker", anchor=CENTER)
        SelectTreeView.heading("Spec", text="Maker Specification", anchor=CENTER)
        SelectTreeView.heading("REQ", text="Qty.", anchor=CENTER)
        SelectTreeView.heading("Tax", text="Tax", anchor=CENTER)
        SelectTreeView.heading("Vendor", text="Vendor", anchor=CENTER)
        SelectTreeView.heading("UnitCost", text="Unit Cost", anchor=CENTER)
        SelectTreeView.heading("SGDCost", text="SGD Cost", anchor=CENTER)
        SelectTreeView.heading("TransCost", text="Trans. Cost", anchor=CENTER)

        queryTreeSelect()

        SelectTreeView.bind('<Double-Button-1>', selectUnitData)
        SelectTreeView.bind('<Button-3>', deselectUnitData)
        ValTreeView.bind('<Button-3>', deselectUnitValue)
        
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
    TransCcyLabel = Label(OrderDataFrame, text="Transaction Currency")
    TransExRateLabel = Label(OrderDataFrame, text="Trans. Exchange Rate")
    ProgressStatLabel = Label(OrderDataFrame, text="Completed")
    ApproveStatLabel = Label(OrderDataFrame, text="Approved")
    IssueStatLabel = Label(OrderDataFrame, text="Issued")
    OrderStatLabel = Label(OrderDataFrame, text="Status")
    TotalSGDLabel = Label(OrderDataFrame, text="Total Cost")



    PurOrderNumBox = Entry(OrderDataFrame, width=20)
    OrderNumGenButton = Button(OrderDataFrame, text="Gen", width=5, command=genOrderNum)
    PaymentTermBox = Entry(OrderDataFrame, width=20, state="readonly")
    PaymentTermButton = Button(OrderDataFrame, text="List", width=5, command=payLstGen)

    OrderDateBox = Entry(OrderDataFrame, width=20, state="readonly")
    OrderDateCalendarButton = Button(OrderDataFrame, text="Cal", width=5, 
                                     command=orderCalPro)
    VendorRemarkBox = Entry(OrderDataFrame, width=20, state="readonly")
    VendorSelectButton = Button(OrderDataFrame, text="List", width=5, command=vendorLstGen)

    TransCcyBox = ttk.Combobox(OrderDataFrame, value=CountryRef.getCcyLst(), 
                               width=10, state="readonly")
    TransExRateBox = Entry(OrderDataFrame, width=12, state="readonly")
    TransSGDLabel = Label(OrderDataFrame, text="SGD")
    
    ProgressStatBox = ttk.Combobox(OrderDataFrame, value=["No", "Yes"],
                                   width=10, state="readonly")
    ApproveStatBox = ttk.Combobox(OrderDataFrame, value=["No", "Yes", "Reject"],
                                  width=10, state="readonly")
    IssueStatBox = ttk.Combobox(OrderDataFrame, value=["No", "Yes"],
                                width=10, state="readonly")
    OrderStatBox = Entry(OrderDataFrame, width=20, state="readonly")
    TotalSGDBox = Entry(OrderDataFrame, width=12, state="readonly")
    TotalSGDValLabel = Label(OrderDataFrame, text="SGD")
    


    def getTransExRate(e):
        Ccy = TransCcyBox.get()
        if Ccy == "SGD":
            TransExRateBox.config(state="normal")
            TransExRateBox.delete(0, END)
            TransExRateBox.insert(0, 1.0)
            TransExRateBox.config(state="readonly")
        elif Ccy in CountryRef.getCcyLst():
            ExRate = CountryRef.getExRate(Ccy)
            TransExRateBox.config(state="normal")
            TransExRateBox.delete(0, END)
            TransExRateBox.insert(0, ExRate)
            TransExRateBox.config(state="readonly")
        else:
            TransExRateBox.config(state="normal")
            TransExRateBox.delete(0, END)
            TransExRateBox.insert(0, 0.0)
            TransExRateBox.config(state="readonly")

    TransCcyBox.bind("<<ComboboxSelected>>", getTransExRate)

    PurOrderNumLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    PaymentTermLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    OrderDateLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    VendorRemarkLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)
    
    TransCcyLabel.grid(row=0, column=3, padx=10, pady=5, sticky=E)
    TransExRateLabel.grid(row=1, column=3, padx=10, pady=5, sticky=E)
    
    ProgressStatLabel.grid(row=0, column=6, padx=10, pady=5, sticky=E)
    ApproveStatLabel.grid(row=1, column=6, padx=10, pady=5, sticky=E)
    IssueStatLabel.grid(row=2, column=6, padx=10, pady=5, sticky=E)
    OrderStatLabel.grid(row=3, column=6, padx=10, pady=5, sticky=E)
    TotalSGDLabel.grid(row=2, column=3, padx=10, pady=5, sticky=E)



    PurOrderNumBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    OrderNumGenButton.grid(row=0, column=2, padx=(0,10), pady=5, sticky=W)
    
    PaymentTermBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    PaymentTermButton.grid(row=1, column=2, padx=(0,10), pady=5, sticky=W)
    
    OrderDateBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    OrderDateCalendarButton.grid(row=2, column=2, padx=(0,10), pady=5, sticky=W)
    
    VendorRemarkBox.grid(row=3, column=1, padx=10, pady=5, sticky=W)
    VendorSelectButton.grid(row=3, column=2, padx=(0,10), pady=5, sticky=W)
    # ABOVE GOT 2 columnspan = 4
    
    TransCcyBox.grid(row=0, column=4, columnspan=2, padx=10, pady=5, sticky=W)
    TransExRateBox.grid(row=1, column=4, padx=10, pady=5, sticky=W)
    TransSGDLabel.grid(row=1, column=5, padx=(0,10), pady=5, sticky=W)
    
    ProgressStatBox.grid(row=0, column=7, padx=10, pady=5, sticky=W)
    ApproveStatBox.grid(row=1, column=7, padx=10, pady=5, sticky=W)
    IssueStatBox.grid(row=2, column=7, padx=10, pady=5, sticky=W)
    OrderStatBox.grid(row=3, column=7, padx=10, pady=5, sticky=W)
    
    TotalSGDBox.grid(row=2, column=4, padx=10, pady=5, sticky=W)
    TotalSGDValLabel.grid(row=2, column=5, padx=(0,10), pady=5, sticky=W)
    
    
    TransCcyBox.current(0)
    TransExRateBox.config(state="normal")
    TransExRateBox.delete(0, END)
    TransExRateBox.insert(0, 1.0)
    TransExRateBox.config(state="readonly")

    ProgressStatBox.current(0)
    ApproveStatBox.current(0)
    IssueStatBox.current(0)
    
    
    
    def ProgressStatSelect(e):
        ProgressVal = ProgressStatBox.current()
        if ProgressVal == 0:
            ApproveStatBox.current(0)
            IssueStatBox.current(0)
        else:
            pass
    
    def ApproveStatSelect(e):
        ProgressVal = ProgressStatBox.current()
        ApproveVal = ApproveStatBox.current()
        if ProgressVal == 0:
            if ApproveVal == 1 or ApproveVal == 2:
                ApproveStatBox.current(0)
                messagebox.showerror("Error",
                                     "You are NOT ALLOWED to change Approval Status when incomplete",
                                     parent=RepWin)
            
            else:
                if ApproveVal == 0:
                    IssueStatBox.current(0)
                else:
                    pass
    
    ProgressStatBox.bind("<<ComboboxSelected>>", ProgressStatSelect)
    ApproveStatBox.bind("<<ComboboxSelected>>", ApproveStatSelect)

    

    
    
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
        comWin.geometry("450x500")
        comWin.columnconfigure(0, weight=1)
        comWin.rowconfigure(0, weight=1)
        
        CompanyDataFrame = LabelFrame(comWin, text="Record")
        CompanyDataFrame.grid(row=0, column=0, padx=20, pady=(20,5), ipadx=5, ipady=5, sticky=W+E)
        # CompanyDataFrame.pack(fill="x", expand="yes", padx=20)
        
        CompanyButtonFrame = LabelFrame(comWin, text="Command")
        CompanyButtonFrame.grid(row=1, column=0, padx=20, pady=(5,20), ipadx=5, ipady=5, sticky=W+E)
        # CompanyButtonFrame.pack(fill="x", expand="yes", padx=20)

        ComNameLabel = Label(CompanyDataFrame, text="Company Name")
        AddressLabel = Label(CompanyDataFrame, text="Address")
        CenterALabel = Label(CompanyDataFrame, text="Center A Info")
        CenterBLabel = Label(CompanyDataFrame, text="Center B Info")
        BuildingLabel = Label(CompanyDataFrame, text="Building")
        PosCodeLabel = Label(CompanyDataFrame, text="Postal Code")
        ComRegNumLabel = Label(CompanyDataFrame, text="Company Reg No.")
        BuyerLabel = Label(CompanyDataFrame, text="Buyer")
        ContactNumLabel = Label(CompanyDataFrame, text="Contact No.")
        EmailLabel = Label(CompanyDataFrame, text="Contact Email")

        ComNameBox = Entry(CompanyDataFrame, width=30)
        AddressBox = Entry(CompanyDataFrame, width=30)
        CenterABox = Entry(CompanyDataFrame, width=30)
        CenterBBox = Entry(CompanyDataFrame, width=30)
        BuildingBox = Entry(CompanyDataFrame, width=30)
        PosCodeBox = Entry(CompanyDataFrame, width=30)
        ComRegNumBox = Entry(CompanyDataFrame, width=30)
        BuyerBox = Entry(CompanyDataFrame, width=30)
        ContactNumBox = Entry(CompanyDataFrame, width=30)
        EmailBox = Entry(CompanyDataFrame, width=30)

        ComNameLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
        AddressLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
        CenterALabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
        CenterBLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)
        BuildingLabel.grid(row=4, column=0, padx=10, pady=5, sticky=E)
        PosCodeLabel.grid(row=5, column=0, padx=10, pady=5, sticky=E)
        ComRegNumLabel.grid(row=6, column=0, padx=10, pady=5, sticky=E)
        BuyerLabel.grid(row=7, column=0, padx=10, pady=5, sticky=E)
        ContactNumLabel.grid(row=8, column=0, padx=10, pady=5, sticky=E)
        EmailLabel.grid(row=9, column=0, padx=10, pady=5, sticky=E)

        ComNameBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
        AddressBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
        CenterABox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
        CenterBBox.grid(row=3, column=1, padx=10, pady=5, sticky=W)
        BuildingBox.grid(row=4, column=1, padx=10, pady=5, sticky=W)
        PosCodeBox.grid(row=5, column=1, padx=10, pady=5, sticky=W)
        ComRegNumBox.grid(row=6, column=1, padx=10, pady=5, sticky=W)
        BuyerBox.grid(row=7, column=1, padx=10, pady=5, sticky=W)
        ContactNumBox.grid(row=8, column=1, padx=10, pady=5, sticky=W)
        EmailBox.grid(row=9, column=1, padx=10, pady=5, sticky=W)
        
        def queryComInfo():
            curCom = connCom.cursor()
            curCom.execute("SELECT * FROM COMPANY_MWA")
            comInfo = curCom.fetchall()
            curCom.close()
            
            if comInfo == []:
                pass
            else:
                ComNameBox.insert(0, comInfo[0][1])
                AddressBox.insert(0, comInfo[0][2])
                CenterABox.insert(0, comInfo[0][3])
                CenterBBox.insert(0, comInfo[0][4])
                BuildingBox.insert(0, comInfo[0][5])
                PosCodeBox.insert(0, comInfo[0][6])
                ComRegNumBox.insert(0, comInfo[0][7])
                BuyerBox.insert(0, comInfo[0][8])
                ContactNumBox.insert(0, comInfo[0][9])
                EmailBox.insert(0, comInfo[0][10])
        
        def updateComInfo():
            if Login.AUTHLVL == 3 or Login.AUTHLVL == 2:
                respCompanyInfo = messagebox.askokcancel("Confirmation",
                                                         "Submit this Info?",
                                                         parent=comWin)
                if respCompanyInfo == True:
                    updateComInfoCom()
                else:
                    pass
            
            else:
                messagebox.showerror("Insufficient Clearance",
                                     "You Don't Have Enough Clearance for This Action",
                                     parent=comWin)
        
        def updateComInfoCom():
            curCom = connCom.cursor()
            curCom.execute("SELECT * FROM `COMPANY_MWA` WHERE oid = 1")
            result = curCom.fetchall()
            if result == []:
                createInfoCom = """ INSERT INTO `COMPANY_MWA`
                (ComName, Address, CenterA, CenterB, Building, 
                 PosCode, ComRegNum, Buyer, ContactNum, Email)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
            
                InfoComValue = (ComNameBox.get(), AddressBox.get(), CenterABox.get(), 
                                CenterBBox.get(), BuildingBox.get(), PosCodeBox.get(), 
                                ComRegNumBox.get(), BuyerBox.get(), ContactNumBox.get(), 
                                EmailBox.get())
                
                curCom.execute(createInfoCom, InfoComValue)
                connCom.commit()
                messagebox.showinfo("Update Successful", 
                                    "You Have Successfuly Updated Company Info", 
                                    parent=comWin)
                comWin.destroy()
            
            else:
                updateInfoCom = """ UPDATE `COMPANY_MWA` SET
                `ComName` = %s,
                `Address` = %s,
                `CenterA` = %s,
                `CenterB` = %s,
                `Building` = %s,
                `PosCode` = %s,
                `ComRegNum` = %s,
                `Buyer` = %s,
                `ContactNum` = %s,
                `Email` = %s
                
                WHERE `oid` = 1"""
                
                UpdateComValue = (ComNameBox.get(), AddressBox.get(), CenterABox.get(), 
                                  CenterBBox.get(), BuildingBox.get(), PosCodeBox.get(), 
                                  ComRegNumBox.get(), BuyerBox.get(), ContactNumBox.get(), 
                                  EmailBox.get())
                
                curCom.execute(updateInfoCom, UpdateComValue)
                connCom.commit()
                messagebox.showinfo("Update Successful", 
                                    "You Have Successfuly Updated Company Info", 
                                    parent=comWin)
                comWin.destroy()
            
        def clearEntryMWA():
            if Login.AUTHLVL == 3 or Login.AUTHLVL == 2:
                ComNameBox.delete(0, END)
                AddressBox.delete(0, END)
                CenterABox.delete(0, END)
                CenterBBox.delete(0, END)
                BuildingBox.delete(0, END)
                PosCodeBox.delete(0, END)
                ComRegNumBox.delete(0, END)
                BuyerBox.delete(0, END)
                ContactNumBox.delete(0, END)
                EmailBox.delete(0, END)
                
            else:
                messagebox.showerror("Insufficient Clearance",
                                     "You Don't Have Enough Clearance for This Action",
                                     parent=comWin)
            



        
        buttonUpdateComInfo = Button(CompanyButtonFrame, 
                                     text="Update Company Information", 
                                     command=updateComInfo)
        buttonUpdateComInfo.grid(row=0, column=0, padx=10, pady=10)
        buttonClearEntryInfo = Button(CompanyButtonFrame, 
                                     text="Clear Entry", 
                                     command=clearEntryMWA)
        buttonClearEntryInfo.grid(row=0, column=1, padx=10, pady=10)

        queryComInfo()
        
        if Login.AUTHLVL == 0 or Login.AUTHLVL == 1:
            ComNameBox.config(state="readonly")
            AddressBox.config(state="readonly")
            CenterABox.config(state="readonly")
            CenterBBox.config(state="readonly")
            BuildingBox.config(state="readonly")
            PosCodeBox.config(state="readonly")
            ComRegNumBox.config(state="readonly")
            BuyerBox.config(state="readonly")
            ContactNumBox.config(state="readonly")
            EmailBox.config(state="readonly")
            # buttonClearEntryInfo.grid_forget()
    


    menuRep = Menu(RepWin)
    RepWin.config(menu=menuRep)
        
    fileMenu = Menu(menuRep, tearoff=0)
    menuRep.add_cascade(label="File", menu=fileMenu)
    fileMenu.add_command(label="Edit Company Info", command=editCompanyInfo)
    fileMenu.add_separator()
    fileMenu.add_command(label="Exit", command=RepWin.destroy)
    
    queryTreeOrder()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


























