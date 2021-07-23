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
import ConnConfig
import CountryRef
from xerogui.config import OverWrite,CLIENT_ID,CLIENT_SECRET,XERO_EMAIL,CHROME_DRIVER_LOCATION

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

import Login
if Login.AUTH:
    import EmployeeTest
    import Vendor
    import RFQ
    import ReportTest
    import StockTest
    import Instruction

# FIRST LAYER framePro, frameBOM etc. (Represent the Tabs)
# SECOND LAYER ProTreeFrame (Treeview), DataFrame, ButtonFrame etc. Under framePro
# NEW LAYER frameBOM (Select BOM)

# connInit NO DATABASE
# connMain Projectmaster
# connEmp Employee

# 0 NOT STARTED
# 1 IN PROGRESS
# 2 COMPLETED

# 0 DESIGNER
# 1 PURCHASER
# 2 MANAGER
# 3 ADMIN

# 0 Inactive
# 1 Active
    
    root = Tk()
    root.iconbitmap("MWA_Icon.ico")
    # root.tk.call("tk", "scaling", 2)
    root.title("Management") 
    root.state("zoomed")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    global tabNote
    tabNote = ttk.Notebook(root)
    tabNote.grid(row=0, column=0, sticky="NSEW") 
     
    framePro = Frame(tabNote)
    tabNote.add(framePro, text="Project Selection")
    
    framePro.columnconfigure(0, weight=1)
    
    # host = '192.168.1.249'
    
    connInit = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password = logininfo[2])
    
# =============================================================================
#     curInit = connInit.cursor()
#     
#     SetupCommand = ["""CREATE SCHEMA IF NOT EXISTS `INDEX_PRO_MASTER` 
#                     DEFAULT CHARACTER SET utf8mb4 
#                     COLLATE utf8mb4_0900_ai_ci""",
#                     
#                     """
#                     CREATE TABLE IF NOT EXISTS `INDEX_PRO_MASTER`.`PROJECT_INFO`
#                     (`oid` INT AUTO_INCREMENT PRIMARY KEY,
#                      `PROJECT_CLS_ID` VARCHAR(10),
#                      `PROJECT_NAME` VARCHAR(50),
#                      `PROJECT_MANAGER` INT,
#                      `PROJECT_SUPPORT` INT,
#                      `START_DATE` DATE,
#                      `DELIVERY_DATE` DATE,
#                      `BASE_PLAN_LOCK` BOOL,
#                      `PROJECT_LOCK` BOOL,
#                      `STATUS` INT,
#                      `NUM_DELETED` INT DEFAULT 0,
#                      `NUM_ONHOLD` INT DEFAULT 0,
#                      `EST_PART` INT DEFAULT 0,
#                      `EST_COST` FLOAT DEFAULT 0.0,
#                      `DATE_OF_ENTRY` DATETIME NOT NULL)
#                     
#                     ENGINE = InnoDB
#                     DEFAULT CHARACTER SET = utf8mb4
#                     COLLATE = utf8mb4_0900_ai_ci"""]
#     
#     
#     
#     for com in SetupCommand:
#         
#         curInit.execute(com)
#         connInit.commit()
#     
#     curInit.close()
# =============================================================================

    
    
    
    connMain = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database= "INDEX_PRO_MASTER")
    
    connEmp = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                      database= "INDEX_EMP_MASTER")
    
    #0 BOM Cost
    #1 Export CSV
    #2 Total Cost
    #3 Delete Project
    #4 Edit Employee
    #5 Approve PurOrder
    #6 Generate PurOrder
    
    def AuthLevel(Auth, index):
        AuthDic = {0:[0,0,0,0,0,0,0], 1:[1,1,0,0,0,0,1],
                   2:[1,1,1,1,0,0,0], 3:[1,1,1,1,1,1,1]}
        AuthBool = AuthDic.get(Auth)[index]
        return AuthBool

        
    

    
    
    
    
    
    
    
    
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background="silver",
                    rowheight=20,
                    fieldbackground="light grey")
    
    style.map("Treeview")
    
    ProTabTitleLabel = Label(framePro, text="Project List", font=("Arial", 12))
    ProTabTitleLabel.grid(row=0, column=0, padx=0, pady=(10,3), ipadx=0, ipady=0, sticky=W+E)
    ProTreeFrame = Frame(framePro)
    ProTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
    # ProTreeFrame.pack()
    
    ProTreeScroll = Scrollbar(ProTreeFrame)
    ProTreeScroll.pack(side=RIGHT, fill=Y)
    
    ProjectTreeView = ttk.Treeview(ProTreeFrame, yscrollcommand=ProTreeScroll.set, selectmode="browse")
    ProTreeScroll.config(command=ProjectTreeView.yview)
    # ProjectTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
    ProjectTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
    
    ProjectTreeView['columns'] = ("ProClassID", 
                                  "ProName", 
                                  "ProManager", 
                                  "ProSupport", 
                                  "StartDate", 
                                  "DeliveryDate", 
                                  "Lock", 
                                  "Status",
                                  "Deleted",
                                  "OnHold",
                                  "EstPart",
                                  "EstCost",
                                  "DateEntry")
    
    ProjectTreeView.column("#0", anchor = CENTER, width =60, minwidth = 0)
    # ProjectTreeView.column("#0",  width=0, stretch=NO)
    ProjectTreeView.column("ProClassID", anchor = CENTER, width = 70, minwidth = 30)
    ProjectTreeView.column("ProName", anchor = W, width= 135, minwidth = 30)
    ProjectTreeView.column("ProManager", anchor = W, width = 95, minwidth = 30)
    ProjectTreeView.column("ProSupport", anchor = W, width = 95, minwidth = 30)
    ProjectTreeView.column("StartDate", anchor = CENTER, width = 100, minwidth = 30)
    ProjectTreeView.column("DeliveryDate", anchor = CENTER, width = 100, minwidth = 30)
    ProjectTreeView.column("Lock", anchor = CENTER, width = 60, minwidth = 30)
    ProjectTreeView.column("Status", anchor = CENTER, width = 80, minwidth = 30)
    ProjectTreeView.column("Deleted", anchor = CENTER, width = 70, minwidth = 30)
    ProjectTreeView.column("OnHold", anchor = CENTER, width = 70, minwidth = 30)
    ProjectTreeView.column("EstPart", anchor = CENTER, width = 70, minwidth = 30)
    ProjectTreeView.column("EstCost", anchor = CENTER, width = 95, minwidth = 30)
    ProjectTreeView.column("DateEntry", anchor = CENTER, width = 90, minwidth = 30)
    
    ProjectTreeView.heading("#0", text = "Index")
    ProjectTreeView.heading("ProClassID", text = "Project ID") #0
    ProjectTreeView.heading("ProName", text = "Project Name") #1
    ProjectTreeView.heading("ProManager", text = "Manager") #2
    ProjectTreeView.heading("ProSupport", text = "Support") #3
    ProjectTreeView.heading("StartDate", text = "Start Date") #4
    ProjectTreeView.heading("DeliveryDate", text = "Delivery Date") #5
    ProjectTreeView.heading("Lock", text = "Lock") #6
    ProjectTreeView.heading("Status", text = "Status") #7
    ProjectTreeView.heading("Deleted", text = "Deleted") #8
    ProjectTreeView.heading("OnHold", text = "On Hold") #9
    ProjectTreeView.heading("EstPart", text = "Est. Part") #10
    ProjectTreeView.heading("EstCost", text = "Est. Cost") #11
    ProjectTreeView.heading("DateEntry", text = "Date of Entry") #12
    
    if AuthLevel(Login.AUTHLVL, 2) == False:
        ProjectTreeView.config(displaycolumns = [0,1,2,3,4,5,6,7,8,9,10,12])
    
    
    
    
    
    
    
    
    

    def deselectProClick(e):
        deselectPro()
    
    def selectProClick(e):
        selectPro()
    
    def updateProReturn(e):
        if buttonUpdatePro["state"] == "disabled":
            messagebox.showwarning("Unable to Update", 
                                   "Please Select a Project", parent=framePro) 
        else:
            updatePro()
    
    def deleteProDel(e):
        if AuthLevel(Login.AUTHLVL, 3) == False:
            messagebox.showerror("Insufficient Clearance",
                                 "You are NOT authorized to delete", parent=framePro)
        
        else:
            if buttonDeletePro["state"] == "disabled":
                messagebox.showwarning("Unable to Delete", 
                                       "Please Select a Project", parent=framePro) 
            else:
                deletePro()
            
    def loadProClick(e):
        if buttonLoadPro["state"] == "normal":
            loadPro()

    def exitAppEsc(e):
        tabNum = tabNote.index("current")
        if tabNum == 0:
            respExitApp = messagebox.askokcancel("Confirmation",
                                                 "Exit Management App Now?",
                                                 parent=framePro)
            if respExitApp == True:
                root.destroy()
            else:
                pass



    ProjectTreeView.bind("<Button-3>", deselectProClick)
    ProjectTreeView.bind("<Double-Button-1>", selectProClick)
    ProjectTreeView.bind("<Return>", updateProReturn)
    ProjectTreeView.bind("<Delete>", deleteProDel)
    ProjectTreeView.bind("<Button-2>", loadProClick)
    ProjectTreeView.bind("<Escape>", exitAppEsc)





    def queryTreePro():
        curMain = connMain.cursor()
        curMain.execute("SELECT * FROM PROJECT_INFO")
        recLst = curMain.fetchall()     

        boolLst = ["N", "Y"]
        statLst = ["Not Started", "In Progress", "Completed"]
        
        checkYearLst = []
        for rec in recLst:
            yearVal = f"Y20{rec[1][1:3]}"
            if yearVal not in checkYearLst:
                checkYearLst.append(yearVal)

        for yr in checkYearLst:
            ProjectTreeView.insert(parent="", index=END, iid=yr, text=yr,
                                    values=("",))
        
        sortYear = sorted(checkYearLst)
        yearNumLst = list(range(len(checkYearLst)))
        dictYearAndNum = dict(zip(sortYear, yearNumLst))
        
        for rec in recLst:
            EstCost = f"{rec[13]} SGD"
            ProjectTreeView.insert(parent=f"Y20{rec[1][1:3]}", index=END, iid=rec[0], text="", 
                                    values=(rec[1], rec[2], EmpOidToName(rec[3]), 
                                            EmpOidToName(rec[4]), rec[5], rec[6], 
                                            f"{boolLst[rec[7]]} / {boolLst[rec[8]]}", 
                                            statLst[rec[9]], rec[10], rec[11], rec[12],
                                            EstCost, rec[14]))
        
        yearVal = ProjectTreeView.get_children()
        for yr in yearVal:
            iidLst = ProjectTreeView.get_children(yr)
            YrNum = dictYearAndNum.get(yr)
            ProjectTreeView.move(yr, "", YrNum)
            
            oidLst = list(iidLst)
            indexLst = []
            for rec in recLst:
                if str(rec[0]) in iidLst:
                    indexVal = int(rec[1][1:])
                    indexLst.append(indexVal)
            
            numLst = list(range(len(indexLst)))

            sortIndex = sorted(indexLst)
            dictIndexAndOid = dict(zip(indexLst, oidLst))
            dictNumAndIndex = dict(zip(numLst, sortIndex))
            
            for i in numLst:
                IndexVal = dictNumAndIndex.get(i)
                oidVal = dictIndexAndOid.get(IndexVal)             
                ProjectTreeView.move(oidVal, yr, i)
                
        connMain.commit()
        curMain.close()
        
        
        
    def checkProID():
        curMain = connMain.cursor()
        curMain.execute("SELECT * FROM PROJECT_INFO")
        recLst = curMain.fetchall()  
        connMain.commit()
        curMain.close()
        ProIDLst = []
        for ele in recLst:
            ProIDLst.append(ele[1])
        return ProIDLst
    
    def checkEmpPro():
        # if FetchEmployees()[0] == "Please Add Employees":
        if ProManagerBox.get() == "Please Add Employees" or ProSupportBox.get() == "Please Add Employees":
            buttonCreatePro.config(state=DISABLED)
        else:
            buttonCreatePro.config(state=NORMAL)
    
    def FetchEmployees(): # Fetch Active Employees
        Fetch = mysql.connector.connect(host = logininfo[0],
                                        user = logininfo[1], 
                                        password =logininfo[2],
                                        database= "INDEX_EMP_MASTER")
        FetchCursor = Fetch.cursor()
        FetchCursor.execute("""SELECT * FROM `EMP_DATA` WHERE STATUS = 1""")
        EmpVal = FetchCursor.fetchall()[1:]
        
        if EmpVal == []:
            EmployeeNameList = ["Please Add Employees"]
            return EmployeeNameList
        else:
            EmployeeNameList = []
            for tup in EmpVal:
                EmployeeNameList.append(tup[2])
            
            return EmployeeNameList
        Fetch.commit()
        FetchCursor.close()
        
    def FetchAllEmployees(): # Fetch All Employees
        Fetch = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                        database= "INDEX_EMP_MASTER")
        
        FetchCursor = Fetch.cursor()[1:]
        FetchCursor.execute("""SELECT * FROM EMP_DATA""")
        EmpVal = FetchCursor.fetchall()
        connEmp.commit()
        curEmp.close()
        if EmpVal == []:
            EmployeeNameList = ["Please Add Employees"]
            return EmployeeNameList
        else:
            EmployeeNameList = []
            for tup in EmpVal:
                EmployeeNameList.append(tup[2])
            return EmployeeNameList
        Fetch.commit()
        FetchCursor.close()
        
    def EmpNameToOid(name):
        curEmp = connEmp.cursor()
        if name not in FetchEmployees():
            newName = name[:-11]
            selectSqlCom = """SELECT * FROM EMP_DATA WHERE EMPLOYEE_NAME = %s"""
            val = (newName, )
            curEmp.execute(selectSqlCom, val)
            EmpVal = curEmp.fetchall()
            return EmpVal[0][0]
        else:
            newName = name
            selectSqlCom = """SELECT * FROM EMP_DATA WHERE EMPLOYEE_NAME = %s"""
            val = (newName, )
            curEmp.execute(selectSqlCom, val)
            EmpVal = curEmp.fetchall()
            return EmpVal[0][0]
    
        connEmp.commit()
        curEmp.close()
        
    def EmpOidToName(oid):    
        curEmp = connEmp.cursor()
        selectSqlCom = """SELECT * FROM EMP_DATA WHERE oid = %s"""
        val = (oid, )
        curEmp.execute(selectSqlCom, val)
        EmpVal = curEmp.fetchall()
        connEmp.commit()
        curEmp.close()
        return EmpVal[0][2]
    
    def EmpOidToNameInactive(oid): # Append INACTIVE if inactive
        selectSqlCom = """SELECT * FROM EMP_DATA WHERE oid = %s"""
        val = (oid, )
        curEmp = connEmp.cursor()
        curEmp.execute(selectSqlCom, val)
        EmpVal = curEmp.fetchall()
        connEmp.commit()
        curEmp.close()
        empName = EmpVal[0][2]
        if empName not in FetchEmployees():
            newEmpName = f"{empName} (Inactive)"
            return newEmpName
        else:
            return empName
    
    def updatePro():
        respUpdatePro = messagebox.askokcancel("Confirmation",
                                               "Update This Project?", parent=framePro)
        if respUpdatePro == True:
            sqlCommand = f"""UPDATE PROJECT_INFO SET
            `PROJECT_NAME` = %s,
            `PROJECT_MANAGER` = %s,
            `PROJECT_SUPPORT` = %s,
            `START_DATE` = %s,
            `DELIVERY_DATE` = %s,
            `BASE_PLAN_LOCK` = %s,
            `PROJECT_LOCK` = %s,
            `STATUS` = %s
            
            WHERE `oid` = %s"""
            
            ProID = ProIDBox.get()
            ProName = ProNameBox.get().upper()
            Stat = ProStatusBox.current()
            
            def checkDatePro(dateVar):
                if dateVar.get() == "":
                    return None
                else:
                    return dateVar.get()
            
            ProMan = EmpNameToOid(ProManagerBox.get())
            ProSupp = EmpNameToOid(ProSupportBox.get())
            
            BasePlanLock = ProBasePlanLockVal.get()
            ProLock = ProLockVal.get()
            
            selected = ProjectTreeView.selection()[0]
            
            inputs = (ProName, ProMan, ProSupp, checkDatePro(ProStartDateBox), 
                      checkDatePro(ProDeliveryDateBox), BasePlanLock, ProLock, Stat, selected)
            
            curMain = connMain.cursor()
            curMain.execute(sqlCommand, inputs)
            connMain.commit()
            curMain.close()
            
            clearEntryPro()
            ProjectTreeView.delete(*ProjectTreeView.get_children())
            queryTreePro()
            
            messagebox.showinfo("Update Successful", 
                                f"You Have Updated Project {ProID}", parent=framePro) 
        else:
            pass
    
    def createProCom():
        timeNow = datetime.now()
        # formatDate = timeNow.strftime("%Y-%m-%d %H:%M:%s")
        formatDate = timeNow.strftime("%Y-%m-%d")
        
        createCom = """INSERT INTO PROJECT_INFO (
        PROJECT_CLS_ID, PROJECT_NAME, PROJECT_MANAGER, PROJECT_SUPPORT, 
        START_DATE, DELIVERY_DATE, BASE_PLAN_LOCK, PROJECT_LOCK, 
        STATUS, DATE_OF_ENTRY)
    
        VALUES (%s, %s ,%s, %s, %s, %s, %s, %s, %s, %s)"""
        
        def checkDatePro(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        values = (ProIDBox.get(), ProNameBox.get().upper(), EmpNameToOid(ProManagerBox.get()), 
                  EmpNameToOid(ProSupportBox.get()), checkDatePro(ProStartDateBox), 
                  checkDatePro(ProDeliveryDateBox), ProBasePlanLockVal.get(), 
                  ProLockVal.get(), ProStatusBox.current(), formatDate)
        curMain = connMain.cursor()
        curMain.execute(createCom, values)
        connMain.commit()
        curMain.close()
        
        ProID = ProIDBox.get()
        
        curInit = connInit.cursor()
        curInit.execute(f""" CREATE SCHEMA IF NOT EXISTS `{ProID}`
                        DEFAULT CHARACTER SET utf8mb4""")
        connInit.commit()
        curInit.close()
        
        connPro = mysql.connector.connect(host = logininfo[0],
                                          user = logininfo[1], 
                                          password =logininfo[2],
                                          database= f"{ProID}")

        curPro = connPro.cursor()

        newProComLst = [""" CREATE TABLE IF NOT EXISTS `MACH_INDEX`(
            `oid` TINYINT AUTO_INCREMENT PRIMARY KEY,
            `MACH_ID` VARCHAR(5),
            `MACH_NAME` VARCHAR(50),
            `LEAD_DESIGNER` INT,
            `LEAD_LOCKER` INT,
            `ORDER_QTY` INT DEFAULT 0,
            `COMP_QTY` INT DEFAULT 0,
            `MACH_DES_DUE` DATE,
            `MACH_DES_COMP` DATE,
            `STATUS` INT,
            `NUM_UA` INT DEFAULT 0,
            `NUM_BOM_LOAD` INT DEFAULT 0,
            `NUM_DELETED` INT DEFAULT 0,
            `NUM_ONHOLD` INT DEFAULT 0,
            `EST_PART_MACH` INT DEFAULT 0,
            `EST_COST_MACH` FLOAT DEFAULT 0.00)
    
            ENGINE = InnoDB
            DEFAULT CHARACTER SET = utf8mb4
            COLLATE = utf8mb4_0900_ai_ci"""]
            
        for com in newProComLst:
            curPro.execute(com)
            connPro.commit()
    
        curPro.close()
        
        clearEntryPro()
        ProjectTreeView.delete(*ProjectTreeView.get_children())
        queryTreePro()
        
        messagebox.showinfo("Create Successful", 
                            f"You Have Created Project {ProID}", parent=framePro) 
        
    
    def createPro():
        if ProIDBox.get() in checkProID():
            messagebox.showerror("Create Error", 
                                 f"You Cannot Have Duplicate Project ID {ProIDBox.get()}", parent=framePro) 
        elif ProIDBox.get() == "":
            messagebox.showerror("Create Error", 
                                 "Invalid Project ID", parent=framePro) 
        else:
            if ProNameBox.get() == "":
                resDesc = messagebox.askokcancel("Yet to Enter a Project Name", 
                                                 "Create a Project Without Name?", parent=framePro)
                if resDesc == True:
                    resConA = messagebox.askokcancel("Confirmation", 
                                                     "Create This Project?", parent=framePro)
                    if resConA == True:
                        createProCom()
            else:
                resConB = messagebox.askokcancel("Confirmation", 
                                                 "Create This Project?", parent=framePro)
                if resConB == True:
                    createProCom()                
                
    def deletePro():
        if AuthLevel(Login.AUTHLVL, 3) == False:
            messagebox.showerror("Insufficient Clearance",
                                 "You are NOT authorized to delete", parent=framePro)
        
        else:
            respDelPro = messagebox.askokcancel("Confirmation",
                                                "Delete The Project?", parent=framePro)
            if respDelPro == True:
                curMain = connMain.cursor()
                
                selected = ProjectTreeView.selection()[0]
                if selected[0] == "Y":
                    selectAll = ProjectTreeView.get_children(selected)
                    
                    respDelAllPro = messagebox.askokcancel("Confirmation",
                                                           "Delete ALL Project Under This Year?",
                                                           parent=framePro)
                    if respDelAllPro == True:
                        for index in selectAll:
                            sqlSelect = "SELECT * FROM PROJECT_INFO WHERE oid = %s"
                            valSelect = (index, )
            
                            curMain.execute(sqlSelect, valSelect)
                            
                            recLst = curMain.fetchall()
                            connMain.commit()
                        
                            proID = recLst[0][1]
                                
                            sqlDelete = "DELETE FROM PROJECT_INFO WHERE oid = %s"
                            valDelete = (index, )
                            curMain.execute(sqlDelete, valDelete)
                            connMain.commit()
                            
                            curInit = connInit.cursor()
                            curInit.execute(f"DROP DATABASE IF EXISTS `{proID}`")
                            connInit.commit()
                            curInit.close()
                            
                            clearEntryPro()
                            ProjectTreeView.delete(*ProjectTreeView.get_children())
                            queryTreePro()
                            
                        messagebox.showinfo("Delete Successful", 
                                            f"You Have Deleted {len(selectAll)} Projects", parent=framePro) 
                    else:
                        pass
        
                else:
                    sqlSelect = "SELECT * FROM PROJECT_INFO WHERE oid = %s"
                    valSelect = (selected, )
                    curMain.execute(sqlSelect, valSelect)
                    
                    recLst = curMain.fetchall()
                    connMain.commit()
        
                    proID = recLst[0][1]
                        
                    sqlDelete = "DELETE FROM PROJECT_INFO WHERE oid = %s"
                    valDelete = (selected, )
                    curMain.execute(sqlDelete, valDelete)
                    connMain.commit()
                    
                    curInit = connInit.cursor()
                    curInit.execute(f"DROP DATABASE IF EXISTS `{proID}`")
                    connInit.commit()
                    curInit.close()
                    
                    clearEntryPro()
                    ProjectTreeView.delete(*ProjectTreeView.get_children())
                    queryTreePro()
                    
                    messagebox.showinfo("Delete Successful", 
                                        f"You Have Deleted Project {proID}", parent=framePro) 
                curMain.close()
            else:
                pass
        
    def selectPro():
        selected = ProjectTreeView.selection()
        if selected == ():
            messagebox.showerror("No Item Selected", "Please Select an Item",
                                 parent=framePro)
        else:
            selected = selected[0]

            if selected[0] == "Y":
                respYear = messagebox.askokcancel("You Have Selected a Year", 
                                                  f"This will select EVERYTHING under {selected}", 
                                                  parent=framePro)
                if respYear == True:
                    selectAll = ProjectTreeView.get_children(selected)
                    messagebox.showinfo("Selected", f"{len(selectAll)} Items under {selected} selected",
                                        parent=framePro)
                    
                    buttonUpdatePro.config(state=DISABLED)
                    buttonDeletePro.config(state=NORMAL)
                    buttonCreatePro.config(state=DISABLED)
                    ProIDBox.config(state="readonly")
                    ProIDGenButton.config(state=DISABLED)
                    ProNameBox.delete(0, END)
                    ProNameBox.insert(0, f"{len(selectAll)} Items in {selected}")
                    ProjectTreeView.config(selectmode="none")
                else:
                    pass
    
            else:
                sqlSelect = "SELECT * FROM PROJECT_INFO WHERE oid = %s"
                valSelect = (selected, )
                curMain = connMain.cursor()
                curMain.execute(sqlSelect, valSelect)
                recLst = curMain.fetchall()
                connMain.commit()
                curMain.close()
                
                clearEntryPro()
                
                ProIDBox.insert(0, recLst[0][1])
                ProNameBox.insert(0, recLst[0][2])
                
                ProStatusBox.current(recLst[0][9])
                
                if recLst[0][5] == None:
                    ProStartDateBox.config(state="normal")
                    ProStartDateBox.insert(0, "")
                    ProStartDateBox.config(state="readonly")
                else:
                    ProStartDateBox.config(state="normal")
                    ProStartDateBox.insert(0, recLst[0][5])
                    ProStartDateBox.config(state="readonly")
                if recLst[0][6] == None:
                    ProDeliveryDateBox.config(state="normal")
                    ProDeliveryDateBox.insert(0, "")
                    ProDeliveryDateBox.config(state="readonly")
                else:
                    ProDeliveryDateBox.config(state="normal")
                    ProDeliveryDateBox.insert(0, recLst[0][6])
                    ProDeliveryDateBox.config(state="readonly")      
                
                ProManagerBox.config(state="normal")
                ProManagerBox.delete(0, END)
                ProManagerBox.insert(0, EmpOidToNameInactive(recLst[0][3]))
                ProManagerBox.config(state="readonly") 
                
                ProSupportBox.config(state="normal")
                ProSupportBox.delete(0, END)
                ProSupportBox.insert(0, EmpOidToNameInactive(recLst[0][4]))
                ProSupportBox.config(state="readonly") 
                
                ProBasePlanLockVal.set(recLst[0][7])
                ProLockVal.set(recLst[0][8])
                
                ProDeletedBox.config(state="normal")
                ProDeletedBox.insert(0, recLst[0][10])
                ProDeletedBox.config(state="readonly")
                ProOnHoldBox.config(state="normal")
                ProOnHoldBox.insert(0, recLst[0][11])
                ProOnHoldBox.config(state="readonly")
                
                ProEstPartBox.config(state="normal")
                ProEstPartBox.insert(0, recLst[0][12])
                ProEstPartBox.config(state="readonly")
                
                if AuthLevel(Login.AUTHLVL, 2) == False:
                    pass
                else:
                    ProEstCostBox.config(state="normal")
                    ProEstCostBox.insert(0, recLst[0][13])
                    ProEstCostBox.config(state="readonly")
                
                buttonUpdatePro.config(state=NORMAL)
                buttonDeletePro.config(state=NORMAL)
                buttonCreatePro.config(state=DISABLED)
                ProIDBox.config(state="readonly")
                ProIDGenButton.config(state=DISABLED)
                ProjectTreeView.config(selectmode="none")
        
    
    def deselectPro():
        selected = ProjectTreeView.selection()
        if len(selected) > 0:
            ProjectTreeView.selection_remove(selected[0])
            clearEntryPro()
        else:
            clearEntryPro()
    
    def clearEntryPro():
        buttonUpdatePro.config(state=DISABLED)
        buttonDeletePro.config(state=DISABLED)
        checkEmpPro()
        ProIDBox.config(state="normal")
        ProIDGenButton.config(state=NORMAL)
        ProjectTreeView.config(selectmode="browse")
        
        ProIDBox.delete(0, END)
        ProNameBox.delete(0, END)
        ProStatusBox.current(0)
        
        ProStartDateBox.config(state="normal")
        ProStartDateBox.delete(0, END)
        ProStartDateBox.config(state="readonly")
        ProDeliveryDateBox.config(state="normal")
        ProDeliveryDateBox.delete(0, END)
        ProDeliveryDateBox.config(state="readonly")
        
        ProManagerBox.current(0)
        ProSupportBox.current(0)
        
        ProBasePlanLockVal.set(0)
        ProLockVal.set(0)
        
        ProDeletedBox.config(state="normal")
        ProDeletedBox.delete(0, END)
        ProDeletedBox.config(state="readonly")
        ProOnHoldBox.config(state="normal")
        ProOnHoldBox.delete(0, END)
        ProOnHoldBox.config(state="readonly")
        
        ProEstPartBox.config(state="normal")
        ProEstPartBox.delete(0, END)
        ProEstPartBox.config(state="readonly")
        ProEstCostBox.config(state="normal")
        ProEstCostBox.delete(0, END)
        ProEstCostBox.config(state="readonly")
    
    def refreshPro():
        ProManagerBox.config(value=FetchEmployees())
        ProSupportBox.config(value=FetchEmployees())
        clearEntryPro()
        ProjectTreeView.delete(*ProjectTreeView.get_children())
        queryTreePro()
        checkEmpPro()
    
    def genProID():
        curMain = connMain.cursor()
        curMain.execute("""SELECT MAX(oid) FROM PROJECT_INFO """)
        maxOID = curMain.fetchall()[0][0]
        
        if maxOID == None:
            nextInt = 1
        
        else:
            curMain.execute(f"""SELECT * FROM PROJECT_INFO WHERE oid = {maxOID} """)
            currentNum = curMain.fetchall()[0][1]        
            currentInt = int(str(currentNum)[-3:])
            nextInt = currentInt + 1
        
        connMain.commit()
        curMain.close()
        
        Currentyear = str((datetime.now().year)%100)
        threeDigit = str(nextInt).rjust(3,'0')
        newProID = f"P{Currentyear}{threeDigit}"
        
        ProIDBox.delete(0, END)
        ProIDBox.insert(0, newProID)
    
    def startCalPro():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("270x260")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            ProStartDateBox.config(state="normal")
            ProStartDateBox.delete(0, END)
            ProStartDateBox.insert(0, val)
            ProStartDateBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            ProStartDateBox.config(state="normal")
            ProStartDateBox.delete(0, END)
            ProStartDateBox.config(state="readonly")
            calWin.destroy()
        
        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
    def deCalPro():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("270x260")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            ProDeliveryDateBox.config(state="normal")
            ProDeliveryDateBox.delete(0, END)
            ProDeliveryDateBox.insert(0, val)
            ProDeliveryDateBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            ProDeliveryDateBox.config(state="normal")
            ProDeliveryDateBox.delete(0, END)
            ProDeliveryDateBox.config(state="readonly")
            calWin.destroy()
        
        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
    
    
    
    
    def loadPro():
        selectVal = ProjectTreeView.selection()
        if selectVal == ():
            messagebox.showwarning("Unable to Load", "Please Select a Project",
                                   parent=framePro)
        else:
            if selectVal[0][0] == "Y":
                messagebox.showwarning("Unable to Load", "You Have Selected a Year",
                                        parent=framePro)
            else:
                loadProCom()

    def loadProCom():
        buttonLoadPro.config(state=DISABLED)

        curMain = connMain.cursor()
        loadProSelect = ProjectTreeView.selection()[0]
        
        sqlSelect = "SELECT * FROM PROJECT_INFO WHERE oid = %s"
        valSelect = (loadProSelect, )
        curMain = connMain.cursor()
        curMain.execute(sqlSelect, valSelect)
        resultSelect = curMain.fetchall()
        connMain.commit()
        curMain.close() 
        
        proName = resultSelect[0][1]
        proDesc = resultSelect[0][2]
        
        connLoad = mysql.connector.connect(host = logininfo[0],
                                           user = logininfo[1], 
                                           password =logininfo[2],
                                           database=proName)
        
        
        
        frameMach = Frame(tabNote)
        tabNote.add(frameMach, text="Machine Selection")
        
        frameMach.columnconfigure(0, weight=1)
        
        tabObjDict = tabNote.children
        tabNote.select(len(tabObjDict)-1)
    
    
    
    
       
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview",
                        background="silver",
                        rowheight=20,
                        fieldbackground="light grey")
        
        style.map("Treeview")
        
        # MachTabTitleLabel = Label(frameMach, 
        #                           text=f"Machine List for Project {proName} {proDesc}", 
        #                           font=("Arial", 12))
        
        MachTabTitleLabel = Label(frameMach, 
                                  text=f"Machine List for Project {proName}", 
                                  font=("Arial", 12))
        MachTabTitleLabel.grid(row=0, column=0, padx=0, pady=(10,3), ipadx=0, ipady=0, sticky=W+E) 
        MachTreeFrame = Frame(frameMach)
        MachTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
        # MachTreeFrame.pack()
        
        MachTreeScroll = Scrollbar(MachTreeFrame)
        MachTreeScroll.pack(side=RIGHT, fill=Y)
        
        MachTreeView = ttk.Treeview(MachTreeFrame, yscrollcommand=MachTreeScroll.set, 
                                    selectmode="browse")
        MachTreeScroll.config(command=MachTreeView.yview)
        # MachTreeView.grid(row=2, column=0, columnspan=10, padx=10)
        MachTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
        
        MachTreeView['columns'] = ("MachID","MachName", 
                                   "LeadDesigner",
                                   "OrderQty", "CompQty",
                                   "DesDue", "DesComp",
                                   "NumUA", "NumBOMLoad", "Status",
                                   "Deleted", "OnHold",
                                   "EstPartMach", "EstCostMach")
        
        # MachTreeView.column("#0", anchor = CENTER, width=45, minwidth=0)
        MachTreeView.column("#0",  width=0, stretch=NO)
        MachTreeView.column("MachID",anchor = CENTER, width = 50, minwidth = 30)
        MachTreeView.column("MachName",anchor = W, width = 130, minwidth = 50)
        MachTreeView.column("LeadDesigner",anchor = W, width = 130, minwidth = 50)
        MachTreeView.column("OrderQty",anchor = CENTER, width = 80, minwidth = 50)
        MachTreeView.column("CompQty",anchor = CENTER, width = 80, minwidth = 50)
        MachTreeView.column("DesDue",anchor = CENTER, width = 90, minwidth = 50)
        MachTreeView.column("DesComp",anchor = CENTER, width = 90, minwidth = 50)
        MachTreeView.column("NumUA",anchor = CENTER, width = 55, minwidth = 50)
        MachTreeView.column("NumBOMLoad",anchor = CENTER, width = 75, minwidth = 50)
        MachTreeView.column("Status",anchor = CENTER,width = 100, minwidth = 50)
        MachTreeView.column("Deleted",anchor = CENTER,width = 70, minwidth = 50)
        MachTreeView.column("OnHold",anchor = CENTER,width = 70, minwidth = 50)
        MachTreeView.column("EstPartMach",anchor = CENTER,width = 70, minwidth = 50)
        MachTreeView.column("EstCostMach",anchor = CENTER,width = 80, minwidth = 50)
        
        MachTreeView.heading("#0",text = "Index")
        MachTreeView.heading("MachID",text = "ID") #0
        MachTreeView.heading("MachName",text = "Machine Name") #1
        MachTreeView.heading("LeadDesigner",text = "Designer") #2
        MachTreeView.heading("OrderQty",text = "Order Qty.") #3
        MachTreeView.heading("CompQty",text = "Comp. Qty.") #4
        MachTreeView.heading("DesDue",text = "Design Due") #5
        MachTreeView.heading("DesComp",text = "Design Comp") #6
        MachTreeView.heading("NumUA",text = "No. UA") #7
        MachTreeView.heading("NumBOMLoad",text = "BOM Load") #8
        MachTreeView.heading("Status",text = "Status") #9
        MachTreeView.heading("Deleted",text = "Deleted") #10
        MachTreeView.heading("OnHold",text = "On Hold") #11
        MachTreeView.heading("EstPartMach",text = "Est. Part") #12
        MachTreeView.heading("EstCostMach",text = "Est. Cost") #13
        
        if AuthLevel(Login.AUTHLVL, 2) == False:
            MachTreeView.config(displaycolumns = [0,1,2,3,4,5,6,7,8,9,10,11,12])
    
        MachDataFrame = LabelFrame(frameMach, text="Record")
        MachDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
        # MachDataFrame.pack(fill="x", expand="yes", padx=20)
        
        MachButtonFrame = LabelFrame(frameMach, text="Command")
        MachButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
        # MachDataFrame.pack(fill="x", expand="yes", padx=20)





        def deselectMachClick(e):
            deselectMach()
        
        def selectMachClick(e):
            selectMach()
        
        def updateMachReturn(e):
            if buttonUpdateMach["state"] == "disabled":
                messagebox.showwarning("Unable to Update", 
                                       "Please Select a Machine", parent=frameMach) 
            else:
                updateMach()
        
        def deleteMachDel(e):
            if buttonDeleteMach["state"] == "disabled":
                messagebox.showwarning("Unable to Delete", 
                                       "Please Select a Machine", parent=frameMach) 
            else:
                deleteMach()

        def loadMachClick(e):
            if buttonLoadMach["state"] == "normal":
                loadMach()

        def closeMachTabEsc(e):
            tabNum = tabNote.index("current")
            if tabNum == 1:
                closeTabMach()

        MachTreeView.bind("<Button-3>", deselectMachClick)
        MachTreeView.bind("<Double-Button-1>", selectMachClick)
        MachTreeView.bind("<Return>", updateMachReturn)
        MachTreeView.bind("<Delete>", deleteMachDel)
        MachTreeView.bind("<Button-2>", loadMachClick)
        MachTreeView.bind("<Escape>", closeMachTabEsc) 





        def queryTreeMach():
            curLoad = connLoad.cursor()
            curLoad.execute("SELECT * FROM MACH_INDEX")
            recLst = curLoad.fetchall()
            connLoad.commit()
            curLoad.close()

            statLst = ["Not Started", "In Progress", "Completed"]
            
            def showDelOnHoldMach(val):
                if int(val) == 0:
                    return "N"
                else:
                    return val
            
            for rec in recLst:
                estMachCost = f"{rec[15]} SGD"
                MachTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                    values=(rec[1], rec[2], 
                                            EmpOidToName(rec[3]), 
                                            rec[5], rec[6], rec[7], rec[8],
                                            rec[10], rec[11], 
                                            statLst[rec[9]], 
                                            showDelOnHoldMach(rec[12]), 
                                            showDelOnHoldMach(rec[13]),
                                            rec[14], estMachCost))

        def checkProTotal():
            curLoad = connLoad.cursor()
            curLoad.execute(f"SELECT * FROM `MACH_INDEX`")
            recLst = curLoad.fetchall()
            
            DeleteProLst = []
            OnHoldProLst = []
            EstPartProLst = []
            EstCostProLst = []
            
            for mach in recLst:
                DeleteProLst.append(mach[12])
                OnHoldProLst.append(mach[13])
                EstPartProLst.append(mach[14])
                EstCostProLst.append(mach[15])
            
            DeleteProTotal = sum(DeleteProLst)
            OnHoldProTotal = sum(OnHoldProLst)
            EstPartProTotal = sum(EstPartProLst)
            EstCostProTotal = sum(EstCostProLst)
            
            curLoad.close()
            curMain = connMain.cursor()
            
            sqlProNum = f"""UPDATE `PROJECT_INFO` SET
            `NUM_DELETED` = %s,
            `NUM_ONHOLD` = %s,
            `EST_PART` = %s,
            `EST_COST` = %s
            
            WHERE `oid` = %s"""  
            
            inputs = (DeleteProTotal, OnHoldProTotal, EstPartProTotal, 
                      EstCostProTotal, loadProSelect)
            
            curMain.execute(sqlProNum, inputs)
            connMain.commit()
            curMain.close()
            
            clearEntryPro()
            ProjectTreeView.delete(*ProjectTreeView.get_children())
            queryTreePro()
    
        def genMachID():
            curLoad = connLoad.cursor()
            curLoad.execute("""SELECT MAX(oid) FROM MACH_INDEX """)
            maxOID = curLoad.fetchall()[0][0]
    
            
            if maxOID == None:
                nextInt = 1
            
            else:
                curLoad.execute(f"""SELECT * FROM MACH_INDEX WHERE oid = {maxOID} """)
                currentNum = curLoad.fetchall()[0][1]        
                nextInt = int(currentNum) + 1
            
            twoDigitMach = str(nextInt).rjust(2,'0')
            newMachID = f"{twoDigitMach}"
            
            connLoad.commit()
            curLoad.close()
            
            MachIDBox.delete(0, END)
            MachIDBox.insert(0, newMachID)
        
        def checkEmpMach():
            if MachDesignerBox.get() == "Please Add Employees" or MachLockerBox.get() == "Please Add Employees":
                buttonCreateMach.config(state=DISABLED)
            else:
                buttonCreateMach.config(state=NORMAL)
        
        def dueCalMach():
            calWin = Toplevel()
            calWin.title("Select the Date")
            calWin.geometry("270x260")
            
            cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
            cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
            
            def confirmDate():
                val = cal.get_date()
                MachDueDateBox.config(state="normal")
                MachDueDateBox.delete(0, END)
                MachDueDateBox.insert(0, val)
                MachDueDateBox.config(state="readonly")
                calWin.destroy()
            
            def emptyDate():
                MachDueDateBox.config(state="normal")
                MachDueDateBox.delete(0, END)
                MachDueDateBox.config(state="readonly")
                calWin.destroy()
    
            buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
            buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
            
            buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
            buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
            
            buttonClose = Button(calWin, text="Close", command=calWin.destroy)
            buttonClose.grid(row=1, column=2, padx=5, pady=5)
        
        def compCalMach():
            calWin = Toplevel()
            calWin.title("Select the Date")
            calWin.geometry("270x260")
            
            cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
            cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
            
            def confirmDate():
                val = cal.get_date()
                MachCompDateBox.config(state="normal")
                MachCompDateBox.delete(0, END)
                MachCompDateBox.insert(0, val)
                MachCompDateBox.config(state="readonly")
                calWin.destroy()
            
            def emptyDate():
                MachCompDateBox.config(state="normal")
                MachCompDateBox.delete(0, END)
                MachCompDateBox.config(state="readonly")
                calWin.destroy()
    
            buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
            buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
            
            buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
            buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
            
            buttonClose = Button(calWin, text="Close", command=calWin.destroy)
            buttonClose.grid(row=1, column=2, padx=5, pady=5)
        
        def updateOrderQty():
            selectMach = MachTreeView.selection()[0]
            curLoad = connLoad.cursor()
            curLoad.execute(f"SELECT MACH_ID FROM MACH_INDEX WHERE oid = '{selectMach}'")
            machResult = curLoad.fetchall()[0][0] # This Returns Machine ID Table
            
            curLoad.execute(f"SELECT oid, ASSEM_FULL, DES_QTY FROM `{machResult}`")
            assemResult = curLoad.fetchall()
            
            assemOidLst = []
            assemNameLst = []
            DesQtyLst = []
            for assem in assemResult:
                assemOidLst.append(assem[0])
                assemNameLst.append(f"{machResult}_{assem[1]}")
                DesQtyLst.append(assem[2])
            
            OrderQty = int(MachOrderBox.get())
            
            assemTotalPartLst = []
            assemTotalCostLst = []
            
            for i in range(len(assemNameLst)):
                curLoad.execute(f"SELECT * FROM `{assemNameLst[i]}` WHERE D = ''")
                unitData = curLoad.fetchall()
            
                def sumTotalCost(unitCst, num):
                    if unitCst == None or unitCst == "":
                        return 0
                    else:
                        totalVal = unitCst * num
                        return round(totalVal, 2)
                    
                def totalSGDConvert(totalVal, exRate):
                    if exRate == None or exRate == "":
                        return 0
                    else:
                        totalSGDVal = float(totalVal) * float(exRate)
                        return round(totalSGDVal, 2)
                
                unitOIDLst = []
                OHNewLst = []
                REQNewLst = []
                BALNewLst = []
                totalCostNewLst = []
                totalSGDNewLst = []
            
                for unit in unitData:
                    maxOH = ((unit[8])+(unit[9]))*(OrderQty)*(DesQtyLst[i])
                    if unit[10] > maxOH:
                        newOH = maxOH
                    else:
                        newOH = unit[10]
                                        
                    newREQ = (((unit[8])+(unit[9]))*(OrderQty)*(DesQtyLst[i]))-(unit[10])
                    if newREQ < 0:
                        newREQ = 0
                        
                    newBAL = newREQ - unit[12]
                    if newBAL < 0:
                        newBAL = 0
                    
                    newTotalCost = sumTotalCost(unit[18], newREQ)
                    newTotalSGD = totalSGDConvert(newTotalCost, unit[21])
                    
                    unitOIDLst.append(unit[0])
                    OHNewLst.append(newOH)
                    REQNewLst.append(newREQ)
                    BALNewLst.append(newBAL)
                    totalCostNewLst.append(newTotalCost)
                    totalSGDNewLst.append(newTotalSGD)
                
                OHTotalNew = sum(OHNewLst)
                REQTotalNew = sum(REQNewLst)
                totalSGDNew = sum(totalSGDNewLst)
                assemTotalPartLst.append(REQTotalNew)
                assemTotalCostLst.append(totalSGDNew)
            
                updateUnitCost = f"""UPDATE `{assemNameLst[i]}` SET
                `OH` = %s,
                `REQ` = %s,
                `BAL` = %s,
                `TotalUnitCost` = %s,
                `TotalSGD` = %s
                
                WHERE `oid` = %s"""
                
                for j in range(len(unitOIDLst)):
                    inputs = (OHNewLst[j], REQNewLst[j], BALNewLst[j], 
                              totalCostNewLst[j], totalSGDNewLst[j],
                              unitOIDLst[j])
                    curLoad.execute(updateUnitCost, inputs)
                connLoad.commit()
                
                updateAssemCost = f"""UPDATE `{machResult}` SET
                `NUM_PART` = %s,
                `TOTAL_COST` = %s,
                `OH_Qty` = %s
                
                WHERE `oid` = %s"""
                
                inputAssem = (REQTotalNew, totalSGDNew, OHTotalNew, assemOidLst[i])

                curLoad.execute(updateAssemCost, inputAssem)
                connLoad.commit()
            
            MachPartEst = sum(assemTotalPartLst)
            MachCostEst = sum(assemTotalCostLst)
            
            updateMachCost = f"""UPDATE `MACH_INDEX` SET
            `EST_PART_MACH` = %s,
            `EST_COST_MACH` = %s
                
            WHERE `oid` = %s"""
            
            inputMach = (MachPartEst, MachCostEst, selectMach)
            curLoad.execute(updateMachCost, inputMach)
            
            curLoad.close()
            
            clearEntryMach()
            MachTreeView.delete(*MachTreeView.get_children())
            queryTreeMach()

        def updateMach():
            sqlUpdateMach = f"""UPDATE MACH_INDEX SET
            `MACH_NAME` = %s,
            `LEAD_DESIGNER` = %s,
            `LEAD_LOCKER` = %s,
            `ORDER_QTY` = %s,
            `MACH_DES_DUE` = %s,
            `MACH_DES_COMP` = %s,
            `STATUS` = %s
            
            WHERE `oid` = %s"""
            
            MachName = MachNameBox.get().upper()
            MachLeadDes = EmpNameToOid(MachDesignerBox.get())
            MachLeadLock = EmpNameToOid(MachLockerBox.get())
            MachOrderQty = MachOrderBox.get()
            
            def checkDateMach(dateVar):
                if dateVar.get() == "":
                    return None
                else:
                    return dateVar.get()
            
            MachStat = MachStatusBox.current()
            
            selected = MachTreeView.selection()[0]
            
            inputs = (MachName, MachLeadDes, MachLeadLock, MachOrderQty,
                      checkDateMach(MachDueDateBox), checkDateMach(MachCompDateBox), 
                      MachStat, selected)
            
            respUpdateMach = messagebox.askokcancel("Confirmation",
                                                    "Update This Machine?",
                                                    parent=frameMach)
            if respUpdateMach == True:
                curLoad = connLoad.cursor()
                curLoad.execute(sqlUpdateMach, inputs)
                connLoad.commit()
                curLoad.close()
                
                updateOrderQty()
                
                clearEntryMach()
                MachTreeView.delete(*MachTreeView.get_children())
                queryTreeMach()
                
                checkProTotal()
                
                messagebox.showinfo("Update Successful", 
                                    f"You Have Updated This Machine", parent=frameMach) 
    
        def createMachCom():
            createComMach = """INSERT INTO MACH_INDEX (
            MACH_ID, MACH_NAME, LEAD_DESIGNER, LEAD_LOCKER, ORDER_QTY,
            MACH_DES_DUE, MACH_DES_COMP, STATUS)
            
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
            
            def checkDateMach(dateVar):
                if dateVar.get() == "":
                    return None
                else:
                    return dateVar.get()
            
            valueMach = (MachIDBox.get(), MachNameBox.get().upper(), 
                         EmpNameToOid(MachDesignerBox.get()), 
                         EmpNameToOid(MachLockerBox.get()), 
                         MachOrderBox.get(),
                         checkDateMach(MachDueDateBox), checkDateMach(MachCompDateBox), 
                         MachStatusBox.current())
            
            curLoad = connLoad.cursor()
            curLoad.execute(createComMach, valueMach)
            connLoad.commit()
            
            machID = MachIDBox.get()
            
            curLoad.execute(f""" CREATE TABLE IF NOT EXISTS `{machID}` (
                `oid` TINYINT AUTO_INCREMENT PRIMARY KEY,
                `ASSEM_TYPE` VARCHAR(3),
                `ASSEM_NUM` VARCHAR(5),
                `ASSEM_FULL` VARCHAR(10),
                `ASSEM_NAME` VARCHAR(50),
                `DES_STAT` INT,
                `PUR_STAT` INT,
                `ASSEM_STAT` INT,
                `DESIGNER` INT,
                `LOCKER` INT,
                `DES_DUE` DATE,
                `DES_COMP` DATE,
                `POI_DUE` DATE,
                `POI_COMP` DATE,
                `PART_DUE` DATE,
                `PART_COMP` DATE,
                `DES_APPROVE` BOOL,
                `APPROVE_NAME` INT,
                `NUM_PART` INT DEFAULT 0,
                `PART_PUR` INT DEFAULT 0,
                `PART_REC` INT DEFAULT 0,
                `NUM_DELETED` INT DEFAULT 0,
                `NUM_ONHOLD` INT DEFAULT 0,
                `DES_QTY` INT DEFAULT 0,
                `COMP_QTY` INT DEFAULT 0,
                `TOTAL_COST` FLOAT DEFAULT 0.00,
                `OH_Qty` INT DEFAULT 0)
            
                ENGINE = InnoDB
                DEFAULT CHARACTER SET = utf8mb4
                COLLATE = utf8mb4_0900_ai_ci""")
                
            connLoad.commit()
            curLoad.close()
            
            clearEntryMach()
            MachTreeView.delete(*MachTreeView.get_children())
            queryTreeMach()
            
            checkProTotal()
            
            messagebox.showinfo("Create Successful", 
                                f"You Have Created Machine {machID}", parent=frameMach)
        
        def createMach():
            if MachIDBox.get() == "":
                messagebox.showerror("Create Error", 
                                     "Invalid Machine Number", parent=frameMach) 
            else:
                if MachNameBox.get() == "":
                    respNameMach = messagebox.askokcancel("Yet to Enter Name", 
                                                          "Create a Machine Without Name?", parent=frameMach)
                    if respNameMach == True:
                        respMachA = messagebox.askokcancel("Confirmation", 
                                                           "Create This Machine?", parent=frameMach)
                        if respMachA == True:
                            createMachCom()
                else:
                    respMachB = messagebox.askokcancel("Confirmation", 
                                                       "Create This Machine?", parent=frameMach)
                    if respMachB == True:
                        createMachCom()
        
        def deleteMach():
            selected = MachTreeView.selection()[0]
            sqlSelect = "SELECT * FROM MACH_INDEX WHERE oid = %s"
            valSelect = (selected, )
            curLoad = connLoad.cursor()
            curLoad.execute(sqlSelect, valSelect)
            recLst = curLoad.fetchall()
            
            MachID = recLst[0][1]
            
            respDeleteMach = messagebox.askokcancel("Confirmation",
                                                    f"Delete Machine {MachID}?",
                                                    parent=frameMach)
            
            if respDeleteMach == True:
                curLoad.execute(f"SELECT * FROM `{MachID}`")
                assemInfo = curLoad.fetchall()
                
                assemNameLst = []
                for tup in assemInfo:
                    assemNameLst.append(f"{MachID}_{tup[3]}")            
                
                curLoad.execute(f"DROP TABLE IF EXISTS `{MachID}`")
                
                sqlDelete = "DELETE FROM MACH_INDEX WHERE oid = %s"
                valDelete = (selected, )
                curLoad.execute(sqlDelete, valDelete)
                
                for assem in assemNameLst:
                    curLoad.execute(f"DROP TABLE IF EXISTS `{assem}`")
                    
                connLoad.commit()
                curLoad.close()
                
                clearEntryMach()
                MachTreeView.delete(*MachTreeView.get_children())
                queryTreeMach()
                
                checkProTotal()
                
                messagebox.showinfo("Delete Successful", 
                                    f"You Have Deleted Machine {MachID}", parent=frameMach)
    
        def selectMach():        
            selectVal = MachTreeView.selection()
            if selectVal == ():
                messagebox.showerror("No Item Selected",
                                     "Please Select an Item",
                                     parent=frameMach)
            else:
                selected = selectVal[0]
                sqlSelect = "SELECT * FROM MACH_INDEX WHERE oid = %s"
                valSelect = (selected, )
                curLoad = connLoad.cursor()
                curLoad.execute(sqlSelect, valSelect)
                recLst = curLoad.fetchall()
        
                connLoad.commit()
                curLoad.close()
                
                clearEntryMach()
                
                MachIDBox.insert(0, recLst[0][1])
                MachNameBox.insert(0, recLst[0][2])
                
                MachDesignerBox.config(state="normal")
                MachDesignerBox.delete(0, END)
                MachDesignerBox.insert(0, EmpOidToNameInactive(recLst[0][3]))
                MachDesignerBox.config(state="readonly") 
                
                MachLockerBox.config(state="normal")
                MachLockerBox.delete(0, END)
                MachLockerBox.insert(0, EmpOidToNameInactive(recLst[0][4]))
                MachLockerBox.config(state="readonly") 
                
                MachOrderBox.config(state="normal")
                MachOrderBox.delete(0, END)
                MachOrderBox.insert(0, recLst[0][5])
                MachOrderBox.config(state="readonly")
                
                MachCompBox.config(state="normal")
                MachCompBox.insert(0, recLst[0][6])
                MachCompBox.config(state="readonly")
                
                if recLst[0][7] == None:
                    MachDueDateBox.config(state="normal")
                    MachDueDateBox.insert(0, "")
                    MachDueDateBox.config(state="readonly")
                else:
                    MachDueDateBox.config(state="normal")
                    MachDueDateBox.insert(0, recLst[0][7])
                    MachDueDateBox.config(state="readonly")
                if recLst[0][8] == None:
                    MachCompDateBox.config(state="normal")
                    MachCompDateBox.insert(0, "")
                    MachCompDateBox.config(state="readonly")
                else:
                    MachCompDateBox.config(state="normal")
                    MachCompDateBox.insert(0, recLst[0][8])
                    MachCompDateBox.config(state="readonly")    
                
                MachStatusBox.current(recLst[0][9])
        
                MachUABox.config(state="normal")
                MachUABox.insert(0, recLst[0][10])
                MachUABox.config(state="readonly")
                MachBOMBox.config(state="normal")
                MachBOMBox.insert(0, recLst[0][11])
                MachBOMBox.config(state="readonly")
                MachDeletedBox.config(state="normal")
                MachDeletedBox.insert(0, recLst[0][12])
                MachDeletedBox.config(state="readonly")
                MachOnHoldBox.config(state="normal")
                MachOnHoldBox.insert(0, recLst[0][13])
                MachOnHoldBox.config(state="readonly")
                
                MachEstPartBox.config(state="normal")
                MachEstPartBox.insert(0, recLst[0][14])
                MachEstPartBox.config(state="readonly")
                
                if AuthLevel(Login.AUTHLVL, 2) == False:
                    pass
                else:
                    MachEstCostBox.config(state="normal")
                    MachEstCostBox.insert(0, recLst[0][15])
                    MachEstCostBox.config(state="readonly")
                
                buttonUpdateMach.config(state=NORMAL)
                buttonCreateMach.config(state=DISABLED)
                buttonDeleteMach.config(state=NORMAL)
                MachIDBox.config(state="readonly")
                MachIDGenButton.config(state=DISABLED)
                MachTreeView.config(selectmode="none")
    
        def deselectMach():
            selected = MachTreeView.selection()
            if len(selected) > 0:
                MachTreeView.selection_remove(selected[0])
                clearEntryMach()
            else:
                clearEntryMach()
    
        def clearEntryMach():
            buttonUpdateMach.config(state=DISABLED)
            buttonCreateMach.config(state=NORMAL)
            buttonDeleteMach.config(state=DISABLED)
            MachIDBox.config(state="normal")
            MachIDGenButton.config(state=NORMAL)
            MachTreeView.config(selectmode="browse")
            
            MachIDBox.delete(0, END)
            MachNameBox.delete(0, END)
            MachStatusBox.current(0)
            
            MachDueDateBox.config(state="normal")
            MachDueDateBox.delete(0, END)
            MachDueDateBox.config(state="readonly")
            MachCompDateBox.config(state="normal")
            MachCompDateBox.delete(0, END)
            MachCompDateBox.config(state="readonly")
            
            MachDesignerBox.current(0)
            MachLockerBox.current(0)
            
            MachOrderBox.config(state="normal")
            MachOrderBox.delete(0, END)
            MachOrderBox.insert(0, 1)
            MachOrderBox.config(state="readonly")
            MachCompBox.config(state="normal")
            MachCompBox.delete(0, END)
            MachCompBox.config(state="readonly")
    
            MachUABox.config(state="normal")
            MachUABox.delete(0, END)
            MachUABox.config(state="readonly")
            MachBOMBox.config(state="normal")
            MachBOMBox.delete(0, END)
            MachBOMBox.config(state="readonly")
            MachDeletedBox.config(state="normal")
            MachDeletedBox.delete(0, END)
            MachDeletedBox.config(state="readonly")
            MachOnHoldBox.config(state="normal")
            MachOnHoldBox.delete(0, END)
            MachOnHoldBox.config(state="readonly")
            
            MachEstPartBox.config(state="normal")
            MachEstPartBox.delete(0, END)
            MachEstPartBox.config(state="readonly")
            MachEstCostBox.config(state="normal")
            MachEstCostBox.delete(0, END)
            MachEstCostBox.config(state="readonly")
    
        def refreshMach():
            MachDesignerBox.config(value=FetchEmployees())
            MachLockerBox.config(value=FetchEmployees())
            clearEntryMach()
            MachTreeView.delete(*MachTreeView.get_children())
            queryTreeMach()
            checkEmpMach()
    
        def closeTabMach():
            # frameMach.destroy()
            respCloseTabMach = messagebox.askokcancel("Confirmation",
                                                      "Close Machine Tab? This will also close ALL Tabs after it",
                                                      parent=frameMach)
            if respCloseTabMach == True:
                tabObjDict = tabNote.children.copy()
                tabkey = list(tabNote.children)
                tab_names = ["Assembly Selection","Machine Selection","Bill of Materials"]
                
                for Tab in tabObjDict:
                    
                    if tabNote.tab('.!notebook.' + Tab, "text") in tab_names:
                        
                        tabObjDict[Tab].destroy()
                    
                buttonLoadPro.config(state=NORMAL)
                
        def loadMach():
            selectMachVal = MachTreeView.selection()
            if selectMachVal == ():
                messagebox.showwarning("Unable to Load", "Please Select a Machine",
                                       parent=frameMach)
            else:
                loadMachCom()
                
        def loadMachCom():
            buttonLoadMach.config(state=DISABLED)
            
            loadMachSelect = MachTreeView.selection()[0]
            sqlSelect = "SELECT * FROM MACH_INDEX WHERE oid = %s"
            valSelect = (loadMachSelect, )
            curLoad = connLoad.cursor()
            curLoad.execute(sqlSelect, valSelect)
            resultSelect = curLoad.fetchall()
            connLoad.commit()
            curLoad.close()
            
            machName = resultSelect[0][1]
            machDesc = resultSelect[0][2]
            OrderQtyMach = int(resultSelect[0][5])
            
            frameAssem = Frame(tabNote)
            tabNote.add(frameAssem, text="Assembly Selection")
            frameAssem.columnconfigure(0, weight=1)
                    
            tabObjDict = tabNote.children
            tabNote.select(len(tabObjDict)-1)
           
            style = ttk.Style()
            style.theme_use("clam")
            
            style.configure("Treeview",
                            background="silver",
                            rowheight=20,
                            fieldbackground="light grey")
            
            style.map("Treeview")
            
            # AssemTabTitleLabel = Label(frameAssem, 
            #                            text=f"Assembly List for Project {proName} {proDesc} - Machine {machName} {machDesc}", 
            #                            font=("Arial", 12))
            
            AssemTabTitleLabel = Label(frameAssem, 
                                       text=f"Assembly List for Machine {proName} - {machName}", 
                                       font=("Arial", 12))
            AssemTabTitleLabel.grid(row=0, column=0, padx=0, pady=(10,3), ipadx=0, ipady=0, sticky=W+E) 
            AssemTreeFrame = Frame(frameAssem)
            AssemTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
            # AssemTreeFrame.pack()
            
            AssemTreeScroll = Scrollbar(AssemTreeFrame)
            AssemTreeScroll.pack(side=RIGHT, fill=Y)
            
            AssemTreeView = ttk.Treeview(AssemTreeFrame, yscrollcommand=AssemTreeScroll.set, 
                                         selectmode="browse")
            AssemTreeScroll.config(command=AssemTreeView.yview)
            # AssemTreeView.grid(row=2, column=0, columnspan=10, padx=10)
            AssemTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
            
            AssemTreeView['columns'] = ("Num", "Description", 
                                        "Design", "Purchase", "Assembly",
                                        "Designer", "Locker", "Approved", 
                                        "Num of Parts", "Parts Purchased", "Parts Received",
                                        "Deleted", "OnHold", "DesQty", "CompQty",
                                        "ApproxCost")

            # AssemTreeView.column("#0", anchor=CENTER, width=45)
            AssemTreeView.column("#0", width=0, stretch=NO)
            AssemTreeView.column("Num", anchor=CENTER, width=40)
            AssemTreeView.column("Description", anchor=W, width=80)
            AssemTreeView.column("Design", anchor=CENTER, width=75)
            AssemTreeView.column("Purchase", anchor=CENTER, width=75)
            AssemTreeView.column("Assembly", anchor=CENTER, width=75)
            AssemTreeView.column("Designer", anchor=W, width=90)
            AssemTreeView.column("Locker", anchor=W, width=90)
            AssemTreeView.column("Approved", anchor=CENTER, width=50)
            AssemTreeView.column("Num of Parts", anchor=CENTER, width=70)
            AssemTreeView.column("Parts Purchased", anchor=CENTER, width=80)
            AssemTreeView.column("Parts Received", anchor=CENTER, width=70)
            AssemTreeView.column("Deleted", anchor=CENTER, width=70)
            AssemTreeView.column("OnHold", anchor=CENTER, width=70)
            AssemTreeView.column("DesQty", anchor=CENTER, width=70)
            AssemTreeView.column("CompQty", anchor=CENTER, width=70)
            AssemTreeView.column("ApproxCost", anchor=CENTER, width=80)
            
            AssemTreeView.heading("#0", text="Index", anchor=CENTER)
            AssemTreeView.heading("Num", text="No.", anchor=CENTER) #0
            AssemTreeView.heading("Description", text="Description", anchor=CENTER) #1
            AssemTreeView.heading("Design", text="Design", anchor=CENTER) #2
            AssemTreeView.heading("Purchase", text="Purchase", anchor=CENTER) #3
            AssemTreeView.heading("Assembly", text="Assembly", anchor=CENTER) #4
            AssemTreeView.heading("Designer", text="Designer", anchor=CENTER) #5
            AssemTreeView.heading("Locker", text="Locker", anchor=CENTER) #6
            AssemTreeView.heading("Approved", text="Appd.", anchor=CENTER) #7
            AssemTreeView.heading("Num of Parts", text="Required", anchor=CENTER) #8
            AssemTreeView.heading("Parts Purchased", text="Purchased", anchor=CENTER) #9
            AssemTreeView.heading("Parts Received", text="Received", anchor=CENTER) #10
            AssemTreeView.heading("Deleted", text="Deleted", anchor=CENTER) #11
            AssemTreeView.heading("OnHold", text="On Hold", anchor=CENTER) #12
            AssemTreeView.heading("DesQty", text="Des. Qty", anchor=CENTER) #13
            AssemTreeView.heading("CompQty", text="Comp. Qty", anchor=CENTER) #14
            AssemTreeView.heading("ApproxCost", text="Est. Cost", anchor=CENTER) #15
            
            if AuthLevel(Login.AUTHLVL, 2) == False:
                AssemTreeView.config(displaycolumns = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14])
    




            def deselectAssemClick(e):
                deselectAssem()
            
            def selectAssemClick(e):
                selectAssem()
            
            def updateAssemReturn(e):
                if buttonUpdateAssem["state"] == "disabled":
                    messagebox.showwarning("Unable to Update", 
                                           "Please Select a Assembly", parent=frameAssem) 
                else:
                    updateAssem()
            
            def deleteAssemDel(e):
                if buttonDeleteAssem["state"] == "disabled":
                    messagebox.showwarning("Unable to Delete", 
                                           "Please Select a Assembly", parent=frameAssem) 
                else:
                    deleteAssem()

            def loadAssemClick(e):
                if buttonLoadAssem["state"] == "normal":
                    loadAssem()
    
            def closeAssemTabEsc(e):
                tabNum = tabNote.index("current")
                if tabNum == 2:
                    closeTabAssem()

            AssemTreeView.bind("<Button-3>", deselectAssemClick)
            AssemTreeView.bind("<Double-Button-1>", selectAssemClick)
            AssemTreeView.bind("<Return>", updateAssemReturn)
            AssemTreeView.bind("<Delete>", deleteAssemDel)
            AssemTreeView.bind("<Button-2>", loadAssemClick)
            AssemTreeView.bind("<Escape>", closeAssemTabEsc)   
    




            def queryTreeAssem():
                machRef = machName
                curLoad = connLoad.cursor()
                curLoad.execute(f"SELECT * FROM `{machRef}`")
                recLst = curLoad.fetchall()
                connLoad.commit()
                curLoad.close()

                statLst = ["NS", "IP", "Comp."]
                boolLst = ["No", "Yes"]
                
                def showDelOnHold(val):
                    if int(val) == 0:
                        return "N"
                    else:
                        return val
                
                for rec in recLst:
                    estCost = f"{rec[25]} SGD"
                    AssemTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                          values=(rec[3], rec[4], statLst[rec[5]], 
                                                  statLst[rec[6]], statLst[rec[7]], 
                                                  EmpOidToName(rec[8]), 
                                                  EmpOidToName(rec[9]), boolLst[rec[16]], 
                                                  f"{rec[18]} + {rec[26]}", 
                                                  rec[19], rec[20], 
                                                  showDelOnHold(rec[21]), showDelOnHold(rec[22]),
                                                  f"{OrderQtyMach} × {rec[23]}", 
                                                  rec[24], estCost))
            
            def checkMachTotal():
                machRef = machName

                curLoad = connLoad.cursor()
                curLoad.execute(f"SELECT * FROM `{machRef}`")
                recLst = curLoad.fetchall()
                
                UAMachLst = []
                BomLoadMachLst = []
                DeleteMachLst = []
                OnHoldMachLst = []
                EstPartMachLst = []
                EstCostMachLst = []
                setMachLst = []
                
                for assem in recLst:
                    UAMachLst.append(assem[0])
                    if assem[18] != 0 or assem[26] != 0:
                        BomLoadMachLst.append(assem[0])
                    DeleteMachLst.append(assem[21])
                    OnHoldMachLst.append(assem[22])
                    EstPartMachLst.append(assem[18])
                    EstCostMachLst.append(assem[25])
                    
                    if assem[23] == 0:
                        pass
                    else:
                        setMachLst.append(assem[24]//assem[23])
                
                UATotal = len(UAMachLst)
                BomLoadTotal = len(BomLoadMachLst)
                DeleteMachTotal = sum(DeleteMachLst)
                OnHoldMachTotal = sum(OnHoldMachLst)
                EstPartMachTotal = sum(EstPartMachLst)
                EstCostMachTotal = sum(EstCostMachLst)
                if len(setMachLst) == 0:
                    machCompQty = 0
                else:
                    machCompQty = min(setMachLst)
                
                sqlMachNum = f"""UPDATE `MACH_INDEX` SET
                `COMP_QTY` = %s,
                `NUM_UA` = %s,
                `NUM_BOM_LOAD` = %s,
                `NUM_DELETED` = %s,
                `NUM_ONHOLD` = %s,
                `EST_PART_MACH` = %s,
                `EST_COST_MACH` = %s
                
                WHERE `oid` = %s"""  
                
                inputs = (machCompQty, UATotal, BomLoadTotal, 
                          DeleteMachTotal,OnHoldMachTotal, 
                          EstPartMachTotal, EstCostMachTotal, loadMachSelect)
                
                curLoad.execute(sqlMachNum, inputs)
                connLoad.commit()
                curLoad.close()
                
                clearEntryMach()
                MachTreeView.delete(*MachTreeView.get_children())
                queryTreeMach()
            
            def genAssemID():
                machRef = machName
                TypeLst = ["M", "E", "P"]
                assemType = TypeLst[TypeBox.current()]
                sqlComAssemGen = f"""SELECT * FROM `{machRef}` WHERE ASSEM_TYPE = %s"""
                valAssemTup = (assemType, )
                
                curLoad = connLoad.cursor()
                curLoad.execute(sqlComAssemGen, valAssemTup)
                assemRefLst = curLoad.fetchall()
                connLoad.commit()
                curLoad.close()
                
                nextInt = len(assemRefLst) + 1
                twoDigitAssem = str(nextInt).rjust(2,'0')
                newAssemID = f"{twoDigitAssem}"
                
                NumBox.delete(0, END)
                NumBox.insert(0, newAssemID)
            
            def checkEmpAssem():
                if DesignerBox.get() == "Please Add Employees" \
                or LockerBox.get() == "Please Add Employees" \
                or ApproveNameBox.get() == "Please Add Employees":
                    buttonCreateAssem.config(state=DISABLED)
                else:
                    buttonCreateAssem.config(state=NORMAL)
                    
            def checkApproveName(name):
                if name == "NIL":
                    return None
                else:
                    newName = EmpNameToOid(name)
                    return newName
            
            def approveClick(event):
                val = DesignApprovedBox.get()
                if val == "No":
                    ApproveNameBox.config(value=["NIL"])
                    ApproveNameBox.current(0)
    
                else:
                    ApproveNameBox.config(value=FetchEmployees())
                    ApproveNameBox.current(0)

            def DesDueAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    DesignDueBox.config(state="normal")
                    DesignDueBox.delete(0, END)
                    DesignDueBox.insert(0, val)
                    DesignDueBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    DesignDueBox.config(state="normal")
                    DesignDueBox.delete(0, END)
                    DesignDueBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
            def POIDueAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    POIssueDueBox.config(state="normal")
                    POIssueDueBox.delete(0, END)
                    POIssueDueBox.insert(0, val)
                    POIssueDueBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    POIssueDueBox.config(state="normal")
                    POIssueDueBox.delete(0, END)
                    POIssueDueBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
            def PartDueAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    PartsReceivedDueBox.config(state="normal")
                    PartsReceivedDueBox.delete(0, END)
                    PartsReceivedDueBox.insert(0, val)
                    PartsReceivedDueBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    PartsReceivedDueBox.config(state="normal")
                    PartsReceivedDueBox.delete(0, END)
                    PartsReceivedDueBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
                
            def DesCompAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    DesignCompletedBox.config(state="normal")
                    DesignCompletedBox.delete(0, END)
                    DesignCompletedBox.insert(0, val)
                    DesignCompletedBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    DesignCompletedBox.config(state="normal")
                    DesignCompletedBox.delete(0, END)
                    DesignCompletedBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
            def POICompAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    POIssueCompletedBox.config(state="normal")
                    POIssueCompletedBox.delete(0, END)
                    POIssueCompletedBox.insert(0, val)
                    POIssueCompletedBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    POIssueCompletedBox.config(state="normal")
                    POIssueCompletedBox.delete(0, END)
                    POIssueCompletedBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
            
            def PartCompAssem():
                calWin = Toplevel()
                calWin.title("Select the Date")
                calWin.geometry("270x260")
                
                cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
                cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
                
                def confirmDate():
                    val = cal.get_date()
                    PartsReceivedCompletedBox.config(state="normal")
                    PartsReceivedCompletedBox.delete(0, END)
                    PartsReceivedCompletedBox.insert(0, val)
                    PartsReceivedCompletedBox.config(state="readonly")
                    calWin.destroy()
                
                def emptyDate():
                    PartsReceivedCompletedBox.config(state="normal")
                    PartsReceivedCompletedBox.delete(0, END)
                    PartsReceivedCompletedBox.config(state="readonly")
                    calWin.destroy()
        
                buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
                buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
                
                buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
                buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
                
                buttonClose = Button(calWin, text="Close", command=calWin.destroy)
                buttonClose.grid(row=1, column=2, padx=5, pady=5)
            
            def updateDesQty():
                machRef = machName
                curLoad = connLoad.cursor()
                
                curLoad.execute(f"SELECT * FROM `MACH_INDEX` WHERE MACH_ID = '{machRef}'")
                machResult = curLoad.fetchall()
                
                selectedUpdate = AssemTreeView.selection()[0]
                curLoad.execute(f"SELECT * FROM `{machRef}` WHERE oid = '{selectedUpdate}'")
                assemSelect = curLoad.fetchall()
                
                OrderQty = int(machResult[0][5])
                DesQty = int(DesignQtyBox.get())
                
                assemNameRef = f"{machRef}_{assemSelect[0][3]}"
                curLoad.execute(f"SELECT * FROM `{assemNameRef}` WHERE D = ''")
                unitData = curLoad.fetchall()
                
                def sumTotalCost(unitCst, num):
                    if unitCst == None or unitCst == "":
                        return 0
                    else:
                        totalVal = unitCst * num
                        return round(totalVal, 2)
                    
                def totalSGDConvert(totalVal, exRate):
                    if exRate == None or exRate == "":
                        return 0
                    else:
                        totalSGDVal = float(totalVal) * float(exRate)
                        return round(totalSGDVal, 2)
                
                unitOIDLst = []
                OHNewLst = []
                REQNewLst = []
                BALNewLst = []
                totalCostNewLst = []
                totalSGDNewLst = []
                
                for unit in unitData:
                    maxOH = ((unit[8])+(unit[9]))*(OrderQty)*(DesQty)
                    if unit[10] > maxOH:
                        newOH = maxOH
                    else:
                        newOH = unit[10]
                    
                    newREQ = (((unit[8])+(unit[9]))*(OrderQty)*(DesQty))-(newOH)
                    if newREQ < 0:
                        newREQ = 0
                        
                    newBAL = newREQ - unit[12]
                    if newBAL < 0:
                        newBAL = 0
                    
                    newTotalCost = sumTotalCost(unit[18], newREQ)
                    newTotalSGD = totalSGDConvert(newTotalCost, unit[21])
                    
                    unitOIDLst.append(unit[0])
                    OHNewLst.append(newOH)
                    REQNewLst.append(newREQ)
                    BALNewLst.append(newBAL)
                    totalCostNewLst.append(newTotalCost)
                    totalSGDNewLst.append(newTotalSGD)
                
                OHTotalNew = sum(OHNewLst)
                REQTotalNew = sum(REQNewLst)
                totalSGDNew = sum(totalSGDNewLst)
                
                sqlUnitCost = f"""UPDATE `{assemNameRef}` SET
                `OH` = %s,
                `REQ` = %s,
                `BAL` = %s,
                `TotalUnitCost` = %s,
                `TotalSGD` = %s
                
                WHERE `oid` = %s"""
                
                for i in range(len(unitOIDLst)):
                    inputs = (OHNewLst[i], REQNewLst[i], BALNewLst[i], 
                              totalCostNewLst[i], totalSGDNewLst[i],
                              unitOIDLst[i])
                    curLoad.execute(sqlUnitCost, inputs)
                connLoad.commit()
                
                sqlAssemCost = f"""UPDATE `{machRef}` SET
                `NUM_PART` = %s,
                `TOTAL_COST` = %s,
                `OH_Qty` = %s
                
                WHERE `oid` = %s"""
                
                inputAssem = (REQTotalNew, totalSGDNew, OHTotalNew, selectedUpdate)
                
                curLoad.execute(sqlAssemCost, inputAssem)
                connLoad.commit()
                curLoad.close()
                
                clearEntryAssem()
                AssemTreeView.delete(*AssemTreeView.get_children())
                queryTreeAssem()
            
            def updateAssem():
                machRef = machName
                sqlUpdateAssem = f"""UPDATE `{machRef}` SET
                `ASSEM_NAME` = %s,
                `DES_STAT` = %s,
                `PUR_STAT` = %s,
                `ASSEM_STAT` = %s,
                `DESIGNER` = %s,
                `LOCKER` = %s,
                `DES_DUE` = %s,
                `DES_COMP` = %s,
                `POI_DUE` = %s,
                `POI_COMP` = %s,
                `PART_DUE` = %s,
                `PART_COMP` = %s,
                `DES_APPROVE` = %s,
                `APPROVE_NAME` = %s,
                `DES_QTY` = %s
                
                WHERE `oid` = %s"""            
                
                def checkDateAssem(dateVar):
                    if dateVar.get() == "":
                        return None
                    else:
                        return dateVar.get()
                
                AssemName = DescBox.get().upper()
                DesStat = DesignBox.current()
                PurStat = PurchaseBox.current()
                AssemStat = AssemBox.current()
                Designer = EmpNameToOid(DesignerBox.get())
                Locker = EmpNameToOid(LockerBox.get())
                
                DesApprove = DesignApprovedBox.current()
                ApproveName = checkApproveName(ApproveNameBox.get())
                DesQty = DesignQtyBox.get()
                
                selected = AssemTreeView.selection()[0]
                
                inputs = (AssemName, DesStat, PurStat, AssemStat, Designer, Locker,
                          checkDateAssem(DesignDueBox), checkDateAssem(DesignCompletedBox), 
                          checkDateAssem(POIssueDueBox), checkDateAssem(POIssueCompletedBox), 
                          checkDateAssem(PartsReceivedDueBox), checkDateAssem(PartsReceivedCompletedBox), 
                          DesApprove, ApproveName, DesQty, selected)
                
                respUpdateAssem = messagebox.askokcancel("Confirmation",
                                                         "Update This Assembly?",
                                                         parent=frameAssem)
                
                if respUpdateAssem == True:       
                    assemVal = AssemTreeView.item(selected)["values"][0]
                    curLoad = connLoad.cursor()
                    curLoad.execute(sqlUpdateAssem, inputs)
                    connLoad.commit()
                    curLoad.close()
                    
                    updateDesQty()
                    
                    clearEntryAssem()
                    AssemTreeView.delete(*AssemTreeView.get_children())
                    queryTreeAssem()
                    
                    checkMachTotal()
                    checkProTotal()
                    
                    messagebox.showinfo("Update Successful",
                                        f"You Have Updated Assem {assemVal}",
                                        parent=frameAssem)
    
            def createAssemCom():
                machRef = machName
                createComAssem = f"""INSERT INTO `{machRef}` (
                ASSEM_TYPE, ASSEM_NUM, ASSEM_FULL, ASSEM_NAME, DES_STAT, PUR_STAT, 
                ASSEM_STAT, DESIGNER, LOCKER, DES_DUE, DES_COMP, POI_DUE, POI_COMP, 
                PART_DUE, PART_COMP, DES_APPROVE, APPROVE_NAME, DES_QTY)
                
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                        %s, %s, %s, %s, %s, %s, %s, %s)"""
                
                TypeLst = ["M", "E", "P"]
                
                def checkDateAssem(dateVar):
                    if dateVar.get() == "":
                        return None
                    else:
                        return dateVar.get()
                
                assemFullName = f"{TypeLst[TypeBox.current()]}{NumBox.get()}"
                
                valueAssem = (TypeLst[TypeBox.current()], NumBox.get(), assemFullName, DescBox.get().upper(),
                              DesignBox.current(), PurchaseBox.current(), AssemBox.current(),
                              EmpNameToOid(DesignerBox.get()), EmpNameToOid(LockerBox.get()),
                              checkDateAssem(DesignDueBox), checkDateAssem(DesignCompletedBox), 
                              checkDateAssem(POIssueDueBox), checkDateAssem(POIssueCompletedBox),
                              checkDateAssem(PartsReceivedDueBox), checkDateAssem(PartsReceivedCompletedBox), 
                              DesignApprovedBox.current(), checkApproveName(ApproveNameBox.get()),
                              DesignQtyBox.get())
                
                curLoad = connLoad.cursor()
                curLoad.execute(createComAssem, valueAssem)
                connLoad.commit()
                
                assemType = TypeLst[TypeBox.current()]
                assemNum = NumBox.get()
                assemFullName = f"{machRef}_{assemType}{assemNum}"         
                
                curLoad.execute(f""" CREATE TABLE IF NOT EXISTS `{assemFullName}` (
                    `oid` INT AUTO_INCREMENT PRIMARY KEY,
                    `PartNum` VARCHAR(30),
                    `Description` VARCHAR(100),
                    `D` VARCHAR(10),
                    `CLS` VARCHAR(20),
                    `V` INT DEFAULT 1,
                    `Maker` VARCHAR(100),
                    `Spec` VARCHAR(150),
                    `DES` INT DEFAULT 0,
                    `SPA` INT DEFAULT 0,
                    `OH` INT DEFAULT 0,
                    `REQ` INT DEFAULT 0,
                    `PCH` INT DEFAULT 0,
                    `BAL` INT DEFAULT 0,
                    `RCV` INT DEFAULT 0,
                    `OS` INT DEFAULT 0,
                    `REMARK` VARCHAR(100),
                    `Vendor` VARCHAR(100),
                    `UnitCost` FLOAT DEFAULT 0.00,
                    `TotalUnitCost` FLOAT DEFAULT 0.00,
                    `Currency` VARCHAR(30),
                    `ExchangeRate` FLOAT DEFAULT 0.00,
                    `TotalSGD` FLOAT DEFAULT 0.00)
                
                    ENGINE = InnoDB
                    DEFAULT CHARACTER SET = utf8mb4
                    COLLATE = utf8mb4_0900_ai_ci""")
                
                connLoad.commit()
                curLoad.close()
                clearEntryAssem()
                AssemTreeView.delete(*AssemTreeView.get_children())
                queryTreeAssem()
                
                checkMachTotal()
                checkProTotal()
                
                messagebox.showinfo("Create Successful",
                                    f"You Have Created Assem {assemFullName}",
                                    parent=frameAssem)
            
            def createAssem():
                if NumBox.get() == "":
                    messagebox.showerror("Create Error", 
                                         "Invalid Assembly Number", parent=frameAssem) 
                else:
                    if DescBox.get() == "":
                        respDescAssem = messagebox.askokcancel("Yet to Enter Description", 
                                                               "Create a Assembly Without Description?", parent=frameAssem)
                        if respDescAssem == True:
                            respAssemA = messagebox.askokcancel("Confirmation", 
                                                                "Create This Project?", parent=frameAssem)
                            if respAssemA == True:
                                createAssemCom()
                    else:
                        respAssemB = messagebox.askokcancel("Confirmation", 
                                                            "Create This Project?", parent=frameAssem)
                        if respAssemB == True:
                            createAssemCom()
            
            def deleteAssem():
                machRef = machName
                selected = AssemTreeView.selection()[0]
                sqlSelect = f"SELECT * FROM `{machRef}` WHERE oid = %s"
                valSelect = (selected, )
                curLoad = connLoad.cursor()
                curLoad.execute(sqlSelect, valSelect)
                recLst = curLoad.fetchall()
                connLoad.commit()
            
                assemType = recLst[0][1]
                assemName = recLst[0][2]
                
                respDeleteAssem = messagebox.askokcancel("Confirm",
                                                         "Delete this Assembly?",
                                                         parent=frameAssem)
                if respDeleteAssem == True:
                    assemFull = f"{machRef}_{assemType}{assemName}"
                    
                    curLoad.execute(f"DROP TABLE IF EXISTS `{assemFull}`")
                    
                    sqlDelete = f"DELETE FROM `{machRef}` WHERE oid = %s"
                    valDelete = (selected, )
                    curLoad.execute(sqlDelete, valDelete)
                    connLoad.commit()
                    curLoad.close()
                
                    clearEntryAssem()
                    AssemTreeView.delete(*AssemTreeView.get_children())
                    queryTreeAssem()
                    
                    checkMachTotal()
                    checkProTotal()
                    
                    messagebox.showinfo("Delete Successful",
                                        f"You Have Deleted Assembly {assemType}{assemName}",
                                        parent=frameAssem)
    
            def selectAssem():
                machRef = machName
                selectVal = AssemTreeView.selection()
                if selectVal == ():
                    messagebox.showerror("No Item Selected",
                                         "Please Select an Item",
                                         parent=frameAssem)
                    clearEntryAssem()
                else:
                    selected = selectVal[0]
                    sqlSelect = f"SELECT * FROM `{machRef}` WHERE oid = %s"
                    valSelect = (selected, )
                    curLoad = connLoad.cursor()
                    curLoad.execute(sqlSelect, valSelect)
                    recLst = curLoad.fetchall()
                    connLoad.commit()
                    curLoad.close()
            
                    clearEntryAssem()
                    
                    if recLst[0][1] == "M":
                        TypeBox.current(0)
                    elif recLst[0][1] == "E":
                        TypeBox.current(1)
                    elif recLst[0][1] == "P":
                        TypeBox.current(2)
                    
                    DesignBox.current(recLst[0][5])
                    PurchaseBox.current(recLst[0][6])
                    AssemBox.current(recLst[0][7])
                    NumBox.insert(0, recLst[0][2])
                    
                    DesignerBox.config(state="normal")
                    DesignerBox.delete(0, END)
                    DesignerBox.insert(0, EmpOidToNameInactive(recLst[0][8]))
                    DesignerBox.config(state="readonly")
                    
                    LockerBox.config(state="normal")
                    LockerBox.delete(0, END)
                    LockerBox.insert(0, EmpOidToNameInactive(recLst[0][9]))
                    LockerBox.config(state="readonly") 
                    
                    DescBox.insert(0, recLst[0][4])
                    
                    if recLst[0][10] == None:
                        DesignDueBox.config(state="normal")
                        DesignDueBox.insert(0, "")
                        DesignDueBox.config(state="readonly")
                    else:
                        DesignDueBox.config(state="normal")
                        DesignDueBox.insert(0, recLst[0][10])
                        DesignDueBox.config(state="readonly")
                    if recLst[0][12] == None:
                        POIssueDueBox.config(state="normal")
                        POIssueDueBox.insert(0, "")
                        POIssueDueBox.config(state="readonly")
                    else:
                        POIssueDueBox.config(state="normal")
                        POIssueDueBox.insert(0, recLst[0][12])
                        POIssueDueBox.config(state="readonly") 
                    if recLst[0][14] == None:
                        PartsReceivedDueBox.config(state="normal")
                        PartsReceivedDueBox.insert(0, "")
                        PartsReceivedDueBox.config(state="readonly")
                    else:
                        PartsReceivedDueBox.config(state="normal")
                        PartsReceivedDueBox.insert(0, recLst[0][14])
                        PartsReceivedDueBox.config(state="readonly")
                        
                    if recLst[0][11] == None:
                        DesignCompletedBox.config(state="normal")
                        DesignCompletedBox.insert(0, "")
                        DesignCompletedBox.config(state="readonly")
                    else:
                        DesignCompletedBox.config(state="normal")
                        DesignCompletedBox.insert(0, recLst[0][11])
                        DesignCompletedBox.config(state="readonly") 
                    if recLst[0][13] == None:
                        POIssueCompletedBox.config(state="normal")
                        POIssueCompletedBox.insert(0, "")
                        POIssueCompletedBox.config(state="readonly")
                    else:
                        POIssueCompletedBox.config(state="normal")
                        POIssueCompletedBox.insert(0, recLst[0][13])
                        POIssueCompletedBox.config(state="readonly")
                    if recLst[0][15] == None:
                        PartsReceivedCompletedBox.config(state="normal")
                        PartsReceivedCompletedBox.insert(0, "")
                        PartsReceivedCompletedBox.config(state="readonly")
                    else:
                        PartsReceivedCompletedBox.config(state="normal")
                        PartsReceivedCompletedBox.insert(0, recLst[0][15])
                        PartsReceivedCompletedBox.config(state="readonly") 
                        
                    DesignApprovedBox.current(recLst[0][16])
    
                    if recLst[0][17] == None:
                        ApproveNameBox.config(state="normal")
                        ApproveNameBox.config(value="NIL")
                        ApproveNameBox.current(0)
                        ApproveNameBox.config(state="readonly") 
                    else:
                        ApproveNameBox.config(state="normal")
                        ApproveNameBox.config(value=FetchEmployees())
                        ApproveNameBox.delete(0, END)
                        ApproveNameBox.insert(0, EmpOidToNameInactive(recLst[0][17]))
                        ApproveNameBox.config(state="readonly")  
                    
                    NumOfPartsBox.config(state="normal")
                    NumOfPartsBox.insert(0, f"{recLst[0][18]} + {recLst[0][26]}")
                    NumOfPartsBox.config(state="readonly")
                    PartsPurchasedBox.config(state="normal")
                    PartsPurchasedBox.insert(0, recLst[0][19])
                    PartsPurchasedBox.config(state="readonly")
                    PartsReceivedBox.config(state="normal")
                    PartsReceivedBox.insert(0, recLst[0][20])
                    PartsReceivedBox.config(state="readonly")
                    
                    DeletedBox.config(state="normal")
                    DeletedBox.insert(0, recLst[0][21])
                    DeletedBox.config(state="readonly")
                    OnHoldBox.config(state="normal")
                    OnHoldBox.insert(0, recLst[0][22])
                    OnHoldBox.config(state="readonly")
                    
                    DesignQtyBox.config(state="normal")
                    DesignQtyBox.delete(0, END)
                    DesignQtyBox.insert(0, recLst[0][23])
                    DesignQtyBox.config(state="readonly")
                    CompletedBox.config(state="normal")
                    CompletedBox.delete(0, END)
                    CompletedBox.insert(0, recLst[0][24])
                    CompletedBox.config(state="readonly")
                    
                    if AuthLevel(Login.AUTHLVL, 2) == False:
                        pass
                    else:
                        EstCostAssemBox.config(state="normal")
                        EstCostAssemBox.insert(0, recLst[0][25])
                        EstCostAssemBox.config(state="readonly")
                    
                    buttonUpdateAssem.config(state=NORMAL)
                    buttonCreateAssem.config(state=DISABLED)
                    buttonDeleteAssem.config(state=NORMAL)
                    NumBox.config(state="readonly")
                    AssemIDGenButton.config(state=DISABLED)
                    AssemTreeView.config(selectmode="none")
                
            def deselectAssem():
                selected = AssemTreeView.selection()
                if len(selected) > 0:
                    AssemTreeView.selection_remove(selected[0])
                    clearEntryAssem()
                else:
                    clearEntryAssem()
    
            def clearEntryAssem():
                buttonUpdateAssem.config(state=DISABLED)
                buttonCreateAssem.config(state=NORMAL)
                buttonDeleteAssem.config(state=DISABLED)
                NumBox.config(state="normal")
                AssemIDGenButton.config(state=NORMAL)
                AssemTreeView.config(selectmode="extended")
                
                TypeBox.current(0)
                DesignBox.current(0)
                PurchaseBox.current(0)
                AssemBox.current(0)
                NumBox.delete(0, END)
                DesignerBox.current(0)
                LockerBox.current(0)
                DescBox.delete(0, END)
                
                DesignDueBox.config(state="normal")
                DesignDueBox.delete(0, END)
                DesignDueBox.config(state="readonly")
                POIssueDueBox.config(state="normal")
                POIssueDueBox.delete(0, END)
                POIssueDueBox.config(state="readonly")
                PartsReceivedDueBox.config(state="normal")
                PartsReceivedDueBox.delete(0, END)
                PartsReceivedDueBox.config(state="readonly")
                DesignCompletedBox.config(state="normal")
                DesignCompletedBox.delete(0, END)
                DesignCompletedBox.config(state="readonly")
                POIssueCompletedBox.config(state="normal")
                POIssueCompletedBox.delete(0, END)
                POIssueCompletedBox.config(state="readonly")
                PartsReceivedCompletedBox.config(state="normal")
                PartsReceivedCompletedBox.delete(0, END)
                PartsReceivedCompletedBox.config(state="readonly")
                
                DesignApprovedBox.current(0)
                ApproveNameBox.config(value=["NIL"])
                ApproveNameBox.current(0)
                
                NumOfPartsBox.config(state="normal")
                NumOfPartsBox.delete(0, END)
                NumOfPartsBox.config(state="readonly")
                PartsPurchasedBox.config(state="normal")
                PartsPurchasedBox.delete(0, END)
                PartsPurchasedBox.config(state="readonly")
                PartsReceivedBox.config(state="normal")
                PartsReceivedBox.delete(0, END)
                PartsReceivedBox.config(state="readonly")
                
                DeletedBox.config(state="normal")
                DeletedBox.delete(0, END)
                DeletedBox.config(state="readonly")
                OnHoldBox.config(state="normal")
                OnHoldBox.delete(0, END)
                OnHoldBox.config(state="readonly")

                DesignQtyBox.config(state="normal")
                DesignQtyBox.delete(0, END)
                DesignQtyBox.insert(0, 1)
                DesignQtyBox.config(state="readonly")
                CompletedBox.config(state="normal")
                CompletedBox.delete(0, END)
                CompletedBox.insert(0, 0)
                CompletedBox.config(state="readonly")
                
                EstCostAssemBox.config(state="normal")
                EstCostAssemBox.delete(0, END)
                EstCostAssemBox.config(state="readonly")
            
            def refreshAssem():
                DesignerBox.config(value=FetchEmployees())
                LockerBox.config(value=FetchEmployees())
                clearEntryAssem()
                AssemTreeView.delete(*AssemTreeView.get_children())
                queryTreeAssem()
                checkEmpAssem()
    
            def closeTabAssem():
                # frameAssem.destroy()
                respCloseTabAssem = messagebox.askokcancel("Confirmation",
                                                           "Close Assembly Tab? This will also close ALL Tabs after it",
                                                           parent=frameAssem)
                if respCloseTabAssem == True:
                    tabObjDict = tabNote.children.copy()
                    tabkey = list(tabNote.children)
                    tab_names = ["Assembly Selection","Bill of Materials"]
                    
                    for Tab in tabObjDict:
                        
                        if tabNote.tab('.!notebook.' + Tab, "text") in tab_names:
                            
                            tabObjDict[Tab].destroy()
                    buttonLoadMach.config(state=NORMAL)

            def loadAssem():
                selectAssemVal = AssemTreeView.selection()
                if selectAssemVal == ():
                    messagebox.showwarning("Unable to Load", "Please Select a Assembly",
                                           parent=frameAssem)
                else:
                    loadAssemCom()

            def loadAssemCom():
                buttonLoadAssem.config(state=DISABLED)
                
                machRef = machName
                curLoad = connLoad.cursor()
                curLoad.execute(f"SELECT * FROM `MACH_INDEX` WHERE MACH_ID = '{machRef}'")
                machResult = curLoad.fetchall()
                
                loadAssemSelect = AssemTreeView.selection()[0]
                sqlSelect = f"SELECT * FROM `{machRef}` WHERE oid = %s"
                valSelect = (loadAssemSelect, )
                
                curLoad.execute(sqlSelect, valSelect)
                resultSelect = curLoad.fetchall()
                connLoad.commit()
                curLoad.close()            

                def checkQty(Qty):
                    if Qty == None or Qty == "":
                        return 0
                    else:
                        return int(Qty)
                
                assemType = resultSelect[0][1]
                assemName = resultSelect[0][2]
                assemDesc = resultSelect[0][4]
                orderQty = checkQty(machResult[0][5])
                desQty = checkQty(resultSelect[0][23])
                
                assemFullName = f"{machRef}_{assemType}{assemName}"
                assemFullNoMach = f"{assemType}{assemName}"
                
                frameUnit = Frame(tabNote)
                tabNote.add(frameUnit, text="Bill of Materials")
                frameUnit.columnconfigure(0, weight=1)
                        
                tabObjDict = tabNote.children
                tabNote.select(len(tabObjDict)-1)
                   
                style = ttk.Style()
                style.theme_use("clam")
                
                style.configure("Treeview",
                                background="silver",
                                rowheight=20,
                                fieldbackground="light grey")
                
                style.map("Treeview")
                
                # UnitTabTitleLabel = Label(frameUnit, 
                #                           text=f"Bill of Material for Project {proName} {proDesc} - Machine {machName} {machDesc} - Assembly {assemType}{assemName} {assemDesc}", 
                #                           font=("Arial", 12))
                
                UnitTabTitleLabel = Label(frameUnit, 
                                          text=f"Bill of Material for Assembly {proName} - {machName} - {assemType}{assemName}", 
                                          font=("Arial", 12))
                
                UnitTabTitleLabel.grid(row=0, column=0, padx=0, pady=(10,3), ipadx=0, ipady=0, sticky=W+E) 
                UnitTreeFrame = Frame(frameUnit)
                UnitTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
                # UnitTreeFrame.pack()
                
                UnitTreeScroll = Scrollbar(UnitTreeFrame)
                UnitTreeScroll.pack(side=RIGHT, fill=Y)
                
                # UnitTreeView = CheckboxTreeview(UnitTreeFrame, yscrollcommand=UnitTreeScroll.set, 
                #                             selectmode="browse")
                UnitTreeView = ttk.Treeview(UnitTreeFrame, yscrollcommand=UnitTreeScroll.set, 
                                            selectmode="browse")
                UnitTreeScroll.config(command=UnitTreeView.yview)
                # UnitTreeView.grid(row=2, column=0, columnspan=10, padx=10)
                UnitTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
                
                UnitTreeView["columns"] = ("Part", "Description", "D", "CLS", "V", 
                                            "Maker", "Spec", "DES", "SPA", "OH", "REQ", 
                                            "PCH", "BAL", "RCV", "OS",
                                            "Remark", "Vendor", "UnitCost",
                                            "TotalCost")
                
                UnitTreeView.column("#0", width=0, stretch=NO)
                # UnitTreeView.column("#0", width=50, anchor=W,stretch=NO)
                UnitTreeView.column("Part", anchor=CENTER, width=45)
                UnitTreeView.column("Description", anchor=W, width=120)
                UnitTreeView.column("D", anchor=CENTER, width=30)
                UnitTreeView.column("CLS", anchor=CENTER, width=50)
                UnitTreeView.column("V", anchor=CENTER, width=30)
                UnitTreeView.column("Maker", anchor=W, width=90)
                UnitTreeView.column("Spec", anchor=W, width=180)
                UnitTreeView.column("DES", anchor=CENTER, width=35)
                UnitTreeView.column("SPA", anchor=CENTER, width=35)
                UnitTreeView.column("OH", anchor=CENTER, width=35)
                UnitTreeView.column("REQ", anchor=CENTER, width=35)
                UnitTreeView.column("PCH", anchor=CENTER, width=35)
                UnitTreeView.column("BAL", anchor=CENTER, width=35)
                UnitTreeView.column("RCV", anchor=CENTER, width=35)
                UnitTreeView.column("OS", anchor=CENTER, width=35)
                UnitTreeView.column("Remark", anchor=W, width=100)
                UnitTreeView.column("Vendor", anchor=W, width=100)
                UnitTreeView.column("UnitCost", anchor=CENTER, width=80)
                UnitTreeView.column("TotalCost", anchor=CENTER, width=80)
                
                UnitTreeView.heading("#0", text="Index", anchor=W)
                UnitTreeView.heading("Part", text="Part", anchor=CENTER) #0
                UnitTreeView.heading("Description", text="Description", anchor=CENTER) #1
                UnitTreeView.heading("D", text="D", anchor=CENTER) #2
                UnitTreeView.heading("CLS", text="CLS", anchor=CENTER) #3
                UnitTreeView.heading("V", text="V", anchor=CENTER) #4
                UnitTreeView.heading("Maker", text="Maker", anchor=CENTER) #5
                UnitTreeView.heading("Spec", text="Maker Specification", anchor=CENTER) #6
                UnitTreeView.heading("DES", text="DES", anchor=CENTER) #7
                UnitTreeView.heading("SPA", text="SPA", anchor=CENTER) #8
                UnitTreeView.heading("OH", text="OH", anchor=CENTER) #9
                UnitTreeView.heading("REQ", text="REQ", anchor=CENTER) #10
                UnitTreeView.heading("PCH", text="PCH", anchor=CENTER) #11
                UnitTreeView.heading("BAL", text="BAL", anchor=CENTER) #12
                UnitTreeView.heading("RCV", text="RCV", anchor=CENTER) #13
                UnitTreeView.heading("OS", text="OS", anchor=CENTER) #14
                UnitTreeView.heading("Remark", text="Remark", anchor=CENTER) #15
                UnitTreeView.heading("Vendor", text="Vendor", anchor=CENTER) #16
                UnitTreeView.heading("UnitCost", text="Unit Cost", anchor=CENTER) #17
                UnitTreeView.heading("TotalCost", text="Total Cost", anchor=CENTER) #18
                
                if AuthLevel(Login.AUTHLVL, 0) == False:
                    UnitTreeView.config(displaycolumns = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16])
                
                def deselectUnitClick(e):
                    deselectUnit()
                
                def selectUnitClick(e):
                    selectUnit()
                
                def updateUnitReturn(e):
                    if buttonUpdateUnit["state"] == "disabled":
                        messagebox.showwarning("Unable to Update", 
                                               "Please Select a Part", parent=frameUnit) 
                    else:
                        updateUnit()
                
                def deleteUnitDel(e):
                    if buttonDeleteUnit["state"] == "disabled":
                        messagebox.showwarning("Unable to Delete", 
                                               "Please Select a Part", parent=frameUnit) 
                    else:
                        deleteUnit()
                
                def setDeleteUnit(e):
                    if buttonUpdateUnit["state"] == "disabled":
                        messagebox.showwarning("Unable to Update", 
                                               "Please Select a Part", parent=frameUnit)
                    else:
                        assemRef = assemFullName
                        selected = UnitTreeView.selection()[0]
                        PartNumRef = str(UnitTreeView.item(selected)["values"][0]).rjust(3,"0")
                        
                        curLoad = connLoad.cursor()
                        curLoad.execute(f"SELECT `D` FROM `{assemRef}` WHERE oid = '{selected}'")
                        currentDVal = curLoad.fetchall()[0][0]
                        
                        updateD = f"""UPDATE `{assemRef}` SET
                        D = %s
                        WHERE oid = %s
                        """
                        
                        if currentDVal == "":
                            curLoad.execute(updateD, ("D", selected))
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Update Successful", 
                                                f"You Have Set Part {PartNumRef} as Deleted", parent=frameUnit) 
                    
                        else:
                            curLoad.execute(updateD, ("", selected))
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Update Successful", 
                                                f"You Have Set Part {PartNumRef} as Normal", parent=frameUnit)
                
                def closeUnitTabEsc(e):
                    tabNum = tabNote.index("current")
                    if tabNum == 3:
                        closeTabUnit()
                

                
                UnitTreeView.bind("<Button-3>", deselectUnitClick)
                UnitTreeView.bind("<Double-Button-1>", selectUnitClick)
                UnitTreeView.bind("<Return>", updateUnitReturn)
                UnitTreeView.bind("<Delete>", deleteUnitDel)
                UnitTreeView.bind("<Control-d>", setDeleteUnit)
                UnitTreeView.bind("<Escape>", closeUnitTabEsc)
                
                
                
                def checkAssemTotal():
                    machRef = machName
                    assemRef = assemFullName
                    
                    curLoad = connLoad.cursor()
                    curLoad.execute(f"SELECT * FROM `{assemRef}` WHERE D = ''")
                    recLst = curLoad.fetchall()
                    
                    costAssemLst = []
                    REQAssemLst = []
                    OHAssemLst = []
                    PCHAssemLst = []
                    RCVAssemLst = []
                    setAssemLst = []
                    
                    for unit in recLst:
                        costAssemLst.append(float(unit[22]))
                        REQAssemLst.append(int(unit[11]))
                        OHAssemLst.append(int(unit[10]))
                        PCHAssemLst.append(int(unit[12]))
                        RCVAssemLst.append(int(unit[14]))
                        totalCurrent = int(unit[10]) + int(unit[14])
                        oneSetNeed = int(unit[8]) + int(unit[9])
                        if oneSetNeed == 0:
                            pass
                        else:
                            setAssemLst.append(totalCurrent//oneSetNeed)
                        
                    curLoad.execute(f"SELECT * FROM `{assemRef}` WHERE D = 'D'")
                    deleteAssemLst = curLoad.fetchall()
                    
                    curLoad.execute(f"SELECT * FROM `{assemRef}` WHERE REMARK = 'ON HOLD'")
                    onHoldAssemLst = curLoad.fetchall()

                    costAssemTotal = sum(costAssemLst)
                    REQAssemTotal = sum(REQAssemLst)
                    OHAssemTotal = sum(OHAssemLst)
                    PCHAssemTotal = sum(PCHAssemLst)
                    RCVAssemTotal = sum(RCVAssemLst)
                    
                    OnHoldAssemTotal = len(onHoldAssemLst)
                    deleteAssemTotal = len(deleteAssemLst)
                    
                    if len(setAssemLst) == 0:
                        assemCompQty = 0
                    else:
                        assemCompQty = min(setAssemLst)
                    
                    sqlAssemNum = f"""UPDATE `{machRef}` SET
                    `NUM_PART` = %s,
                    `PART_PUR` = %s,
                    `PART_REC` = %s,
                    `NUM_DELETED` = %s,
                    `NUM_ONHOLD` = %s,
                    `COMP_QTY` = %s,
                    `TOTAL_COST` = %s,
                    `OH_Qty` = %s
                    
                    WHERE `oid` = %s"""  
                    
                    inputs = (REQAssemTotal, PCHAssemTotal, RCVAssemTotal,
                              deleteAssemTotal, OnHoldAssemTotal, assemCompQty,
                              costAssemTotal, OHAssemTotal, loadAssemSelect)
                    
                    curLoad.execute(sqlAssemNum, inputs)
                    connLoad.commit()
                    curLoad.close()
                    
                    clearEntryAssem()
                    AssemTreeView.delete(*AssemTreeView.get_children())
                    queryTreeAssem()
                
                def queryTreeUnit():
                    assemRef = assemFullName
                    curLoad = connLoad.cursor()
                    curLoad.execute(f"SELECT * FROM `{assemRef}`")
                    recLst = curLoad.fetchall()
                    connLoad.commit()
                    curLoad.close()
                    
                    def checkCurrencyNone(num, ccy):
                        if num == None or num == "":
                            return num
                        else:
                            numCurrency = f"{num} {ccy}"
                            return numCurrency
                    
                    for rec in recLst:
                        UnitTreeView.insert(parent="", index=END, iid=rec[0], 
                                            values=(rec[1], rec[2], rec[3], rec[4], rec[5], rec[6], rec[7], 
                                                    rec[8], rec[9], rec[10], rec[11], rec[12], rec[13], 
                                                    rec[14], rec[15], rec[16], rec[17], 
                                                    checkCurrencyNone(rec[18], rec[20]), 
                                                    f"{rec[19]} {rec[20]}"))
        
                def fetchUnitClass():
                    connVend = mysql.connector.connect(host = logininfo[0],
                                                       user = logininfo[1], 
                                                       password =logininfo[2],
                                                       database= "INDEX_VEND_MASTER")
                    curVend = connVend.cursor()
                    curVend.execute("""SELECT * FROM `CLASS_LIST`""")
                    ClassVal = curVend.fetchall()
                    
                    if ClassVal == []:
                        ClassValList = ["Please Input Class"]
                        return ClassValList
                    else:
                        ClassValList = []
                        for tup in ClassVal:
                            ClassValList.append(tup[1])
                        return ClassValList
                    
                    connVend.commit()
                    curVend.close()
                
                def updateUnit():
                    if AuthLevel(Login.AUTHLVL, 0) == False:
                        updateUnitComNoCost()
                    else:
                        updateUnitComCost()
                
                def updateUnitComNoCost():
                    assemRef = assemFullName
                    sqlCommand = f"""UPDATE `{assemRef}` SET
                    PartNum = %s,
                    Description = %s,
                    D = %s,
                    CLS = %s,
                    V = %s,
                    Maker = %s,
                    Spec = %s,
                    DES = %s,
                    SPA = %s,
                    OH = %s,
                    REQ = %s,
                    PCH = %s,
                    BAL = %s,
                    RCV = %s,
                    OS = %s,
                    REMARK = %s,
                    Vendor = %s

                    WHERE oid = %s
                    """

                    DLst = ["", "D"]
                    
                    PartNum = str(PartNumBox.get()).rjust(3,"0")
                    Description = DescriptionBox.get().upper()
                    D = DLst[DBox.current()]
                    CLS = CLSBox.get()
                    V = VBox.get()
                    Maker = MakerBox.get()
                    Spec = SpecBox.get().upper()
                    DES = int(DESBox.get())
                    SPA = int(SPABox.get())
                    OH = int(OHBox.get())
                    REQ = int(REQBox.get())
                    PCH = int(PCHBox.get())
                    BAL = int(BALBox.get())
                    RCV = int(RCVBox.get())
                    OS = int(OSBox.get())
                    Remark = RemarkBox.get()
                    Vendor = VendorBox.get()

                    selected = UnitTreeView.selection()[0]
                    
                    inputs = (PartNum, Description, D, CLS, V,
                              Maker, Spec, DES, SPA, OH, REQ, PCH,
                              BAL, RCV, OS, Remark, Vendor, selected)
                    
                    updateUnitFormat = True
                    try:
                        int(PartNum)
                    except:
                        updateUnitFormat = False
                    
                    if updateUnitFormat == False:
                        messagebox.showerror("Unable to Update",
                                             "Please Key in an Integer",
                                             parent=frameUnit)
                    
                    else:
                        respUpdateUnit = messagebox.askokcancel("Confirmation", "Confirm Update")
                        if respUpdateUnit == True:
                            curLoad = connLoad.cursor()
                            curLoad.execute(sqlCommand, inputs)
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Update Successful", 
                                                f"You Have Updated Part {PartNum}", parent=frameUnit) 
                        else:
                            pass
                
                def updateUnitComCost():
                    assemRef = assemFullName
                    sqlCommand = f"""UPDATE `{assemRef}` SET
                    PartNum = %s,
                    Description = %s,
                    D = %s,
                    CLS = %s,
                    V = %s,
                    Maker = %s,
                    Spec = %s,
                    DES = %s,
                    SPA = %s,
                    OH = %s,
                    REQ = %s,
                    PCH = %s,
                    BAL = %s,
                    RCV = %s,
                    OS = %s,
                    REMARK = %s,
                    Vendor = %s,
                    UnitCost = %s,
                    TotalUnitCost = %s,
                    Currency = %s,
                    ExchangeRate = %s,
                    TotalSGD = %s

                    WHERE oid = %s
                    """
                    
                    def sumTotalCost(unit, num):
                        if unit == None or unit == "":
                            return 0
                        else:
                            return unit * num

                    def checkCost(cost):
                        if cost == "":
                            return None
                        else:
                            return float(cost)
 
                    def totalSGDConvert(totalVal, exRate):
                        if exRate == None or exRate == "":
                            return 0
                        else:
                            totalSGDVal = float(totalVal) * float(exRate)
                            return round(totalSGDVal, 2)

                    DLst = ["", "D"]
                    
                    PartNum = str(PartNumBox.get()).rjust(3,"0")
                    Description = DescriptionBox.get().upper()
                    D = DLst[DBox.current()]
                    CLS = CLSBox.get()
                    V = VBox.get()
                    Maker = MakerBox.get()
                    Spec = SpecBox.get().upper()
                    DES = int(DESBox.get())
                    SPA = int(SPABox.get())
                    OH = int(OHBox.get())
                    REQ = int(REQBox.get())
                    PCH = int(PCHBox.get())
                    BAL = int(BALBox.get())
                    RCV = int(RCVBox.get())
                    OS = int(OSBox.get())
                    Remark = RemarkBox.get()
                    Vendor = VendorBox.get()
                    
                    UnitCost = checkCost(UnitCostBox.get())
                    TotalUnitCost = sumTotalCost(UnitCost, REQ)
                    Currency = CurrencyUnitBox.get()
                    ExchangeRate = ExRateBox.get()
                    TotalSGD = totalSGDConvert(TotalUnitCost, ExchangeRate)

                    selected = UnitTreeView.selection()[0]
                    
                    inputs = (PartNum, Description, D, CLS, V,
                              Maker, Spec, DES, SPA, OH, REQ, PCH,
                              BAL, RCV, OS, Remark, Vendor, UnitCost, 
                              TotalUnitCost, Currency, ExchangeRate, 
                              TotalSGD, selected)
                    
                    updateUnitFormat = True
                    try:
                        int(PartNum)
                    except:
                        updateUnitFormat = False
                    
                    if updateUnitFormat == False:
                        messagebox.showerror("Unable to Update",
                                             "Please Key in an Integer",
                                             parent=frameUnit)
                    
                    else:
                        respUpdateUnit = messagebox.askokcancel("Confirmation", "Confirm Update")
                        if respUpdateUnit == True:
                            curLoad = connLoad.cursor()
                            curLoad.execute(sqlCommand, inputs)
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Update Successful", 
                                                f"You Have Updated Part {PartNum}", parent=frameUnit) 
                        else:
                            pass
                
                def createUnit():
                    if AuthLevel(Login.AUTHLVL, 0) == False:
                        createUnitNoCost()
                    else:
                        createUnitCost()
                
                def createUnitNoCost():
                    assemRef = assemFullName
                    
                    createUnitFormat = True
                    try:
                        int(PartNumBox.get())
                    except:
                        createUnitFormat = False
                    
                    if PartNumBox.get() == "":
                        messagebox.showerror("Unable to Create Unit",
                                             "Please Key In a Part Number",
                                             parent=frameUnit)
                        
                    elif createUnitFormat == False:
                        messagebox.showerror("Incorrect Part Number Format",
                                             "Please Key in an Integer",
                                             parent=frameUnit)
                    
                    else:
                        respCreateUnit = messagebox.askokcancel("Confirmation",
                                                                "Add New Part?",
                                                                parent=frameUnit)
                        if respCreateUnit == True:
                            sqlCommand = f"""INSERT INTO `{assemRef}` (
                            PartNum, Description, D, CLS, V,
                            Maker, Spec, DES, SPA, OH, REQ,
                            PCH, BAL, RCV, OS, REMARK, Vendor, UnitCost, 
                            TotalUnitCost, Currency, ExchangeRate, TotalSGD)
                    
                            VALUES (%s, %s, %s, %s, %s, %s, %s, 
                                    %s, %s, %s, %s, %s, %s, %s, 
                                    %s, %s, %s, %s, %s, %s, %s, %s)"""
                            
                            def sumTotalCost(unit, num):
                                if unit == None or unit == "":
                                    return 0
                                else:
                                    return unit * num
        
                            def checkCost(cost):
                                if cost == "":
                                    return None
                                else:
                                    return float(cost)
         
                            def totalSGDConvert(totalVal, exRate):
                                if exRate == None or exRate == "":
                                    return 0
                                else:
                                    totalSGDVal = float(totalVal) * float(exRate)
                                    return round(totalSGDVal, 2)
                            
                            DLst = ["", "D"]
                            
                            PartNum = str(PartNumBox.get()).rjust(3,"0")
                            Description = DescriptionBox.get().upper()
                            D = DLst[DBox.current()]
                            CLS = CLSBox.get()
                            V = VBox.get()
                            Maker = MakerBox.get()
                            Spec = SpecBox.get().upper()
    
                            DES = int(DESBox.get())
                            SPA = int(SPABox.get())
                            OH = int(OHBox.get())
                            REQ = int(REQBox.get())
                            PCH = int(PCHBox.get())
                            BAL = int(BALBox.get())
                            RCV = int(RCVBox.get())
                            OS = int(OSBox.get())
                            Remark = RemarkBox.get()
                            Vendor = VendorBox.get()
                            
                            UnitCost = None
                            TotalUnitCost = sumTotalCost(UnitCost, REQ)
                            Currency = "SGD"
                            ExchangeRate = 1.0
                            TotalSGD = totalSGDConvert(TotalUnitCost, ExchangeRate)
                            
                            values = (PartNum, Description, D, CLS, V, Maker, Spec,
                                      DES, SPA, OH, REQ, PCH, BAL, RCV, OS, Remark,
                                      Vendor, UnitCost, TotalUnitCost, Currency,
                                      ExchangeRate, TotalSGD)
                        
                            curLoad = connLoad.cursor()
                            curLoad.execute(sqlCommand, values)
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Creation Successful",
                                                f"New Part {PartNum} Added", parent=frameUnit)
                        else:
                            pass                        

                def createUnitCost():
                    assemRef = assemFullName
                    
                    createUnitFormat = True
                    try:
                        int(PartNumBox.get())
                    except:
                        createUnitFormat = False
                    
                    if PartNumBox.get() == "":
                        messagebox.showerror("Unable to Create Unit",
                                             "Please Key In a Part Number",
                                             parent=frameUnit)
                        
                    elif createUnitFormat == False:
                        messagebox.showerror("Incorrect Part Number Format",
                                             "Please Key in an Integer",
                                             parent=frameUnit)
                    
                    else:
                        respCreateUnit = messagebox.askokcancel("Confirmation",
                                                                "Add New Part?",
                                                                parent=frameUnit)
                        if respCreateUnit == True:
                            sqlCommand = f"""INSERT INTO `{assemRef}` (
                            PartNum, Description, D, CLS, V,
                            Maker, Spec, DES, SPA, OH, REQ,
                            PCH, BAL, RCV, OS, REMARK, Vendor, UnitCost, 
                            TotalUnitCost, Currency, ExchangeRate, TotalSGD)
                    
                            VALUES (%s, %s, %s, %s, %s, %s, %s, 
                                    %s, %s, %s, %s, %s, %s, %s, 
                                    %s, %s, %s, %s, %s, %s, %s, %s)"""
                            
                            def sumTotalCost(unit, num):
                                if unit == None or unit == "":
                                    return 0
                                else:
                                    return unit * num
        
                            def checkCost(cost):
                                if cost == "":
                                    return None
                                else:
                                    return float(cost)
         
                            def totalSGDConvert(totalVal, exRate):
                                if exRate == None or exRate == "":
                                    return 0
                                else:
                                    totalSGDVal = float(totalVal) * float(exRate)
                                    return round(totalSGDVal, 2)
                            
                            DLst = ["", "D"]
                            
                            PartNum = str(PartNumBox.get()).rjust(3,"0")
                            Description = DescriptionBox.get().upper()
                            D = DLst[DBox.current()]
                            CLS = CLSBox.get()
                            V = VBox.get()
                            Maker = MakerBox.get()
                            Spec = SpecBox.get().upper()
    
                            DES = int(DESBox.get())
                            SPA = int(SPABox.get())
                            OH = int(OHBox.get())
                            REQ = int(REQBox.get())
                            PCH = int(PCHBox.get())
                            BAL = int(BALBox.get())
                            RCV = int(RCVBox.get())
                            OS = int(OSBox.get())
                            Remark = RemarkBox.get()
                            Vendor = VendorBox.get()
                            
                            UnitCost = checkCost(UnitCostBox.get())
                            TotalUnitCost = sumTotalCost(UnitCost, REQ)
                            Currency = CurrencyUnitBox.get()
                            ExchangeRate = ExRateBox.get()
                            TotalSGD = totalSGDConvert(TotalUnitCost, ExchangeRate)
                            
                            values = (PartNum, Description, D, CLS, V, Maker, Spec,
                                      DES, SPA, OH, REQ, PCH, BAL, RCV, OS, Remark,
                                      Vendor, UnitCost, TotalUnitCost, Currency,
                                      ExchangeRate, TotalSGD)
                        
                            curLoad = connLoad.cursor()
                            curLoad.execute(sqlCommand, values)
                            connLoad.commit()
                            curLoad.close()
                            
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Creation Successful",
                                                f"New Part {PartNum} Added", parent=frameUnit)
                        else:
                            pass
                
                def deleteUnit():
                    assemRef = assemFullName
                    selected = UnitTreeView.selection()
                    if selected == ():
                        messagebox.showwarning("Error", "Please Select a Value",
                                               parent=frameUnit)
                    else:
                        unitNumRef = UnitTreeView.item(selected[0])["values"][0]
                        unitNumDigit = (str(unitNumRef)).rjust(3,"0")
                        
                        respDelUnit = messagebox.askokcancel("Confirmation",
                                                             f"Delete Part {unitNumDigit}?",
                                                             parent=frameUnit)
                        if respDelUnit == True:
                            curLoad = connLoad.cursor()
                            for i in selected:
                                curLoad.execute(f"DELETE from `{assemRef}` WHERE oid={i}")
                                
                            connLoad.commit()
                            curLoad.close() 
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Successful",
                                                f"You Have Deleted Part {unitNumDigit}",
                                                parent=frameUnit)
                        else:
                            pass
                
                def selectUnit():
                    assemRef = assemFullName
                    selected = UnitTreeView.selection()
                    
                    if selected == ():
                        messagebox.showerror("No Item Selected", "Please Select an Item",
                                              parent=frameUnit)
                        clearEntryUnit()
                    else:
                        selected = selected[0]
                        clearEntryUnitNoZero()
                
                        sqlSelect = f"SELECT * FROM `{assemRef}` WHERE oid = %s"
                        valSelect = (selected, )
                        
                        curLoad = connLoad.cursor()
                        curLoad.execute(sqlSelect, valSelect)
                        resultSelect = curLoad.fetchall()
                        connLoad.commit()
                        curLoad.close()
                        
                        PartNumBox.insert(0, resultSelect[0][1])
                        DescriptionBox.insert(0, resultSelect[0][2])
                        
                        DBox.config(state="normal")
                        DBox.delete(0, END)
                        DBox.insert(0, resultSelect[0][3])
                        DBox.config(state="readonly")
                        
                        CLSBox.config(state="normal")
                        CLSBox.delete(0, END)
                        CLSBox.insert(0, resultSelect[0][4])
                        CLSBox.config(state="readonly")
                        
                        VBox.config(state="normal")
                        VBox.delete(0, END)
                        VBox.insert(0, resultSelect[0][5])
                        VBox.config(state="readonly")
                        
                        MakerBox.insert(0, resultSelect[0][6])
                        SpecBox.insert(0, resultSelect[0][7])
                        DESBox.insert(0, resultSelect[0][8])
                        SPABox.insert(0, resultSelect[0][9])
                        OHBox.insert(0, resultSelect[0][10])
                        REQBox.insert(0, resultSelect[0][11])
                        PCHBox.insert(0, resultSelect[0][12])
                        BALBox.insert(0, resultSelect[0][13])
                        RCVBox.insert(0, resultSelect[0][14])
                        OSBox.insert(0, resultSelect[0][15])
                        
                        RemarkBox.config(state="normal")
                        RemarkBox.delete(0, END)
                        RemarkBox.insert(0, resultSelect[0][16])
                        if resultSelect[0][16] == "On Hold" or resultSelect[0][16] == "":
                            RemarkBox.config(state="readonly")
                        
                        VendorBox.config(state="normal")
                        VendorBox.delete(0, END)
                        VendorBox.insert(0, resultSelect[0][17] if resultSelect[0][17] else "")
                        VendorBox.config(state="readonly")
                        
                        if AuthLevel(Login.AUTHLVL, 0) == False:
                            pass
                        else:
                            if resultSelect[0][18] == None:
                                UnitCostBox.insert(0, "")
                            else:
                                UnitCostBox.insert(0, resultSelect[0][18])
                        
                        if AuthLevel(Login.AUTHLVL, 0) == False:
                            pass
                        else:
                            UnitCostCcyLabel.config(text=str(resultSelect[0][20]))
                            
                            CurrencyUnitBox.config(state="normal")
                            CurrencyUnitBox.delete(0, END)
                            CurrencyUnitBox.insert(0, resultSelect[0][20])
                            CurrencyUnitBox.config(state="readonly")
                            
                            ExRateBox.config(state="normal")
                            ExRateBox.delete(0, END)
                            ExRateBox.insert(0, resultSelect[0][21])
                            ExRateBox.config(state="readonly")
                            
                            TotalUnitCostBox.config(state="normal")
                            TotalUnitCostBox.delete(0, END)
                            TotalUnitCostBox.insert(0, resultSelect[0][22])
                            TotalUnitCostBox.config(state="readonly")
                        
                        buttonUpdateUnit.config(state=NORMAL)
                        buttonCreateUnit.config(state=DISABLED)
                        buttonDeleteUnit.config(state=NORMAL)
                        UnitTreeView.config(selectmode="none")
            
                def deselectUnit():
                    selected = UnitTreeView.selection()
                    if len(selected) > 0:
                        UnitTreeView.selection_remove(selected[0])
                        clearEntryUnit()
                    else:
                        clearEntryUnit()
                
                def clearEntryUnit():
                    buttonUpdateUnit.config(state=DISABLED)
                    buttonCreateUnit.config(state=NORMAL)
                    buttonDeleteUnit.config(state=DISABLED)
                    
                    PartNumBox.delete(0, END)
                    DescriptionBox.delete(0, END)
                    DBox.current(0)
                    
                    VBox.config(state="normal")
                    VBox.delete(0, END)
                    VBox.insert(0, 1)
                    VBox.config(state="readonly")
                    
                    MakerBox.delete(0, END)
                    SpecBox.delete(0, END)
                    CLSBox.current(0)
                    RemarkBox.config(state="readonly")
                    RemarkBox.current(0)
                    
                    DESBox.delete(0, END)
                    DESBox.insert(0, 0)
                    SPABox.delete(0, END)
                    SPABox.insert(0, 0)
                    OHBox.delete(0, END)
                    OHBox.insert(0, 0)
                    REQBox.delete(0, END)
                    REQBox.insert(0, 0)
                    PCHBox.delete(0, END)
                    PCHBox.insert(0, 0)
                    BALBox.delete(0, END)
                    BALBox.insert(0, 0)
                    RCVBox.delete(0, END)
                    RCVBox.insert(0, 0)
                    OSBox.delete(0, END)
                    OSBox.insert(0, 0)
                    
                    VendorBox.config(state="normal")
                    VendorBox.delete(0, END)
                    VendorBox.config(state="readonly")
                    
                    UnitCostBox.delete(0, END)
                    UnitCostCcyLabel.config(text="SGD")
                    
                    TotalUnitCostBox.config(state="normal")
                    TotalUnitCostBox.delete(0, END)
                    TotalUnitCostBox.config(state="readonly")
                    
                    CurrencyUnitBox.config(state="readonly")
                    CurrencyUnitBox.current(0)
                    
                    ExRateBox.config(state="normal")
                    ExRateBox.delete(0, END)
                    ExRateBox.insert(0, 1.00)
                    ExRateBox.config(state="readonly")
                    UnitTreeView.config(selectmode="browse")

                def clearEntryUnitNoZero():
                    buttonUpdateUnit.config(state=DISABLED)
                    buttonCreateUnit.config(state=NORMAL)
                    buttonDeleteUnit.config(state=DISABLED)
                    
                    PartNumBox.delete(0, END)
                    DescriptionBox.delete(0, END)
                    DBox.current(0)
                    
                    VBox.config(state="normal")
                    VBox.delete(0, END)
                    VBox.insert(0, 1)
                    VBox.config(state="readonly")
                    
                    MakerBox.delete(0, END)
                    SpecBox.delete(0, END)
                    CLSBox.current(0)
                    RemarkBox.config(state="readonly")
                    RemarkBox.current(0)
                    
                    DESBox.delete(0, END)
                    SPABox.delete(0, END)
                    OHBox.delete(0, END)
                    REQBox.delete(0, END)
                    PCHBox.delete(0, END)
                    BALBox.delete(0, END)
                    RCVBox.delete(0, END)
                    OSBox.delete(0, END)
                    
                    VendorBox.config(state="normal")
                    VendorBox.delete(0, END)
                    VendorBox.config(state="readonly")
                    
                    UnitCostBox.delete(0, END)
                    UnitCostCcyLabel.config(text="SGD")
                    
                    TotalUnitCostBox.config(state="normal")
                    TotalUnitCostBox.delete(0, END)
                    TotalUnitCostBox.config(state="readonly")
                    
                    CurrencyUnitBox.config(state="readonly")
                    CurrencyUnitBox.current(0)
                    
                    ExRateBox.config(state="normal")
                    ExRateBox.delete(0, END)
                    ExRateBox.insert(0, 1.00)
                    ExRateBox.config(state="readonly")
                    UnitTreeView.config(selectmode="browse")

                def removeAllUnit():
                    respA = messagebox.askokcancel("Confirmation", "Confirm Delete All Parts")
                    if respA == True:
                        respB = messagebox.askokcancel("Warning", "This Action is NOT Recoverable")
                        if respB == True:
                            assemRef = assemFullName
                            selectAll = UnitTreeView.get_children()
                            curLoad = connLoad.cursor()
                            for i in selectAll:
                                curLoad.execute(f"DELETE from `{assemRef}` WHERE oid={i}")   
                                connLoad.commit()
                                
                            curLoad.close() 
                            clearEntryUnit()
                            UnitTreeView.delete(*UnitTreeView.get_children())
                            queryTreeUnit()
                            
                            checkAssemTotal()
                            checkMachTotal()
                            checkProTotal()
                            
                            messagebox.showinfo("Delete Successful", "You Have Deleted ALL Parts of This BOM")
                        else:
                            pass
                    else:
                        pass
                
                def fetchCcy():
                    currencyLst = CountryRef.getCcyLst()
                    if currencyLst == []:
                        return ["Not Found"]
                    else:
                        return currencyLst
                
                def refreshUnit():
                    CLSBox.config(value=fetchUnitClass())
                    clearEntryUnit()
                    UnitTreeView.delete(*UnitTreeView.get_children())
                    queryTreeUnit()
                    CurrencyUnitBox.config(value=fetchCcy())
                    CurrencyUnitBox.current(0)
                    
                def checkPartNum(partNum):
                    val = partNum
                    if len(val)>=3:
                        return val[-3:]
                    else:
                        return val
                
                def importCSVUnit():
                    assemRef = assemFullName
                    fileDir = filedialog.askopenfilename(initialdir="~/Desktop",
                                                          title="Select A File",
                                                          filetypes=(("CSV Files", "*.csv"),
                                                                    ("Any Files", "*.*")))
                    
                    if fileDir == "":
                        messagebox.showwarning("Please Select a File",
                                               "No File Selected",
                                               parent=frameUnit)
                    
                    else:
                        with open(f"{fileDir}", encoding="utf-8-sig") as f: 
                            csvFile = csv.reader(f, delimiter=",")
                            
                            rawLst = []
                            for row in csvFile:
                                rawLst.append(row)
                            
                            fullLst = []
                            for i in range(1, len(rawLst)):
                                singleLst = ["", "", "", "", "", "", 
                                              "", "", "", "", "", "", 
                                              "", "", "", "", "", "",
                                              "", "", "", ""]
                                for j in range(len(rawLst[i])):
                                    if rawLst[0][j] == "PART NO":
                                        singleLst[0] = checkPartNum(rawLst[i][j])
                                    elif rawLst[0][j] == "DESCRIPTION":
                                        singleLst[1] = rawLst[i][j].upper()
                                    elif rawLst[0][j] == "D":
                                        singleLst[2] = rawLst[i][j]                    
                                    elif rawLst[0][j] == "CLASS":
                                        singleLst[3] = rawLst[i][j]
                                    elif rawLst[0][j] == "V":
                                        singleLst[4] = rawLst[i][j]
                                    elif rawLst[0][j] == "MAKER":
                                        singleLst[5] = rawLst[i][j]
                                    elif rawLst[0][j] == "MAKER SPEC":
                                        singleLst[6] = rawLst[i][j].upper()
                                    elif rawLst[0][j] == "D. QTY":
                                        singleLst[7] = rawLst[i][j]
                                    elif rawLst[0][j] == "S. QTY":
                                        singleLst[8] = rawLst[i][j]
                                    elif rawLst[0][j] == "OH. QTY":
                                        singleLst[9] = rawLst[i][j]
                                    elif rawLst[0][j] == "REQ. QTY":
                                        singleLst[10] = rawLst[i][j]
                                    elif rawLst[0][j] == "PCH. QTY":
                                        singleLst[11] = rawLst[i][j]
                                    elif rawLst[0][j] == "BAL. QTY":
                                        singleLst[12] = rawLst[i][j]
                                    elif rawLst[0][j] == "RCV. QTY":
                                        singleLst[13] = rawLst[i][j]
                                    elif rawLst[0][j] == "OS. QTY":
                                        singleLst[14] = rawLst[i][j]
                                    elif rawLst[0][j] == "REMARK":
                                        singleLst[15] = rawLst[i][j]
                                    elif rawLst[0][j] == "VENDOR":
                                        singleLst[16] = rawLst[i][j]
                                    elif rawLst[0][j] == "UNIT COST":
                                        singleLst[17] = rawLst[i][j]
    
                                    elif rawLst[0][j] == "TOTAL COST":
                                        singleLst[18] = rawLst[i][j]
                                    elif rawLst[0][j] == "CURRENCY":
                                        singleLst[19] = rawLst[i][j]
                                    elif rawLst[0][j] == "EXCHANGE RATE":
                                        singleLst[20] = rawLst[i][j]
                                    elif rawLst[0][j] == "TOTAL SGD":
                                        singleLst[21] = rawLst[i][j]
                                        
                                if singleLst != ["", "", "", "", "", "", 
                                                  "", "", "", "", "", "", 
                                                  "", "", "", "", "", "",
                                                  "", "", "", ""]:
                                    fullLst.append(singleLst)
                        
                        calcLst = []
                        DesQtyVal = int(desQty)
                        OrderQtyVal = int(orderQty)
                        
                        def removeSymbol(cost):
                            val = re.sub("[^0-9\.]", "", str(cost))
                            return val
                        
                        def sumTotalCost(unitCst, num):
                            if unitCst == None or unitCst == "":
                                return 0
                            else:
                                val = unitCst * num
                                return round(val, 2)
     
                        def totalSGDConvert(totalVal, exRate):
                            if exRate == None or exRate == "":
                                return 0
                            else:
                                totalSGDVal = float(totalVal) * float(exRate)
                                return round(totalSGDVal, 2)
                        
                        for i in range(len(fullLst)):
                            PartNum = (fullLst[i][0]).rjust(3,"0")
                            PartDesc = fullLst[i][1]
                            D = fullLst[i][2]
                            CLS = fullLst[i][3]
                            
                            if fullLst[i][4] == "":
                                V = 1
                            else:
                                V = int(fullLst[i][4])
                                
                            Maker = fullLst[i][5]
                            MakerSpec = fullLst[i][6]
                            
                            if fullLst[i][7] == "":
                                DES = 0
                            else:
                                DES = int(fullLst[i][7])
                            if fullLst[i][8] == "":
                                SPA = 0
                            else:
                                SPA = int(fullLst[i][8])                            
                            if fullLst[i][9] == "":
                                OH = 0
                            else:
                                maxOH = (DES + SPA) * DesQtyVal * OrderQtyVal
                                if int(fullLst[i][9]) > maxOH:
                                    OH = maxOH
                                else:
                                    OH = int(fullLst[i][9])
                            
                            if fullLst[i][10] == "":
                                REQ = ((DES + SPA) * DesQtyVal * OrderQtyVal) - OH
                            else:
                                REQ = int(fullLst[i][10])
                            if fullLst[i][11] == "":
                                PCH = 0
                            else:
                                PCH = int(fullLst[i][11])
                            if fullLst[i][12] == "":
                                BAL = REQ-PCH
                            else:
                                BAL = int(fullLst[i][12])
                                
                            if fullLst[i][13] == "":
                                RCV = 0
                            else:
                                RCV = int(fullLst[i][13])
                            if fullLst[i][14] == "":
                                OS = PCH-RCV
                            else:
                                OS = int(fullLst[i][14])
                            
                            Remark = fullLst[i][15]
                            Vendor = fullLst[i][16]
                            
                            if fullLst[i][17] == "":
                                UnitCost = None
                            else:
                                UnitCost = float(removeSymbol(fullLst[i][17]))
                            
                            if fullLst[i][18] == "":
                                TotalUnitCost = sumTotalCost(UnitCost, REQ)
                            else:
                                TotalUnitCost = float(fullLst[i][18])
                                
                            if fullLst[i][19] == "":
                                Currency = "SGD"
                            else:
                                Currency = fullLst[i][19]
                                
                            if fullLst[i][20] == "":
                                if Currency == "SGD":
                                    ExRate = 1.00
                                elif Currency in CountryRef.getCcyLst():
                                    ExRate = CountryRef.getExRate(Currency)
                                else:
                                    ExRate = 0.00
                            else:
                                ExRate = float(fullLst[i][20])
                                
                            if fullLst[i][21] == "":
                                TotalSGD = totalSGDConvert(TotalUnitCost, ExRate)
                            else:
                                TotalSGD = fullLst[i][21]
                            
                            tup = (PartNum, PartDesc, D, CLS, V, Maker, MakerSpec,
                                    DES, SPA, OH, REQ, PCH, BAL, RCV, OS, Remark, Vendor,
                                    UnitCost, TotalUnitCost, Currency, ExRate, TotalSGD)
                            
                            calcLst.append(tup)
    
                        sqlCommand = f""" REPLACE INTO `{assemRef}` (
                        PartNum, Description, D, CLS, V,
                        Maker, Spec, DES, SPA, OH, REQ,
                        PCH, BAL, RCV, OS, Remark, Vendor, UnitCost,
                        TotalUnitCost, Currency, ExchangeRate, TotalSGD)
                    
                        VALUES (%s, %s, %s, %s, %s, %s, 
                                %s, %s, %s, %s, %s, %s,
                                %s, %s, %s, %s, %s, %s,
                                %s, %s, %s, %s)"""
                    
                        curLoad = connLoad.cursor()
                        for i in range(len(calcLst)):
                            curLoad.execute(sqlCommand, calcLst[i])
                            connLoad.commit()
    
                        curLoad.close()
                        clearEntryUnit()
                        UnitTreeView.delete(*UnitTreeView.get_children())
                        queryTreeUnit()
                        
                        checkAssemTotal()
                        checkMachTotal()
                        checkProTotal()
                        
                        messagebox.showinfo("Import Successful", 
                                            "You Have Imported BOM", parent=frameUnit) 
                
                def exportCSVUnit():
                    proRef = proName
                    machRef = machName
                    assemTypeRef = assemType
                    assemNameRef = assemName
                    assemRef = assemFullName
                    
                    if AuthLevel(Login.AUTHLVL, 1) == False:
                        messagebox.showerror("Insufficient Clearance",
                                             "You are NOT authorized to export", parent=frameUnit)
                    
                    else:
                        respExportUnit = messagebox.askokcancel("Confirmation", "Confirm Export as CSV?")
                        if respExportUnit == True:
                            curLoad = connLoad.cursor()
                            curLoad.execute(f"SELECT * FROM {assemRef}")
                            result = curLoad.fetchall()
                            connLoad.commit()
                            curLoad.close()
                            
                            unitFull = f"{proName}-{machRef}-{assemTypeRef}{assemNameRef}"
                            
                            headingLst = ["PART NO", "DESCRIPTION", "D", "CLASS",
                                          "V", "MAKER", "MAKER SPEC", 
                                          "D. QTY", "S. QTY", "OH. QTY", "REQ. QTY",
                                          "PCH. QTY", "BAL. QTY", "RCV. QTY", "OS. QTY",
                                          "Remark", "Vendor", "Unit Cost", "Total Cost", "Total in SGD"]
                            
                            def threeDigitConverter(val):
                                return val.rjust(3,"0")
                            
                            def checkCurrencyNone(num, ccy):
                                if num == None or num == "":
                                    return num
                                else:
                                    twoDigit = "{:.2f}".format(num)
                                    numCcy = f"{twoDigit} {ccy}"
                                    return numCcy
                            
                            fullLst = []
                            for i in range(len(result)):
                                singleLst = [f"{unitFull}-{threeDigitConverter(result[i][1])}", 
                                             result[i][2], result[i][3], result[i][4],
                                             result[i][5], result[i][6], result[i][7],
                                             result[i][8], result[i][9], result[i][10],
                                             result[i][11], result[i][12], result[i][13],
                                             result[i][14], result[i][15], result[i][16],
                                             result[i][17], 
                                             checkCurrencyNone(result[i][18], result[i][20]),
                                             checkCurrencyNone(result[i][19], result[i][20]),
                                             checkCurrencyNone(result[i][22], "SGD")]
                                fullLst.append(singleLst)
                            
                            with open(f"{unitFull}.csv", "w", newline="") as f:
                                theWriter = csv.writer(f)
                                theWriter.writerow(headingLst)
                                for rec in fullLst:
                                    theWriter.writerow(rec)
    
                            messagebox.showinfo("Export Successful", 
                                                f"You Have Successfuly Exported {unitFull}", parent=frameAssem) 
                
                def closeTabUnit():
                    respCloseUnit = messagebox.askokcancel("Confirmation",
                                                           "Exit BOM Tab Now?",
                                                           parent=frameUnit)
                    if respCloseUnit == True:
                        frameUnit.destroy()
                        buttonLoadAssem.config(state=NORMAL)
                    else:
                        pass
            
                #Vendor Selection Window
                def SelectVendor():
                    connVend = mysql.connector.connect(host = logininfo[0],
                                                       user = logininfo[1], 
                                                       password =logininfo[2],
                                                       database = "INDEX_VEND_MASTER")
                    
                    curVend = connVend.cursor()
                    curVend.execute(f"""SELECT * FROM VENDOR_LIST 
                                    WHERE `CLASS` = '{CLSBox.get()}' """)
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
                            
                    
                    global img
                    #photo = Image.open('mag3.png')
                    photo = PhotoImage(file = r"mag3.png")
                    img = photo.subsample(4, 4)
                    SearchButton = Button(SearchFrame,command = Search, image = img)
                    SearchButton.grid(row=0, column = 2)

                    def SelectedVend():
                        try:
                            VendorBox.config(state="normal")
                            VendorBox.delete(0, END)  
                            VendorBox.insert(END,curVEND[0])
                            VendorBox.config(state="readonly")
                            VendWin.destroy()
                        except:
                            VendorBox.config(state="normal")
                            VendorBox.delete(0, END)  
                            VendorBox.config(state="readonly")
                            messagebox.showwarning("No Vendor Selected", "Please Select a Vendor", 
                                                   parent=VendWin)
                    
                    def ClearVend():
                        VendorBox.config(state="normal")
                        VendorBox.delete(0, END)
                        VendorBox.config(state="readonly")
                        VendWin.destroy()
                    
                    def BackVend():
                        VendWin.destroy()
                
                    SelButton = Button(VendWin, text ="Select", command = SelectedVend, width = 10)
                    SelButton.grid(row = 2, column = 0)
                    
                    ClearButton = Button(VendWin, text ="Clear", command = ClearVend, width = 10)
                    ClearButton.grid(row = 2, column = 1)
                    
                    ExitButton = Button(VendWin, text ="Exit", command = BackVend, width = 10)
                    ExitButton.grid(row = 2, column = 2)
  

                
                def NumBoxClick(e):
                    NumBoxCom()
                    
                def NumBoxCom():
                    DESVal = int(DESBox.get())
                    SPAVal = int(SPABox.get())
                    DesQtyVal = int(desQty)
                    OrderQtyVal = int(orderQty)
                    
                    maxOHVal = (DESVal + SPAVal) * DesQtyVal * OrderQtyVal
                    OHBox.config(to=maxOHVal)
                    OHVal = int(OHBox.get())
                    if OHVal > maxOHVal:
                        OHBox.delete(0, END)
                        OHBox.insert(0, maxOHVal)
                        OHVal = maxOHVal
                    
                    REQVal = ((DESVal + SPAVal) * DesQtyVal * OrderQtyVal) - OHVal
                    
                    maxPCHVal = REQVal
                    PCHBox.config(to=maxPCHVal)
                    PCHVal = int(PCHBox.get())
                    if PCHVal > maxPCHVal:
                        PCHBox.delete(0, END)
                        PCHBox.insert(0, maxPCHVal)
                        PCHVal = maxPCHVal
                    
                    maxRCVVal = PCHVal
                    RCVBox.config(to=maxRCVVal)
                    RCVVal = int(RCVBox.get())
                    if RCVVal > maxRCVVal:
                        RCVBox.delete(0, END)
                        RCVBox.insert(0, maxRCVVal)
                        RCVVal = maxRCVVal
                    
                    BALVal = REQVal - PCHVal
                    OSVal = PCHVal - RCVVal
                    
                    if REQVal < 0:
                        REQVal = 0
                    if BALVal < 0:
                        REQVal = 0
                    if OSVal < 0:
                        REQVal = 0
                    
                    REQBox.delete(0, END)
                    REQBox.insert(0, REQVal)
                    BALBox.delete(0, END)
                    BALBox.insert(0, BALVal)
                    OSBox.delete(0, END)
                    OSBox.insert(0, OSVal)
                
                def CurrencyUnitSelect(e):
                    exRate = CountryRef.getExRate(CurrencyUnitBox.get())
                    ExRateBox.config(state="normal")
                    ExRateBox.delete(0, END)
                    ExRateBox.insert(0, exRate)
                    ExRateBox.config(state="readonly")
                    UnitCostCcyLabel.config(text=str(CurrencyUnitBox.get()))
                
                def RemarkSelect(e):
                    if RemarkBox.get() == "Received":
                        RemarkBox.config(state="normal")
                        RemarkBox.delete(0, END)
                        timeNow = datetime.now()
                        formatDate = timeNow.strftime("%d-%b-%y")
                        RemarkBox.insert(0, f"Rcv. {formatDate}")
                    elif RemarkBox.get() == "Others":
                        RemarkBox.config(state="normal")
                        RemarkBox.delete(0, END)
                    else:
                        RemarkBox.config(state="readonly")
                    
                def setPCHAll(e):
                    ReqQty = REQBox.get()
                    PCHBox.delete(0, END)
                    PCHBox.insert(0, ReqQty)
                    NumBoxCom()
                    
                def setRCVAll(e):
                    PchQty = PCHBox.get()
                    RCVBox.delete(0, END)
                    RCVBox.insert(0, PchQty)
                    NumBoxCom()





                UnitDataFrame = LabelFrame(frameUnit, text="Record")
                UnitDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
                # UnitDataFrame.pack(fill="x", expand="yes", padx=20)
                
                UnitButtonFrame = LabelFrame(frameUnit, text="Command")
                UnitButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
                # UnitButtonFrame.pack(fill="x", expand="yes", padx=20)
    
                PartNumLabel = Label(UnitDataFrame, text="Part No.")
                DescriptionLabel = Label(UnitDataFrame, text="Description")
                DLabel = Label(UnitDataFrame, text="D")
                VLabel = Label(UnitDataFrame, text="V")
                MakerLabel = Label(UnitDataFrame, text="Maker")
                SpecLabel = Label(UnitDataFrame, text="Specification")
                CLSLabel = Label(UnitDataFrame, text="Class")
                RemarkLabel = Label(UnitDataFrame, text="Remark")
                DESLabel = Label(UnitDataFrame, text="Design Qty.")
                SPALabel = Label(UnitDataFrame, text="Spare Qty.")
                OHLabel = Label(UnitDataFrame, text="On Hand Qty.")
                REQLabel = Label(UnitDataFrame, text="Required Qty.")
                PCHLabel = Label(UnitDataFrame, text="Purchased Qty.")
                BALLabel = Label(UnitDataFrame, text="Balance Qty.")
                RCVLabel = Label(UnitDataFrame, text="Received Qty.")
                OSLabel = Label(UnitDataFrame, text="Outstanding Qty.")
                VendorLabel = Label(UnitDataFrame, text="Vendor")
                UnitCostLabel = Label(UnitDataFrame, text="Unit Cost")
                TotalUnitCostLabel = Label(UnitDataFrame, text="Total Cost")
                CurrencyUnitLabel = Label(UnitDataFrame, text="Currency")
            
                PartNumLabel.grid(row=0, column=0, padx=10, sticky=E)
                DescriptionLabel.grid(row=1, column=0, padx=10, sticky=E)
                DLabel.grid(row=2, column=0, padx=10, sticky=E)
                VLabel.grid(row=3, column=0, padx=10, sticky=E)
                MakerLabel.grid(row=0, column=2, padx=10, sticky=E)
                SpecLabel.grid(row=1, column=2, padx=10, sticky=E)
                CLSLabel.grid(row=2, column=2, padx=10, sticky=E)
                RemarkLabel.grid(row=3, column=2, padx=10, sticky=E)
                DESLabel.grid(row=0, column=4, padx=10, sticky=E)
                SPALabel.grid(row=1, column=4, padx=10, sticky=E)
                OHLabel.grid(row=2, column=4, padx=10, sticky=E)
                REQLabel.grid(row=3, column=4, padx=10, sticky=E)
                PCHLabel.grid(row=0, column=6, padx=10, sticky=E)
                BALLabel.grid(row=1, column=6, padx=10, sticky=E)
                RCVLabel.grid(row=2, column=6, padx=10, sticky=E)
                OSLabel.grid(row=3, column=6, padx=10, sticky=E)
                VendorLabel.grid(row=0, column=8, padx=10, sticky=E)
                UnitCostLabel.grid(row=1, column=8, padx=10, sticky=E)
                TotalUnitCostLabel.grid(row=2, column=8, padx=10, sticky=E)
                CurrencyUnitLabel.grid(row=3, column=8, padx=10, sticky=E)
                
                PartNumBox = Entry(UnitDataFrame)
                DescriptionBox = Entry(UnitDataFrame)
                DBox = ttk.Combobox(UnitDataFrame, width=8, 
                                      value=["", "Deleted"], 
                                      state="readonly")
                VBox = Spinbox(UnitDataFrame, width=8, from_=0, to=1000, state="readonly")
                MakerBox = Entry(UnitDataFrame, width=30)
                SpecBox = Entry(UnitDataFrame, width=30)
                CLSBox = ttk.Combobox(UnitDataFrame, width=20, 
                                      value=fetchUnitClass(), state="readonly")
                RemarkBox = ttk.Combobox(UnitDataFrame, width=20, 
                                      value=["", "On Hold", "Received", "Others"], 
                                      state="readonly")
                DESBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                SPABox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                OHBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                REQBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                PCHBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                BALBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                RCVBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                OSBox = Spinbox(UnitDataFrame, from_=0, to=10000, width=8)
                VendorBox = Entry(UnitDataFrame, width=18, state="readonly")
                
                UnitCostBox = Spinbox(UnitDataFrame, from_=0, to=1000, increment=0.01, width=12)
                UnitCostCcyLabel = Label(UnitDataFrame, text="SGD")
                UnitCostCcyLabel.grid(row=1, column=11, padx=0, sticky=W)
                
                TotalUnitCostBox = Entry(UnitDataFrame, width=14, state="readonly")
                TotalUnitCostCcyLabel = Label(UnitDataFrame, text="SGD")
                TotalUnitCostCcyLabel.grid(row=2, column=11, padx=0, sticky=W)
                
                CurrencyUnitBox = ttk.Combobox(UnitDataFrame, width=10, height=5,
                                               value=CountryRef.getCcyLst(), state="readonly")
                
                EqualLabel = Label(UnitDataFrame, text=" = ")
                EqualLabel.grid(row=3, column=10, padx=0, sticky=W)
                ExRateBox = Entry(UnitDataFrame, width=8, state="readonly")
                SGDLabel = Label(UnitDataFrame, text="SGD")
                SGDLabel.grid(row=3, column=12, padx=0, sticky=W)
                
                VendorSelectionButton = Button(UnitDataFrame,text = "List", command = SelectVendor, width=3, height=1)
                if AuthLevel(Login.AUTHLVL, 0) == False:
                    VendorSelectionButton.grid(row=0, column=13, pady=5, padx=(10,5), sticky=E)
                else:
                    VendorSelectionButton.grid(row=0, column=11, pady=5, padx=(10,5), sticky=E)
                
                VBox.config(state="normal")
                VBox.delete(0, END)
                VBox.insert(0, 1)
                VBox.config(state="readonly")
                
                CLSBox.current(0)
                UnitCostBox.delete(0, END)
                CurrencyUnitBox.current(0)
                RemarkBox.current(0)
                
                ExRateBox.config(state="normal")
                ExRateBox.delete(0, END)
                ExRateBox.insert(0, 1.00)
                ExRateBox.config(state="readonly")
                
                PartNumBox.grid(row=0, column=1, pady=5, sticky=W)
                DescriptionBox.grid(row=1, column=1, pady=5, sticky=W)
                DBox.grid(row=2, column=1, pady=5, sticky=W)
                VBox.grid(row=3, column=1, pady=5, sticky=W)
                MakerBox.grid(row=0, column=3, pady=5, sticky=W)
                SpecBox.grid(row=1, column=3, pady=5, sticky=W)
                CLSBox.grid(row=2, column=3, pady=5, sticky=W)
                RemarkBox.grid(row=3, column=3, columnspan=2, pady=5, sticky=W)
                DESBox.grid(row=0, column=5, pady=5, sticky=W)
                SPABox.grid(row=1, column=5, pady=5, sticky=W)
                OHBox.grid(row=2, column=5, pady=5, sticky=W)
                REQBox.grid(row=3, column=5, pady=5, sticky=W)
                PCHBox.grid(row=0, column=7, pady=5, sticky=W)
                BALBox.grid(row=1, column=7, pady=5, sticky=W)
                RCVBox.grid(row=2, column=7, pady=5, sticky=W)
                OSBox.grid(row=3, column=7, pady=5, sticky=W)
                VendorBox.grid(row=0, column=9, columnspan=4, pady=5, sticky=W)
                UnitCostBox.grid(row=1, column=9, columnspan=4, pady=5, sticky=W)
                TotalUnitCostBox.grid(row=2, column=9, columnspan=4, pady=5, sticky=W)
                CurrencyUnitBox.grid(row=3, column=9, pady=5, sticky=W)
                ExRateBox.grid(row=3, column=11, padx=5, pady=5, sticky=W)
                
                if AuthLevel(Login.AUTHLVL, 0) == False:
                    UnitCostLabel.grid_forget()
                    TotalUnitCostLabel.grid_forget()
                    CurrencyUnitLabel.grid_forget()
                    UnitCostBox.grid_forget()
                    UnitCostCcyLabel.grid_forget()
                    TotalUnitCostBox.grid_forget()
                    TotalUnitCostCcyLabel.grid_forget()
                    CurrencyUnitBox.grid_forget()
                    EqualLabel.grid_forget()
                    ExRateBox.grid_forget()
                    SGDLabel.grid_forget()
                    
                    
                    
                buttonUpdateUnit = Button(UnitButtonFrame, text="Update Unit", command=updateUnit, state=DISABLED)
                buttonUpdateUnit.grid(row=0, column=0, padx=10, pady=10)
                
                buttonCreateUnit = Button(UnitButtonFrame, text="Create Unit", command=createUnit)
                buttonCreateUnit.grid(row=0, column=1, padx=10, pady=10)
            
                buttonDeleteUnit = Button(UnitButtonFrame, text="Delete Unit", command=deleteUnit, state=DISABLED)
                buttonDeleteUnit.grid(row=0, column=2, padx=10, pady=10)
            
                buttonSelectUnit = Button(UnitButtonFrame, text="Select Unit", command=selectUnit)
                buttonSelectUnit.grid(row=0, column=3, padx=10, pady=10)
                
                buttonDeselectUnit = Button(UnitButtonFrame, text="Deselect Unit", command=deselectUnit)
                buttonDeselectUnit.grid(row=0, column=4, padx=10, pady=10)
                
                buttonClearEntryUnit = Button(UnitButtonFrame, text="Clear Entry", command=clearEntryUnit)
                buttonClearEntryUnit.grid(row=0, column=5, padx=10, pady=10)
                
                buttonClearRemoveAllUnit = Button(UnitButtonFrame, text="Remove All Unit", command=removeAllUnit)
                buttonClearRemoveAllUnit.grid(row=0, column=6, padx=(90,10), pady=10)
                
                buttonRefreshUnit = Button(UnitButtonFrame, text="Refresh", command=refreshUnit)
                buttonRefreshUnit.grid(row=0, column=7, padx=10, pady=10)
                
                buttonImportCSVUnit = Button(UnitButtonFrame, text="Import CSV", command=importCSVUnit)
                buttonImportCSVUnit.grid(row=0, column=8, padx=10, pady=10)
                
                buttonExportCSVUnit = Button(UnitButtonFrame, text="Export CSV", command=exportCSVUnit)
                buttonExportCSVUnit.grid(row=0, column=9, padx=10, pady=10)
                
                buttonCloseTabUnit = Button(UnitButtonFrame, text="Close Tab", command=closeTabUnit)
                buttonCloseTabUnit.grid(row=0, column=10, padx=10, pady=10)
                
                if AuthLevel(Login.AUTHLVL, 1) == False:
                    buttonExportCSVUnit.grid_forget()
    
                DESBox.bind("<Enter>", NumBoxClick)
                DESBox.bind("<Leave>", NumBoxClick)
                SPABox.bind("<Enter>", NumBoxClick)
                SPABox.bind("<Leave>", NumBoxClick)
                OHBox.bind("<Enter>", NumBoxClick)
                OHBox.bind("<Leave>", NumBoxClick)
                PCHBox.bind("<Enter>", NumBoxClick)
                PCHBox.bind("<Leave>", NumBoxClick)
                RCVBox.bind("<Enter>", NumBoxClick)
                RCVBox.bind("<Leave>", NumBoxClick)
                
                PCHBox.bind("<Double-Button-3>", setPCHAll)
                RCVBox.bind("<Double-Button-3>", setRCVAll)
                
                RemarkBox.bind("<<ComboboxSelected>>", RemarkSelect)
                CurrencyUnitBox.bind("<<ComboboxSelected>>", CurrencyUnitSelect)
                
                
                
                
                
                PartNumBox.bind("<Return>", updateUnitReturn)
                PartNumBox.bind("<Delete>", deleteUnitDel)
                PartNumBox.bind("<Control-d>", setDeleteUnit)
                PartNumBox.bind("<Escape>", closeUnitTabEsc)
                
                DescriptionBox.bind("<Return>", updateUnitReturn)
                DescriptionBox.bind("<Delete>", deleteUnitDel)
                DescriptionBox.bind("<Control-d>", setDeleteUnit)
                DescriptionBox.bind("<Escape>", closeUnitTabEsc)
                
                DBox.bind("<Return>", updateUnitReturn)
                DBox.bind("<Delete>", deleteUnitDel)
                DBox.bind("<Control-d>", setDeleteUnit)
                DBox.bind("<Escape>", closeUnitTabEsc)
                
                VBox.bind("<Return>", updateUnitReturn)
                VBox.bind("<Delete>", deleteUnitDel)
                VBox.bind("<Control-d>", setDeleteUnit)
                VBox.bind("<Escape>", closeUnitTabEsc)
                
                MakerBox.bind("<Return>", updateUnitReturn)
                MakerBox.bind("<Delete>", deleteUnitDel)
                MakerBox.bind("<Control-d>", setDeleteUnit)
                MakerBox.bind("<Escape>", closeUnitTabEsc)
                
                SpecBox.bind("<Return>", updateUnitReturn)
                SpecBox.bind("<Delete>", deleteUnitDel)
                SpecBox.bind("<Control-d>", setDeleteUnit)
                SpecBox.bind("<Escape>", closeUnitTabEsc)
                
                CLSBox.bind("<Return>", updateUnitReturn)
                CLSBox.bind("<Delete>", deleteUnitDel)
                CLSBox.bind("<Control-d>", setDeleteUnit)
                CLSBox.bind("<Escape>", closeUnitTabEsc)
                
                RemarkBox.bind("<Return>", updateUnitReturn)
                RemarkBox.bind("<Delete>", deleteUnitDel)
                RemarkBox.bind("<Control-d>", setDeleteUnit)
                RemarkBox.bind("<Escape>", closeUnitTabEsc)
                
                DESBox.bind("<Return>", updateUnitReturn)
                DESBox.bind("<Delete>", deleteUnitDel)
                DESBox.bind("<Control-d>", setDeleteUnit)
                DESBox.bind("<Escape>", closeUnitTabEsc)
                
                SPABox.bind("<Return>", updateUnitReturn)
                SPABox.bind("<Delete>", deleteUnitDel)
                SPABox.bind("<Control-d>", setDeleteUnit)
                SPABox.bind("<Escape>", closeUnitTabEsc)
                
                OHBox.bind("<Return>", updateUnitReturn)
                OHBox.bind("<Delete>", deleteUnitDel)
                OHBox.bind("<Control-d>", setDeleteUnit)
                OHBox.bind("<Escape>", closeUnitTabEsc)
                
                REQBox.bind("<Return>", updateUnitReturn)
                REQBox.bind("<Delete>", deleteUnitDel)
                REQBox.bind("<Control-d>", setDeleteUnit)
                REQBox.bind("<Escape>", closeUnitTabEsc)
                
                PCHBox.bind("<Return>", updateUnitReturn)
                PCHBox.bind("<Delete>", deleteUnitDel)
                PCHBox.bind("<Control-d>", setDeleteUnit)
                PCHBox.bind("<Escape>", closeUnitTabEsc)
                
                BALBox.bind("<Return>", updateUnitReturn)
                BALBox.bind("<Delete>", deleteUnitDel)
                BALBox.bind("<Control-d>", setDeleteUnit)
                BALBox.bind("<Escape>", closeUnitTabEsc)
                
                RCVBox.bind("<Return>", updateUnitReturn)
                RCVBox.bind("<Delete>", deleteUnitDel)
                RCVBox.bind("<Control-d>", setDeleteUnit)
                RCVBox.bind("<Escape>", closeUnitTabEsc)
                
                OSBox.bind("<Return>", updateUnitReturn)
                OSBox.bind("<Delete>", deleteUnitDel)
                OSBox.bind("<Control-d>", setDeleteUnit)
                OSBox.bind("<Escape>", closeUnitTabEsc)
                
                VendorBox.bind("<Return>", updateUnitReturn)
                VendorBox.bind("<Delete>", deleteUnitDel)
                VendorBox.bind("<Control-d>", setDeleteUnit)
                VendorBox.bind("<Escape>", closeUnitTabEsc)
                
                UnitCostBox.bind("<Return>", updateUnitReturn)
                UnitCostBox.bind("<Delete>", deleteUnitDel)
                UnitCostBox.bind("<Control-d>", setDeleteUnit)
                UnitCostBox.bind("<Escape>", closeUnitTabEsc)
                
                TotalUnitCostBox.bind("<Return>", updateUnitReturn)
                TotalUnitCostBox.bind("<Delete>", deleteUnitDel)
                TotalUnitCostBox.bind("<Control-d>", setDeleteUnit)
                TotalUnitCostBox.bind("<Escape>", closeUnitTabEsc)
                
                CurrencyUnitBox.bind("<Return>", updateUnitReturn)
                CurrencyUnitBox.bind("<Delete>", deleteUnitDel)
                CurrencyUnitBox.bind("<Control-d>", setDeleteUnit)
                CurrencyUnitBox.bind("<Escape>", closeUnitTabEsc)
                
                ExRateBox.bind("<Return>", updateUnitReturn)
                ExRateBox.bind("<Delete>", deleteUnitDel)
                ExRateBox.bind("<Control-d>", setDeleteUnit)
                ExRateBox.bind("<Escape>", closeUnitTabEsc)


    
                queryTreeUnit()

    
    
    
    
    
    
    
    
    
    
            # LABEL FOR ASSEM LAYER
            
            AssemDataFrame = LabelFrame(frameAssem, text="Record")
            AssemDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
            # AssemDataFrame.pack(fill="x", expand="yes", padx=20)
            
            AssemButtonFrame = LabelFrame(frameAssem, text="Command")
            AssemButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
            # AssemButtonFrame.pack(fill="x", expand="yes", padx=20)
            
            TypeLabel = Label(AssemDataFrame, text="Type", font=("Arial", 8))
            NumLabel = Label(AssemDataFrame, text="Number", font=("Arial", 8))
            DescLabel = Label(AssemDataFrame, text="Description", font=("Arial", 8))
            DesignLabel = Label(AssemDataFrame, text="Design", font=("Arial", 8))
            PurchaseLabel = Label(AssemDataFrame, text="Purchase", font=("Arial", 8))
            AssemLabel = Label(AssemDataFrame, text="Assembly", font=("Arial", 8))
            DesignerLabel = Label(AssemDataFrame, text="Designer", font=("Arial", 8))
            LockerLabel = Label(AssemDataFrame, text="Locker", font=("Arial", 8))
            DesignDueLabel = Label(AssemDataFrame, text="Design Due", font=("Arial", 8))
            DesignCompletedLabel = Label(AssemDataFrame, text="Design Completed", font=("Arial", 8))
            POIssueDueLabel = Label(AssemDataFrame, text="PO Issue Due", font=("Arial", 8))
            POIssueCompletedLabel = Label(AssemDataFrame, text="PO Issue Completed", font=("Arial", 8))
            PartsReceivedDueLabel = Label(AssemDataFrame, text="Parts Received Due", font=("Arial", 8))
            PartsReceivedCompletedLabel = Label(AssemDataFrame, text="Parts Received Completed", font=("Arial", 8))
            DesignApprovedLabel = Label(AssemDataFrame, text="Design Approved", font=("Arial", 8))
            ApproveNameLabel = Label(AssemDataFrame, text="Approved By", font=("Arial", 8))
            NumOfPartsLabel = Label(AssemDataFrame, text="Parts Required", font=("Arial", 8))
            PartsPurchasedLabel = Label(AssemDataFrame, text="Parts Purchased", font=("Arial", 8))
            PartsReceivedLabel = Label(AssemDataFrame, text="Parts Received", font=("Arial", 8))
            DeletedLabel = Label(AssemDataFrame, text="Deleted", font=("Arial", 8))
            OnHoldLabel = Label(AssemDataFrame, text="On Hold", font=("Arial", 8))
            DesignQtyLabel = Label(AssemDataFrame, text="Design Qty", font=("Arial", 8))
            CompletedLabel = Label(AssemDataFrame, text="Completed Qty", font=("Arial", 8))
            EstCostAssemLabel = Label(AssemDataFrame, text="Estimated Cost", font=("Arial", 8))
            
            TypeLabel.grid(row=0, column=0, padx=10, pady=2, sticky=E)
            DesignLabel.grid(row=0, column=5, padx=10, pady=2, sticky=E)
            PurchaseLabel.grid(row=0, column=7, padx=10, pady=2, sticky=E)
            AssemLabel.grid(row=0, column=9, padx=10, pady=2, sticky=E)
            NumLabel.grid(row=1, column=0, padx=10, pady=2, sticky=E)
            DesignerLabel.grid(row=2, column=0, padx=10, pady=2, sticky=E)
            LockerLabel.grid(row=3, column=0, padx=10, pady=2, sticky=E)
            DescLabel.grid(row=0, column=2, padx=10, pady=2, sticky=E)
            DesignDueLabel.grid(row=1, column=2, padx=10, pady=2, sticky=E)
            POIssueDueLabel.grid(row=2, column=2, padx=10, pady=2, sticky=E)
            PartsReceivedDueLabel.grid(row=3, column=2, padx=10, pady=2, sticky=E)
            DesignCompletedLabel.grid(row=4, column=2, padx=10, pady=2, sticky=E)
            POIssueCompletedLabel.grid(row=5, column=2, padx=10, pady=2, sticky=E)
            PartsReceivedCompletedLabel.grid(row=6, column=2, padx=10, pady=2, sticky=E)
            DesignApprovedLabel.grid(row=1, column=5, padx=10, pady=2, sticky=E)
            ApproveNameLabel.grid(row=2, column=5, padx=10, pady=2, sticky=E)
            NumOfPartsLabel.grid(row=3, column=5, padx=10, pady=2, sticky=E)
            PartsPurchasedLabel.grid(row=4, column=5, padx=10, pady=2, sticky=E)
            PartsReceivedLabel.grid(row=5, column=5, padx=10, pady=2, sticky=E)
            DeletedLabel.grid(row=1, column=7, padx=10, pady=2, sticky=E)
            OnHoldLabel.grid(row=2, column=7, padx=10, pady=2, sticky=E)
            DesignQtyLabel.grid(row=1, column=9, padx=10, pady=2, sticky=E)
            CompletedLabel.grid(row=2, column=9, padx=10, pady=2, sticky=E)
            EstCostAssemLabel.grid(row=3, column=9, padx=10, pady=2, sticky=E)
            

    

    
            TypeBox = ttk.Combobox(AssemDataFrame, width=15, 
                                   value=["Mechanical", "Electrical", "Pneumatic"], 
                                   font=("Arial", 8), state="readonly")
            DesignBox = ttk.Combobox(AssemDataFrame, width=15, 
                                     value=["Not Started", "In Progress", "Completed"], 
                                     font=("Arial", 8), state="readonly")
            PurchaseBox = ttk.Combobox(AssemDataFrame, width=15, 
                                       value=["Not Started", "In Progress", "Completed"], 
                                       font=("Arial", 8), state="readonly")
            AssemBox = ttk.Combobox(AssemDataFrame, width=15, 
                                    value=["Not Started", "In Progress", "Completed"], 
                                    font=("Arial", 8), state="readonly")
            NumBox = Entry(AssemDataFrame, width=12, font=("Arial", 8))
            AssemIDGenButton = Button(AssemDataFrame, text="Gen", font=("Arial", 7), command=genAssemID)
            DesignerBox = ttk.Combobox(AssemDataFrame, width=15, value=FetchEmployees(), 
                                       font=("Arial", 8), state="readonly")
            LockerBox = ttk.Combobox(AssemDataFrame, width=15, value=FetchEmployees(), 
                                     font=("Arial", 8), state="readonly")
            DescBox = Entry(AssemDataFrame, width=25, font=("Arial", 8))
            
            DesignDueBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            POIssueDueBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            PartsReceivedDueBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            DesignCompletedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            POIssueCompletedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            PartsReceivedCompletedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            DesDueCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=DesDueAssem)
            POIDueCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=POIDueAssem)
            PartDueCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=PartDueAssem)
            DesCompCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=DesCompAssem)
            POICompCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=POICompAssem)
            PartCompCalendarButton = Button(AssemDataFrame, text="Cal", font=("Arial", 7), command=PartCompAssem)
            
            DesignApprovedBox = ttk.Combobox(AssemDataFrame, width=15, value=["No", "Yes"], 
                                             font=("Arial", 8), state="readonly")
            ApproveNameBox = ttk.Combobox(AssemDataFrame, width=15, value=["NIL"], 
                                          font=("Arial", 8), state="readonly")
    
            NumOfPartsBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            PartsPurchasedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            PartsReceivedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            DeletedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            OnHoldBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            
            DesignQtyBox = Spinbox(AssemDataFrame, from_=0, to=1000, width=13, font=("Arial", 8), state="readonly")
            CompletedBox = Entry(AssemDataFrame, width=15, font=("Arial", 8), state="readonly")
            
            EstCostAssemBox = Entry(AssemDataFrame, width=12, font=("Arial", 8), state="readonly")
            EstCostSGDLabel = Label(AssemDataFrame, text="SGD", font=("Arial", 8))
            
            TypeBox.current(0)
            DesignBox.current(0)
            PurchaseBox.current(0)
            AssemBox.current(0)
            DesignerBox.current(0)
            LockerBox.current(0)
            DesignApprovedBox.current(0)
            ApproveNameBox.current(0)
            
            DesignQtyBox.config(state="normal")
            DesignQtyBox.delete(0, END)
            DesignQtyBox.insert(0,1)
            DesignQtyBox.config(state="readonly")
            

            
            
            
            TypeBox.grid(row=0, column=1, columnspan=2, padx=10, pady=2, sticky=W)
            DesignBox.grid(row=0, column=6, padx=10, pady=2, sticky=W)
            PurchaseBox.grid(row=0, column=8, padx=10, pady=2, sticky=W)
            AssemBox.grid(row=0, column=10, columnspan=2, padx=10, pady=2, sticky=W)
            NumBox.grid(row=1, column=1, padx=10, pady=2, sticky=W)
            AssemIDGenButton.grid(row=1, column=2, padx=0, pady=1, sticky=W)
            DesignerBox.grid(row=2, column=1, columnspan=2, padx=10, pady=2, sticky=W)
            LockerBox.grid(row=3, column=1, columnspan=2, padx=10, pady=2, sticky=W)
            DescBox.grid(row=0, column=3, columnspan=2, padx=10, pady=2, sticky=W)
            
            DesignDueBox.grid(row=1, column=3, padx=10, pady=2, sticky=W)
            POIssueDueBox.grid(row=2, column=3, padx=10, pady=2, sticky=W)
            PartsReceivedDueBox.grid(row=3, column=3, padx=10, pady=2, sticky=W)
            DesignCompletedBox.grid(row=4, column=3, padx=10, pady=2, sticky=W)
            POIssueCompletedBox.grid(row=5, column=3, padx=10, pady=2, sticky=W)
            PartsReceivedCompletedBox.grid(row=6, column=3, padx=10, pady=2, sticky=W)
            DesDueCalendarButton.grid(row=1, column=4, padx=0, pady=1, sticky=W)
            POIDueCalendarButton.grid(row=2, column=4, padx=0, pady=1, sticky=W)
            PartDueCalendarButton.grid(row=3, column=4, padx=0, pady=1, sticky=W)
            DesCompCalendarButton.grid(row=4, column=4, padx=0, pady=1, sticky=W)
            POICompCalendarButton.grid(row=5, column=4, padx=0, pady=1, sticky=W)
            PartCompCalendarButton.grid(row=6, column=4, padx=0, pady=1, sticky=W)
            
            DesignApprovedBox.grid(row=1, column=6, padx=10, pady=2, sticky=W)
            ApproveNameBox.grid(row=2, column=6, padx=10, pady=2, sticky=W)
            NumOfPartsBox.grid(row=3, column=6, padx=10, pady=2, sticky=W)
            PartsPurchasedBox.grid(row=4, column=6, padx=10, pady=2, sticky=W)
            PartsReceivedBox.grid(row=5, column=6, padx=10, pady=2, sticky=W)
            DeletedBox.grid(row=1, column=8, padx=10, pady=2, sticky=W)
            OnHoldBox.grid(row=2, column=8, padx=10, pady=2, sticky=W)
            DesignQtyBox.grid(row=1, column=10, columnspan=2, padx=10, pady=2, sticky=W)
            CompletedBox.grid(row=2, column=10, columnspan=2, padx=10, pady=2, sticky=W)
            
            EstCostAssemBox.grid(row=3, column=10, padx=10, pady=2, sticky=W)
            EstCostSGDLabel.grid(row=3, column=11, padx=0, pady=2, sticky=W)
            
            if AuthLevel(Login.AUTHLVL, 2) == False:
                EstCostAssemLabel.grid_forget()
                EstCostAssemBox.grid_forget()
                EstCostSGDLabel.grid_forget()
            
            
    
            buttonUpdateAssem = Button(AssemButtonFrame, text="Update Assembly", command=updateAssem, state=DISABLED)
            buttonCreateAssem = Button(AssemButtonFrame, text="Create Assembly", command=createAssem)
            buttonDeleteAssem = Button(AssemButtonFrame, text="Delete Assembly", command=deleteAssem, state=DISABLED)
            buttonSelectAssem = Button(AssemButtonFrame, text="Select Assembly", command=selectAssem)
            buttonDeselectAssem = Button(AssemButtonFrame, text="Deselect Assembly", command=deselectAssem)
            buttonLoadAssem = Button(AssemButtonFrame, text="Load Assembly", command=loadAssem)
            buttonClearEntryAssem = Button(AssemButtonFrame, text="Clear Entry", command=clearEntryAssem)
            buttonRefreshAssem = Button(AssemButtonFrame, text="Refresh", command=refreshAssem)
            buttonCloseTabAssem = Button(AssemButtonFrame, text="Close Tab", command=closeTabAssem)
        
            buttonUpdateAssem.grid(row=0, column=0, padx=10, pady=10)
            buttonCreateAssem.grid(row=0, column=1, padx=10, pady=10)
            buttonDeleteAssem.grid(row=0, column=2, padx=10, pady=10)
            buttonSelectAssem.grid(row=0, column=3, padx=10, pady=10)
            buttonDeselectAssem.grid(row=0, column=4, padx=10, pady=10)
            buttonLoadAssem.grid(row=0, column=5, padx=10, pady=10)
            buttonClearEntryAssem.grid(row=0, column=6, padx=(90,10), pady=10)
            buttonRefreshAssem.grid(row=0, column=7, padx=10, pady=10)
            buttonCloseTabAssem.grid(row=0, column=8, padx=10, pady=10)
            
            DesignApprovedBox.bind("<<ComboboxSelected>>", approveClick)
            
            queryTreeAssem()
    
    
    
    
    
    
    
    
    
    
        # LABEL FOR MACHINE LAYER
    
        MachIDLabel = Label(MachDataFrame, text="Machine ID")
        MachNameLabel = Label(MachDataFrame, text="Machine Name")
        MachStatusLabel = Label(MachDataFrame, text="Status")
        MachDueDateLabel = Label(MachDataFrame, text="Due Date")
        MachCompDateLabel = Label(MachDataFrame, text="Completed Date")
        MachDesignerLabel = Label(MachDataFrame, text="Designer")
        MachLockerLabel = Label(MachDataFrame, text="Locker")
        MachOrderLabel = Label(MachDataFrame, text="Order Qty")
        MachCompLabel = Label(MachDataFrame, text="Complete Qty")
        MachUALabel = Label(MachDataFrame, text="No. of UA")
        MachBOMLabel = Label(MachDataFrame, text="BOM Load")
        MachDeletedLabel = Label(MachDataFrame, text="Deleted")
        MachOnHoldLabel = Label(MachDataFrame, text="On Hold")
        MachEstPartLabel = Label(MachDataFrame, text="Est. Part")
        MachEstCostLabel = Label(MachDataFrame, text="Est. Cost")
    
        MachIDBox = Entry(MachDataFrame, width=12)
        MachIDGenButton = Button(MachDataFrame, text="Gen", width=4, command=genMachID)
        MachNameBox = Entry(MachDataFrame, width=22)
        MachStatusBox = ttk.Combobox(MachDataFrame, width=10, 
                                     value=["Not Started", "In Progress", "Completed"],
                                     state="readonly")
        
        MachDueDateBox = Entry(MachDataFrame, width=12, state="readonly")
        MachCompDateBox = Entry(MachDataFrame, width=12, state="readonly")
        MachDueCalendarButton = Button(MachDataFrame, text="Cal", width=4, command=dueCalMach)
        MachCompCalendarButton = Button(MachDataFrame, text="Cal", width=4, command=compCalMach)
        
        MachDesignerBox = ttk.Combobox(MachDataFrame, width=20, value=FetchEmployees(), state="readonly")
        MachLockerBox = ttk.Combobox(MachDataFrame, width=20, value=FetchEmployees(), state="readonly")
        
        MachOrderBox = Spinbox(MachDataFrame, from_=0, to=100, width=10, state="readonly")
        MachCompBox = Entry(MachDataFrame, width=12, state="readonly")
    
        MachUABox = Entry(MachDataFrame, width=10, state="readonly")
        MachBOMBox = Entry(MachDataFrame, width=10, state="readonly")
        MachDeletedBox = Entry(MachDataFrame, width=10, state="readonly")
        MachOnHoldBox = Entry(MachDataFrame, width=10, state="readonly")
        MachEstPartBox = Entry(MachDataFrame, width=15, state="readonly")
        MachEstCostBox = Entry(MachDataFrame, width=10, state="readonly")
        MachEstCostSGDLabel = Label(MachDataFrame, text="SGD")
        
        MachStatusBox.current(0)
        MachDesignerBox.current(0)
        MachLockerBox.current(0)
        
        MachOrderBox.config(state="normal")
        MachOrderBox.delete(0, END)
        MachOrderBox.insert(0,1)
        MachOrderBox.config(state="readonly")

        
        
        
        
        
        MachIDLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
        MachNameLabel.grid(row=0, column=3, padx=10, pady=5, sticky=E)
        MachStatusLabel.grid(row=0, column=5, padx=10, pady=5, sticky=E)
        MachDueDateLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
        MachCompDateLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
        MachDesignerLabel.grid(row=1, column=3, padx=10, pady=5, sticky=E)
        MachLockerLabel.grid(row=2, column=3, padx=10, pady=5, sticky=E)
        MachOrderLabel.grid(row=1, column=5, padx=10, pady=5, sticky=E)
        MachCompLabel.grid(row=2, column=5, padx=10, pady=5, sticky=E)
        MachUALabel.grid(row=0, column=7, padx=10, pady=5, sticky=E)
        MachBOMLabel.grid(row=1, column=7, padx=10, pady=5, sticky=E)
        MachDeletedLabel.grid(row=2, column=7, padx=10, pady=5, sticky=E)
        MachOnHoldLabel.grid(row=3, column=7, padx=10, pady=5, sticky=E)
        MachEstPartLabel.grid(row=0, column=9, padx=10, pady=5, sticky=E)
        MachEstCostLabel.grid(row=1, column=9, padx=10, pady=5, sticky=E)
        
        MachIDBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
        MachIDGenButton.grid(row=0, column=2, padx=5, pady=5, sticky=W)
        MachNameBox.grid(row=0, column=4, padx=10, pady=5, sticky=W)
        MachStatusBox.grid(row=0, column=6, padx=10, pady=5, sticky=W)
        
        MachDueDateBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
        MachCompDateBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
        MachDueCalendarButton.grid(row=1, column=2, padx=5, pady=5, sticky=W)
        MachCompCalendarButton.grid(row=2, column=2, padx=5, pady=5, sticky=W)
        
        MachDesignerBox.grid(row=1, column=4, padx=10, pady=5, sticky=W)
        MachLockerBox.grid(row=2, column=4, padx=10, pady=5, sticky=W)
        
        MachOrderBox.grid(row=1, column=6, padx=10, pady=5, sticky=W)
        MachCompBox.grid(row=2, column=6, padx=10, pady=5, sticky=W)
    
        MachUABox.grid(row=0, column=8, padx=10, pady=5, sticky=W)
        MachBOMBox.grid(row=1, column=8, padx=10, pady=5, sticky=W)
        MachDeletedBox.grid(row=2, column=8, padx=10, pady=5, sticky=W)
        MachOnHoldBox.grid(row=3, column=8, padx=10, pady=5, sticky=W)
        MachEstPartBox.grid(row=0, column=10, columnspan=2, padx=10, pady=5, sticky=W)
        MachEstCostBox.grid(row=1, column=10, padx=10, pady=5, sticky=W)
        MachEstCostSGDLabel.grid(row=1, column=11, padx=(0,10), pady=5, sticky=W)
        
        if AuthLevel(Login.AUTHLVL, 2) == False:
            MachEstCostLabel.grid_forget()
            MachEstCostBox.grid_forget()
            MachEstCostSGDLabel.grid_forget()
        
    

    
    
        buttonUpdateMach = Button(MachButtonFrame, text="Update Machine", command=updateMach, state=DISABLED)
        buttonCreateMach = Button(MachButtonFrame, text="Create Machine", command=createMach)
        buttonDeleteMach = Button(MachButtonFrame, text="Delete Machine", command=deleteMach, state=DISABLED)
        buttonSelectMach = Button(MachButtonFrame, text="Select Machine", command=selectMach)
        buttonDeselectMach = Button(MachButtonFrame, text="Deselect Machine", command=deselectMach)
        buttonLoadMach = Button(MachButtonFrame, text="Load Machine", command=loadMach)
        buttonClearEntryMach = Button(MachButtonFrame, text="Clear Entry", command=clearEntryMach)
        buttonRefreshMach = Button(MachButtonFrame, text="Refresh", command=refreshMach)
        buttonCloseTabMach = Button(MachButtonFrame, text="Close Tab", command=closeTabMach)
    
        buttonUpdateMach.grid(row=0, column=0, padx=10, pady=10)
        buttonCreateMach.grid(row=0, column=1, padx=10, pady=10)
        buttonDeleteMach.grid(row=0, column=2, padx=10, pady=10)
        buttonSelectMach.grid(row=0, column=3, padx=10, pady=10)
        buttonDeselectMach.grid(row=0, column=4, padx=10, pady=10)
        buttonLoadMach.grid(row=0, column=5, padx=10, pady=10)
        buttonClearEntryMach.grid(row=0, column=6, padx=(90,10), pady=10)
        buttonRefreshMach.grid(row=0, column=7, padx=10, pady=10)
        buttonCloseTabMach.grid(row=0, column=8, padx=10, pady=10)
    
        queryTreeMach()
    
    
    
    
    # LABEL FOR PROJECT LAYER
    
    proDataFrame = LabelFrame(framePro, text="Record")
    proDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # proDataFrame.pack(fill="x", expand="yes", padx=20)
    
    proButtonFrame = LabelFrame(framePro, text="Command")
    proButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # proButtonFrame.pack(fill="x", expand="yes", padx=20)
    
    ProIDLabel = Label(proDataFrame, text="Project ID")
    ProNameLabel = Label(proDataFrame, text="Project Name")
    ProStatusLabel = Label(proDataFrame, text="Status")
    ProStartDateLabel = Label(proDataFrame, text="Start Date")
    ProDeliveryDateLabel = Label(proDataFrame, text="Delivery Date")
    ProManagerLabel = Label(proDataFrame, text="Project Manager")
    ProSupportLabel = Label(proDataFrame, text="Project Support")
    ProBasePlanLockLabel = Label(proDataFrame, text="Base Plan Lock")
    ProLockLabel = Label(proDataFrame, text="Project Lock")
    ProDeletedLabel = Label(proDataFrame, text="Deleted")
    # ProVersionLabel = Label(proDataFrame, text="Version")
    ProOnHoldLabel = Label(proDataFrame, text="On Hold")
    ProEstPartLabel = Label(proDataFrame, text="Est. Part")
    ProEstCostLabel = Label(proDataFrame, text="Est. Cost")
    
    ProIDBox = Entry(proDataFrame, width=12)
    ProIDGenButton = Button(proDataFrame, text="Gen", width=4, command=genProID)
    ProNameBox = Entry(proDataFrame, width=30)
    ProStatusBox = ttk.Combobox(proDataFrame, width=10, 
                                value=["Not Started", "In Progress", "Completed"], 
                                state="readonly")
    
    ProStartDateBox = Entry(proDataFrame, width=12, state="readonly")
    ProDeliveryDateBox = Entry(proDataFrame, width=12, state="readonly")
    ProStartCalendarButton = Button(proDataFrame, text="Cal", width=4, command=startCalPro)
    ProDeliveryCalendarButton = Button(proDataFrame, text="Cal", width=4, command=deCalPro)
    
    ProManagerBox = ttk.Combobox(proDataFrame, width=28, value=FetchEmployees(), state="readonly")
    ProSupportBox = ttk.Combobox(proDataFrame, width=28, value=FetchEmployees(), state="readonly")
    
    ProBasePlanLockVal = IntVar()
    ProLockVal = IntVar()
    ProBasePlanLockBox = Checkbutton(proDataFrame, variable=ProBasePlanLockVal, onvalue=1, offvalue=0)
    ProLockBox = Checkbutton(proDataFrame, variable=ProLockVal, onvalue=1, offvalue=0)
    
    ProDeletedBox = Entry(proDataFrame, width=12, state="readonly")
    ProOnHoldBox = Entry(proDataFrame, width=12, state="readonly")
    
    ProEstPartBox = Entry(proDataFrame, width=12, state="readonly")
    ProEstCostBox = Entry(proDataFrame, width=12, state="readonly")
    ProSGDLabel = Label(proDataFrame, text="SGD")   
    
    ProStatusBox.current(0)
    ProManagerBox.current(0)
    ProSupportBox.current(0)
    ProBasePlanLockVal.set(0)
    ProLockVal.set(0)
    
    
    
    
    
    ProIDLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    ProNameLabel.grid(row=0, column=3, padx=10, pady=5, sticky=E)
    ProStatusLabel.grid(row=0, column=7, padx=10, pady=5, sticky=E)
    ProStartDateLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    ProDeliveryDateLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    ProManagerLabel.grid(row=1, column=3, padx=10, pady=5, sticky=E)
    ProSupportLabel.grid(row=2, column=3, padx=10, pady=5, sticky=E)
    ProBasePlanLockLabel.grid(row=0, column=5, padx=10, pady=5, sticky=E)
    ProLockLabel.grid(row=1, column=5, padx=10, pady=5, sticky=E)
    ProDeletedLabel.grid(row=1, column=7, padx=10, pady=5, sticky=E)
    ProOnHoldLabel.grid(row=2, column=7, padx=10, pady=5, sticky=E)
    ProEstPartLabel.grid(row=0, column=9, padx=10, pady=5, sticky=E)
    ProEstCostLabel.grid(row=1, column=9, padx=10, pady=5, sticky=E)
    
    ProIDBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    ProNameBox.grid(row=0, column=4, padx=10, pady=5, sticky=W)
    ProStatusBox.grid(row=0, column=8, padx=10, pady=5, sticky=W)
    
    ProIDGenButton.grid(row=0, column=2, padx=5, pady=5, sticky=W)
    ProStartDateBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    ProStartCalendarButton.grid(row=1, column=2, padx=5, pady=5, sticky=W)
    ProDeliveryDateBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    ProDeliveryCalendarButton.grid(row=2, column=2, padx=5, pady=5, sticky=W)
    
    ProManagerBox.grid(row=1, column=4, padx=10, pady=5, sticky=W)
    ProSupportBox.grid(row=2, column=4, padx=10, pady=5, sticky=W)
    ProBasePlanLockBox.grid(row=0, column=6, padx=10, pady=5, sticky=W)
    ProLockBox.grid(row=1, column=6, padx=10, pady=5, sticky=W)
    ProDeletedBox.grid(row=1, column=8, padx=10, pady=5, sticky=W)
    ProOnHoldBox.grid(row=2, column=8, padx=10, pady=5, sticky=W)
    
    ProEstPartBox.grid(row=0, column=10, columnspan=2, padx=10, pady=5, sticky=W)
    ProEstCostBox.grid(row=1, column=10, padx=10, pady=5, sticky=W)
    ProSGDLabel.grid(row=1, column=11, padx=0, pady=5, sticky=W)
    
    if AuthLevel(Login.AUTHLVL, 2) == False:
        ProEstCostLabel.grid_forget()
        ProEstCostBox.grid_forget()
        ProSGDLabel.grid_forget()
    
    
    
    def openVendor():
        Vendor.RunVEND()
        
    def openRFQ():
        RFQ.openRFQ(tabNote)
        
    def reviewRFQ():
        RFQ.reviewRFQ(tabNote)
    
    def ConfigureChrome():
        OverWrite(Reconfigure = True)
    
    menuBar = Menu(root)
    root.config(menu=menuBar)
        
    fileMenu = Menu(menuBar, tearoff=0)
    menuBar.add_cascade(label="File", menu=fileMenu)
    fileMenu.add_command(label="Edit Employee", command=EmployeeTest.openEmployee)
    fileMenu.add_command(label="Edit Vendor", command=openVendor)
    fileMenu.add_command(label="Edit Stock", command=StockTest.openStock)
    fileMenu.add_separator()
    fileMenu.add_command(label="Edit Currency Exchange Rate", command=CountryRef.openCurrency)
    fileMenu.add_separator()
    fileMenu.add_command(label="Edit Chrome/Xero Details", command=ConfigureChrome)
    fileMenu.add_separator()
    fileMenu.add_command(label="Exit", command=root.destroy)
        
    ProcurementMenu = Menu(menuBar, tearoff=0)
    menuBar.add_cascade(label="Procurement", menu=ProcurementMenu)
    ProcurementMenu.add_command(label="Generate RFQ", command=openRFQ)
    ProcurementMenu.add_separator()
    ProcurementMenu.add_command(label="Generate Purchase Order", command=ReportTest.openPurchase)
    
    
    helpMenu = Menu(menuBar, tearoff=0)
    menuBar.add_cascade(label="Help", menu=helpMenu)
    helpMenu.add_command(label="Instructions", command=Instruction.openInstruction)    
    
    if AuthLevel(Login.AUTHLVL, 6) == False:
        ProcurementMenu.entryconfig("Generate Purchase Order", state="disabled")
        ProcurementMenu.entryconfig("Generate RFQ", state="disabled")
    
    
    
    
    
    
    
    
    
    
    buttonUpdatePro = Button(proButtonFrame, text="Update Project", command=updatePro, state=DISABLED)
    buttonUpdatePro.grid(row=0, column=0, padx=10, pady=10, sticky=W)
            
    buttonCreatePro = Button(proButtonFrame, text="Create Project", command=createPro)
    buttonCreatePro.grid(row=0, column=1, padx=10, pady=10, sticky=W)
    
    buttonDeletePro = Button(proButtonFrame, text="Delete Project", command=deletePro, state=DISABLED)
    buttonDeletePro.grid(row=0, column=2, padx=10, pady=10, sticky=W)
    
    buttonSelectPro = Button(proButtonFrame, text="Select Project", command=selectPro)
    buttonSelectPro.grid(row=0, column=3, padx=10, pady=10, sticky=W)
    
    buttonDeselectPro = Button(proButtonFrame, text="Deselect Project", command=deselectPro)
    buttonDeselectPro.grid(row=0, column=4, padx=10, pady=10, sticky=W)
    
    buttonLoadPro = Button(proButtonFrame, text="Load Project", command=loadPro)
    buttonLoadPro.grid(row=0, column=5, padx=10, pady=10, sticky=W)
    
    buttonClearEntryPro = Button(proButtonFrame, text="Clear Entry", command=clearEntryPro)
    buttonClearEntryPro.grid(row=0, column=6, padx=(90,10), pady=10, sticky=W)
    
    buttonRefreshPro = Button(proButtonFrame, text="Refresh", command=refreshPro)
    buttonRefreshPro.grid(row=0, column=7, padx=10, pady=10, sticky=W)
    
    if AuthLevel(Login.AUTHLVL, 3) == False:
        buttonDeletePro.grid_forget()
    
    queryTreePro()
    checkEmpPro()
    
    
    
    root.mainloop()
    
    
    
    
    
        
