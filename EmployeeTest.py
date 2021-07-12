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
import ConnConfig

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

connInit = mysql.connector.connect(host = logininfo[0],
                                   user = logininfo[1], 
                                   password =logininfo[2])

connMain = mysql.connector.connect(host = logininfo[0],
                                   user = logininfo[1], 
                                   password =logininfo[2],
                                   database= "INDEX_PRO_MASTER")

connEmp = mysql.connector.connect(host = logininfo[0],
                                  user = logininfo[1], 
                                  password =logininfo[2],
                                  database= "INDEX_EMP_MASTER")


def openEmployee():
    EmpWin = Toplevel()
    EmpWin.iconbitmap("MWA_Icon.ico")
    EmpWin.title("Employee Menu")
    EmpWin.state("zoomed")
    EmpWin.columnconfigure(0, weight=1)
    
    EmpTitleLabel = Label(EmpWin, text = "Employee List", font=("Arial", 12))
    EmpTitleLabel.grid(row=0, column=0, padx=0, pady=0, ipadx=0, ipady=10, sticky=W+E) 
    
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background="silver",
                    rowheight=20,
                    fieldbackground="light grey")
    
    style.map("Treeview")
    
    EmpTreeFrame = Frame(EmpWin)

    EmpTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
    # EmpTreeFrame.pack()
    
    EmpTreeScroll = Scrollbar(EmpTreeFrame)
    EmpTreeScroll.pack(side=RIGHT, fill=Y)
    
    EmpTreeView = ttk.Treeview(EmpTreeFrame, yscrollcommand=EmpTreeScroll.set, 
                                selectmode="browse")
    EmpTreeScroll.config(command=EmpTreeView.yview)
    # EmpTreeView.grid(row=2, column=0, columnspan=10, padx=10)
    EmpTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill=X, expand=True)
    
    EmpTreeView['columns'] = ("Employee ID", "Employee Name", "Employee Class", 
                              "Nationality", "Date of Birth", "Home Address",
                              "Join Date", "Status")
    
    # EmpTreeView.column("#0", anchor = CENTER, width =45, minwidth = 0)
    EmpTreeView.column("#0",  width=0, stretch=NO)
    EmpTreeView.column("Employee ID",anchor = CENTER, width =60, minwidth = 50)
    EmpTreeView.column("Employee Name",anchor = CENTER, width =150, minwidth = 50)
    EmpTreeView.column("Employee Class",anchor = CENTER, width =120, minwidth = 50)
    EmpTreeView.column("Nationality",anchor = CENTER, width =80, minwidth = 50)
    EmpTreeView.column("Date of Birth",anchor = CENTER, width =100, minwidth = 50)
    EmpTreeView.column("Home Address",anchor = CENTER, width = 240, minwidth = 50)
    EmpTreeView.column("Join Date",anchor = CENTER, width =100, minwidth = 50)
    EmpTreeView.column("Status",anchor = CENTER, width =50, minwidth = 50)
    
    EmpTreeView.heading("#0",text = "Label")
    EmpTreeView.heading("Employee ID",text = "ID")
    EmpTreeView.heading("Employee Name",text = "Name")
    EmpTreeView.heading("Employee Class",text = "Class")
    EmpTreeView.heading("Nationality",text = "Nationality")
    EmpTreeView.heading("Date of Birth",text = "D.O.B.")
    EmpTreeView.heading("Home Address",text = "Home Address")
    EmpTreeView.heading("Join Date",text = "Joined Date")
    EmpTreeView.heading("Status",text = "Status")
    
    def queryTreeEmp():
        curEmp = connEmp.cursor()
        curEmp.execute("SELECT * FROM EMP_DATA")
        recLst = curEmp.fetchall()    

        classLst = ["Designer", "Purchaser", "Project Manager", "Administrator"]
        statLst = ["Inactive", "Active"]
        
        for rec in recLst:
            EmpTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                values=(rec[1], rec[2], classLst[rec[3]], 
                                        rec[4], rec[5], rec[6], rec[7],
                                        statLst[rec[8]]))
        connEmp.commit()
        curEmp.close()
    
    
    def genEmpID():
        empClass = EmpClassBox.current()
        curEmp = connEmp.cursor()
        curEmp.execute(f"SELECT * FROM EMP_DATA WHERE EMPLOYEE_CLASS = {empClass}")
        empRefLst = curEmp.fetchall()
        connEmp.commit()
        curEmp.close()
        
        numLst = []
        for i in range(len(empRefLst)):
            numLst.append(int(empRefLst[i][1][-2:]))
        
        if len(numLst) == 0:
            currentNum = 0
        else:
            currentNum = max(numLst)
        
        nextNum = currentNum + 1
        twoDigit = str(nextNum).rjust(2,"0")
        
        classLst = ["DS", "PC", "PM", "ADM"]
        nextID = f"{classLst[empClass]}{twoDigit}"
        
        EmpIDBox.delete(0, END)
        EmpIDBox.insert(0, nextID)
        
    def genEmpNat():
        natWin = Toplevel()
        natWin.title("Select Country")
        natWin.geometry("210x300")
        
        sclFrame = Frame(natWin)
        sclFrame.grid(row=0, column=0, columnspan=3, padx=20, pady=10, ipadx=5, ipady=0, sticky=W+E)
                
        scl = Scrollbar(sclFrame, orient=VERTICAL)
        scl.pack(side=RIGHT, fill=Y)
        
        natLstBox = Listbox(sclFrame, width=20, selectmode=SINGLE, yscrollcommand=scl.set)
        scl.config(command=natLstBox.yview)
        natLstBox.pack(ipady=20)
        
        natLst = ["Singaporean", "Chinese", "Malaysian", "Indian",
                  "American", "British", "Canadian", "Australian", "New Zealander",
                  "Japanese", "Korean", "Vietnamese", "Thai", "Burmese", 
                  "Cambodian", "Laotian", "Filipino", "Indonesian", "Bruneian",
                  "Others (Please Specify)"]
        
        for nat in natLst:
            natLstBox.insert(END, nat)
        
        EmpNationBox.delete(0, END)
        EmpNationBox.config(state="readonly")
        
        def confirmNat():
            empNat = natLstBox.get(ANCHOR)
            if empNat == "":
                messagebox.showwarning("No Country Selected", "Please Select a Country", 
                                       parent=natWin)
            elif empNat == "Others (Please Specify)":
                EmpNationBox.config(state="normal")
                EmpNationBox.delete(0, END)
                natWin.destroy()
            else:
                EmpNationBox.config(state="normal")
                EmpNationBox.delete(0, END)
                EmpNationBox.insert(0, empNat)
                EmpNationBox.config(state="readonly")
                natWin.destroy()
        
        def emptyNat():
            EmpNationBox.config(state="normal")
            EmpNationBox.delete(0, END)
            EmpNationBox.config(state="readonly")
            natWin.destroy()
        
        buttonConfirm = Button(natWin, text="Confirm", command=confirmNat)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
        
        buttonEmpty = Button(natWin, text="Clear", command=emptyNat)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
        
        buttonClose = Button(natWin, text="Close", command=natWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
    def DOBCalEmp():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("240x240")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd",
                       year=2000, month=1, day=1)
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            EmpDOBBox.config(state="normal")
            EmpDOBBox.delete(0, END)
            EmpDOBBox.insert(0, val)
            EmpDOBBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            EmpDOBBox.config(state="normal")
            EmpDOBBox.delete(0, END)
            EmpDOBBox.config(state="readonly")
            calWin.destroy()

        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
        
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
        
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
    def DateJoinCalEmp():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("240x240")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            EmpDateJointBox.config(state="normal")
            EmpDateJointBox.delete(0, END)
            EmpDateJointBox.insert(0, val)
            EmpDateJointBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            EmpDateJointBox.config(state="normal")
            EmpDateJointBox.delete(0, END)
            EmpDateJointBox.config(state="readonly")
            calWin.destroy()

        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
        
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
        
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
    
    def updateEmp():
        
        sqlCommand = f"""UPDATE EMP_DATA SET
        `EMPLOYEE_ID` = %s,
        `EMPLOYEE_NAME` = %s, 
        `EMPLOYEE_CLASS` = %s,
        `NATIONALITY` = %s,
        `DOB` = %s,
        `ADDRESS` = %s,
        `JOIN_DATE` = %s,
        `STATUS` = %s
        
        WHERE `oid` = %s"""
        
        
        

        
        def checkDateEmp(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        selected = EmpTreeView.focus()
        
        inputs = (EmpIDBox.get(), EmpNameBox.get(), EmpClassBox.current(), 
                  EmpNationBox.get(), checkDateEmp(EmpDOBBox), EmpAddressBox.get(),
                  checkDateEmp(EmpDateJointBox), EmpStatusBox.current(), selected)
        
        inputs2 =(EmpIDBox.get(), EmpClassBox.current(), selected)
        empFullName = f"{EmpIDBox.get()} - {EmpNameBox.get()}"
        curEmp = connEmp.cursor()

        curEmp.execute(sqlCommand, inputs)
        connEmp.commit()
        curEmp.execute("""UPDATE LOGIN_DATA SET
        `EMPLOYEE_ID` = %s,
        `EMPLOYEE_CLASS` = %s
        
        WHERE `oid` = %s""",inputs2)
        connEmp.commit()
        
        curEmp.close()
        clearEntryEmp()
        EmpTreeView.delete(*EmpTreeView.get_children())
        queryTreeEmp()
        
        messagebox.showinfo("Update Successful", 
                            f"You Have Updated Employee {empFullName}", parent=EmpWin) 
    
    def createEmp():
        createComEmp = """INSERT INTO EMP_DATA (
        EMPLOYEE_ID, EMPLOYEE_NAME, EMPLOYEE_CLASS, NATIONALITY, 
        DOB, ADDRESS, JOIN_DATE, STATUS)
    
        VALUES (%s, %s ,%s, %s, %s, %s, %s, %s)"""
        
        def checkDateEmp(dateVar):
            if dateVar.get() == "":
                return None
            else:
                return dateVar.get()
        
        values = (EmpIDBox.get(), EmpNameBox.get(), EmpClassBox.current(), 
                  EmpNationBox.get(), checkDateEmp(EmpDOBBox),
                  EmpAddressBox.get(), checkDateEmp(EmpDateJointBox),
                  EmpStatusBox.current())
        
        values2 =(EmpIDBox.get(), EmpClassBox.current())
                 
        empFullName = f"{EmpIDBox.get()} - {EmpNameBox.get()}"
        
        curEmp = connEmp.cursor()
        curEmp.execute(createComEmp, values)
        connEmp.commit()
        
        curEmp.execute("""INSERT INTO LOGIN_DATA (
        EMPLOYEE_ID, EMPLOYEE_CLASS)
        
        VALUES (%s, %s)""", values2)
        connEmp.commit()
        curEmp.close()
        clearEntryEmp()
        EmpTreeView.delete(*EmpTreeView.get_children())
        queryTreeEmp()
        
        messagebox.showinfo("Create Successful", 
                            f"You Have Added New Employee {empFullName}", parent=EmpWin) 
    
    def deleteEmp():
        empID = EmpIDBox.get()
        empName = EmpNameBox.get()
        empFullName = f"{empID} - {empName}"
        
        selected = EmpTreeView.focus()
        sqlDelete = "DELETE FROM EMP_DATA WHERE oid = %s"
        valDelete = (selected, )
        curEmp = connEmp.cursor()
        curEmp.execute(sqlDelete, valDelete)
        connEmp.commit()
        curEmp.execute("DELETE FROM LOGIN_DATA WHERE oid = %s", valDelete)
        connEmp.commit()
        curEmp.close()
        clearEntryEmp()
        EmpTreeView.delete(*EmpTreeView.get_children())
        queryTreeEmp()
        
        messagebox.showinfo("Delete Successful", 
                            f"You Have Removed Employee {empFullName}", parent=EmpWin) 
    
    def selectEmp():
        selected = EmpTreeView.focus()
        sqlSelect = "SELECT * FROM EMP_DATA WHERE oid = %s"
        valSelect = (selected, )
        curEmp = connEmp.cursor()
        curEmp.execute(sqlSelect, valSelect)
        recLst = curEmp.fetchall()
        connEmp.commit()
        curEmp.close()
        
        clearEntryEmp()
        
        EmployeeIDRef = recLst[0][1]
    
        if EmployeeIDRef == "ADM01":
            EmpIDGenButton.config(state=DISABLED)
            buttonUpdateEmp.config(state=DISABLED)
            buttonDeleteEmp.config(state=DISABLED)
            buttonCreateEmp.config(state=DISABLED)
        
        else:
            EmpIDBox.insert(0, recLst[0][1])
            EmpNameBox.insert(0, recLst[0][2])
            EmpClassBox.current(recLst[0][3])
            
            EmpNationBox.config(state="normal")
            EmpNationBox.insert(0, recLst[0][4])
            EmpNationBox.config(state="readonly")
            
            if recLst[0][5] == None:
                EmpDOBBox.config(state="normal")
                EmpDOBBox.insert(0, "")
                EmpDOBBox.config(state="readonly")
            else:
                EmpDOBBox.config(state="normal")
                EmpDOBBox.insert(0, recLst[0][5])
                EmpDOBBox.config(state="readonly")
            
            EmpAddressBox.insert(0, recLst[0][6])
            
            if recLst[0][7] == None:
                EmpDateJointBox.config(state="normal")
                EmpDateJointBox.insert(0, "")
                EmpDateJointBox.config(state="readonly")
            else:
                EmpDateJointBox.config(state="normal")
                EmpDateJointBox.insert(0, recLst[0][7])
                EmpDateJointBox.config(state="readonly")
            
            EmpStatusBox.current(recLst[0][8])
            
            EmpIDGenButton.config(state=DISABLED)
            buttonUpdateEmp.config(state=NORMAL)
            buttonDeleteEmp.config(state=NORMAL)
            buttonCreateEmp.config(state=DISABLED)
    
    def deselectEmp():
        selected = EmpTreeView.selection()
        if len(selected) > 0:
            EmpTreeView.selection_remove(selected[0])
            clearEntryEmp()
        else:
            clearEntryEmp()
    
    def clearEntryEmp():
        EmpIDGenButton.config(state=NORMAL)
        buttonUpdateEmp.config(state=DISABLED)
        buttonDeleteEmp.config(state=DISABLED)
        buttonCreateEmp.config(state=NORMAL)
        
        EmpNameBox.delete(0, END)
        EmpClassBox.current(0)
        EmpIDBox.delete(0, END)
        EmpStatusBox.current(1)
        
        EmpNationBox.config(state="normal")
        EmpNationBox.delete(0, END)
        EmpNationBox.config(state="readonly")
        
        EmpDOBBox.config(state="normal")
        EmpDOBBox.delete(0, END)
        EmpDOBBox.config(state="readonly")
        
        EmpAddressBox.delete(0, END)
        
        EmpDateJointBox.config(state="normal")
        EmpDateJointBox.delete(0, END)
        EmpDateJointBox.config(state="readonly")
    
    def refreshEmp():
        clearEntryEmp()
        EmpTreeView.delete(*EmpTreeView.get_children())
        queryTreeEmp()
    
    def closeEmp():
        EmpWin.destroy()
    
    EmpDataFrame = LabelFrame(EmpWin, text="Record")
    EmpDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # EmpDataFrame.pack(fill="x", expand="yes", padx=20)
    
    EmpButtonFrame = LabelFrame(EmpWin, text="Command")
    EmpButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # EmpButtonFrame.pack(fill="x", expand="yes", padx=20)

    EmpNameLabel = Label(EmpDataFrame, text="Name")
    EmpClassLabel = Label(EmpDataFrame, text="Class")
    EmpIDLabel = Label(EmpDataFrame, text="Employee ID")
    EmpStatusLabel = Label(EmpDataFrame, text="Status")
    EmpNationLabel = Label(EmpDataFrame, text="Nationality")
    EmpDOBLabel = Label(EmpDataFrame, text="Date of Birth")
    EmpAddressLabel = Label(EmpDataFrame, text="Address")
    EmpDateJoinLabel = Label(EmpDataFrame, text="Date Join")

    EmpNameBox = Entry(EmpDataFrame, width=20)
    EmpClassBox = ttk.Combobox(EmpDataFrame, width=18, 
                               value=["Designer", "Purchaser", 
                                      "Project Manager", "Administrator"], state="readonly")
    EmpIDBox = Entry(EmpDataFrame, width=12)
    EmpIDGenButton = Button(EmpDataFrame, text="Gen", width=4, command=genEmpID)
    EmpStatusBox = ttk.Combobox(EmpDataFrame, width=18, 
                                value=["Inactive", "Active"], state="readonly")
    
    EmpNationBox = Entry(EmpDataFrame, width=15, state="readonly")
    EmpNationGenButton = Button(EmpDataFrame, text="List", width=4, command=genEmpNat)
    
    EmpDOBBox = Entry(EmpDataFrame, width=15, state="readonly")
    EmpDOBCalendarButton = Button(EmpDataFrame, text="Cal", width=4, command=DOBCalEmp)
    
    EmpAddressBox = Entry(EmpDataFrame, width=25)
    EmpDateJointBox = Entry(EmpDataFrame, width=15, state="readonly")
    EmpDateJoinCalendarButton = Button(EmpDataFrame, text="Cal", width=4, command=DateJoinCalEmp)
    
    EmpClassBox.current(0)
    EmpStatusBox.current(1)

    EmpNameLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    EmpClassLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    EmpIDLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    EmpStatusLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)
    EmpNationLabel.grid(row=0, column=3, padx=10, pady=5, sticky=E)
    EmpDOBLabel.grid(row=1, column=3, padx=10, pady=5, sticky=E)
    EmpAddressLabel.grid(row=2, column=3, padx=10, pady=5, sticky=E)
    EmpDateJoinLabel.grid(row=3, column=3, padx=10, pady=5, sticky=E)
    
    EmpNameBox.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky=W)
    EmpClassBox.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky=W)
    
    EmpIDBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    EmpIDGenButton.grid(row=2, column=2, padx=5, pady=5, sticky=W)
    EmpStatusBox.grid(row=3, column=1, columnspan=2, padx=10, pady=5, sticky=W)
    EmpNationBox.grid(row=0, column=4, padx=10, pady=5, sticky=W)
    EmpNationGenButton.grid(row=0, column=5, padx=5, pady=5, sticky=W)
    
    EmpDOBBox.grid(row=1, column=4, padx=10, pady=5, sticky=W)
    EmpDOBCalendarButton.grid(row=1, column=5, padx=5, pady=5, sticky=W)
    
    EmpAddressBox.grid(row=2, column=4, columnspan=2, padx=10, pady=5, sticky=W)
    EmpDateJointBox.grid(row=3, column=4, padx=10, pady=5, sticky=W)
    EmpDateJoinCalendarButton.grid(row=3, column=5, padx=5, pady=5, sticky=W)


    
    buttonUpdateEmp = Button(EmpButtonFrame, text="Update Employee Info", command=updateEmp, state=DISABLED)
    buttonUpdateEmp.grid(row=0, column=0, padx=10, pady=10, sticky=W)
            
    buttonCreateEmp = Button(EmpButtonFrame, text="Add New Employee", command=createEmp)
    buttonCreateEmp.grid(row=0, column=1, padx=10, pady=10, sticky=W)
    
    buttonDeleteEmp = Button(EmpButtonFrame, text="Delete Employee", command=deleteEmp, state=DISABLED)
    buttonDeleteEmp.grid(row=0, column=2, padx=10, pady=10, sticky=W)
    
    buttonSelectEmp = Button(EmpButtonFrame, text="Select Employee", command=selectEmp)
    buttonSelectEmp.grid(row=0, column=3, padx=10, pady=10, sticky=W)
    
    buttonDeselectEmp = Button(EmpButtonFrame, text="Deselect Employee", command=deselectEmp)
    buttonDeselectEmp.grid(row=0, column=4, padx=10, pady=10, sticky=W)
    
    buttonClearEntryEmp = Button(EmpButtonFrame, text="Clear Entry", command=clearEntryEmp)
    buttonClearEntryEmp.grid(row=0, column=5, padx=10, pady=10, sticky=W)
    
    buttonRefreshEmp = Button(EmpButtonFrame, text="Refresh", command=refreshEmp)
    buttonRefreshEmp.grid(row=0, column=6, padx=10, pady=10, sticky=W)
    
    buttonCloseEmp = Button(EmpButtonFrame, text="Close", command=closeEmp)
    buttonCloseEmp.grid(row=0, column=7, padx=10, pady=10, sticky=W)

    queryTreeEmp()