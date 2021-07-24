from tkinter import *
from tkinter import ttk
from tkinter import messagebox 
from tkcalendar import *
from datetime import datetime
from mysql import *
import mysql.connector
import ConnConfig

logininfo = (ConnConfig.host, ConnConfig.username, ConnConfig.password)

def openStock():
    StockWin = Toplevel()
    StockWin.title("Stock Inventory")
    StockWin.state("zoomed")
    StockWin.columnconfigure(0, weight=1)
    
    connInit = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2])
    
    SetupCommand = ["""CREATE SCHEMA IF NOT EXISTS `STOCK_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci""",
                    
                    """
                    CREATE TABLE IF NOT EXISTS `STOCK_MASTER`.`STOCK_LIST`
                    (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                     `PartNum` VARCHAR(50),
                     `DescStock` VARCHAR(200),
                     `QtyStock` INT,
                     `PurOrderNum` VARCHAR(50),
                     `PurDate` DATE,
                     `RcvDate` DATE,
                     `Remark` VARCHAR(100))
                    
                    ENGINE = InnoDB
                    DEFAULT CHARACTER SET = utf8mb4
                    COLLATE = utf8mb4_0900_ai_ci"""]

    curInit = connInit.cursor()
    for com in SetupCommand:
        curInit.execute(com)
        connInit.commit()
    
    curInit.close()
    
    connStock = mysql.connector.connect(host = logininfo[0],
                                        user = logininfo[1], 
                                        password =logininfo[2],
                                        database= "STOCK_MASTER")


    
    
    
    StockTitleLabel = Label(StockWin, text="Stock Inventory List", font=("Arial", 12))
    StockTitleLabel.grid(row=0, column=0, columnspan=2, padx=0, pady=0, ipadx=0, ipady=10, sticky=W+E)
    
    StockTreeFrame = Frame(StockWin)
    StockTreeFrame.grid(row=1, column=0, columnspan=2, padx=10, pady=0, ipadx=10, ipady=5, sticky="EW")
    # StockTreeFrame.pack(fill="x", expand=True)
    
    StockTreeScroll = Scrollbar(StockTreeFrame)
    StockTreeScroll.pack(side=RIGHT, fill=Y)
    
    StockTreeView = ttk.Treeview(StockTreeFrame, yscrollcommand=StockTreeScroll.set, selectmode="browse")
    StockTreeScroll.config(command=StockTreeView.yview)
    # StockTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
    StockTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
    
    StockTreeView['columns'] = ("PartNum", 
                                "Description", 
                                "InStock", 
                                "PurOrderNum",
                                "PurDate",
                                "RcvDate",
                                "Remark")
    
    # StockTreeView.column("#0", anchor = CENTER, width =50, minwidth = 0)
    StockTreeView.column("#0",  width=0, stretch=NO)
    StockTreeView.column("PartNum", anchor = CENTER, width = 160, minwidth = 50)
    StockTreeView.column("Description", anchor = W, width= 200, minwidth = 50)
    StockTreeView.column("InStock", anchor = CENTER, width = 80, minwidth = 50)
    StockTreeView.column("PurOrderNum", anchor = CENTER, width = 160, minwidth = 50)
    StockTreeView.column("PurDate", anchor = CENTER, width = 100, minwidth = 50)
    StockTreeView.column("RcvDate", anchor = CENTER, width = 100, minwidth = 50)
    StockTreeView.column("Remark", anchor = W, width = 140, minwidth = 50)
    
    StockTreeView.heading("#0", text = "Index")
    StockTreeView.heading("PartNum", text = "Part No.")
    StockTreeView.heading("Description", text = "Description")
    StockTreeView.heading("InStock", text = "Qty. In Stock")
    StockTreeView.heading("PurOrderNum", text = "Purchase Order No.")
    StockTreeView.heading("PurDate", text = "Purchased On")
    StockTreeView.heading("RcvDate", text = "Received On")
    StockTreeView.heading("Remark", text = "Remark")
    
    
    
    
    
    def deselectStockClick(e):
        deselectStock()
    
    def selectStockClick(e):
        selectStock()
    
    def updateStockReturn(e):
        if buttonUpdateStock["state"] == "disabled":
            messagebox.showwarning("Unable to Update", 
                                   "Please Select a Part", parent=StockWin) 
        else:
            updateStock()
    
    def deleteStockDel(e):
        if buttonDeleteStock["state"] == "disabled":
            messagebox.showwarning("Unable to Delete", 
                                   "Please Select a Part", parent=StockWin) 
        else:
            deleteStock()

    def exitStockEsc(e):
        closeStock()

    StockTreeView.bind("<Button-3>", deselectStockClick)
    StockTreeView.bind("<Double-Button-1>", selectStockClick)
    StockTreeView.bind("<Return>", updateStockReturn)
    StockTreeView.bind("<Delete>", deleteStockDel) 
    StockTreeView.bind("<Escape>", exitStockEsc)
    
    
    
    
    
    def queryTreeStock():
        curStock = connStock.cursor()
        curStock.execute("SELECT * FROM STOCK_LIST")
        recLst = curStock.fetchall()

        def dateFormat(dateVal):
            if dateVal == None:
                return None
            else:
                return dateVal.strftime("%d-%b-%Y")

        for rec in recLst:
            StockTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                 values=(rec[1], rec[2], rec[3], rec[4],
                                         dateFormat(rec[5]), dateFormat(rec[6]),
                                         rec[7]))
        curStock.close()
    
    def updateStock():
        selected = StockTreeView.selection()[0]
        PartCode = StockTreeView.item(selected)["values"][0]
        
        updateStockCom = f"""UPDATE `STOCK_LIST` SET
        PartNum = %s,
        DescStock = %s,
        QtyStock = %s,
        PurOrderNum = %s,
        PurDate = %s,
        RcvDate = %s,
        Remark = %s
        
        WHERE oid = %s
        """

        PartNumStock = PartNumStockBox.get()
        PurOrderNumStock = PurOrderNumBox.get()
        QtyStock = QtyStockBox.get()
        DescStock = DescStockBox.get()
        RemarkStock = RemarkBox.get()
        
        try:            
            QtyInt = int(QtyStock)
            
            def checkDateMach(dateVar):
                if dateVar.get() == "":
                    return None
                else:
                    return dateVar.get()
            
            inputs = (PartNumStock, DescStock, QtyInt,
                      PurOrderNumStock, checkDateMach(PurDateBox),
                      checkDateMach(RcvDateBox), RemarkStock, selected)
            
            respUpdateStock = messagebox.askokcancel("Confirmation", 
                                                     f"Confirm Update {PartCode}?", 
                                                     parent=StockWin)
            if respUpdateStock == True:
                curStock = connStock.cursor()
                curStock.execute(updateStockCom, inputs)
                connStock.commit()
                curStock.close()
                
                clearEntryStock()
                StockTreeView.delete(*StockTreeView.get_children())
                queryTreeStock()
                
                messagebox.showinfo("Update Successful", 
                                    f"You Have Updated {PartCode}", parent=StockWin) 
            else:
                pass
        except:
            messagebox.showerror("Error",
                                  "Please Check Format",
                                  parent=StockWin)

    def createStock():
        PartNumStock = PartNumStockBox.get()
        PurOrderNumStock = PurOrderNumBox.get()
        QtyStock = QtyStockBox.get()
        DescStock = DescStockBox.get()
        RemarkStock = RemarkBox.get()

        try:
            QtyInt = int(QtyStock)
            respCreateStock = messagebox.askokcancel("Confirmation",
                                                     "Add New Part?",
                                                     parent=StockWin)
            
            if respCreateStock == True:
                createStock = f"""INSERT INTO `STOCK_LIST` (
                PartNum, DescStock, QtyStock, PurOrderNum, PurDate, RcvDate, Remark)
                
                VALUES (%s, %s, %s, %s, %s, %s, %s)"""
                
                def checkDateMach(dateVar):
                    if dateVar.get() == "":
                        return None
                    else:
                        return dateVar.get()
                
                valueStock = (PartNumStock, DescStock, QtyInt,
                              PurOrderNumStock, checkDateMach(PurDateBox),
                              checkDateMach(RcvDateBox), RemarkStock)
                
                curStock = connStock.cursor()
                curStock.execute(createStock, valueStock)
                connStock.commit()
                curStock.close()
                
                clearEntryStock()
                StockTreeView.delete(*StockTreeView.get_children())
                queryTreeStock()
                
                messagebox.showinfo("Create Successful", 
                                    f"You Have Added New Part {PartNumStock}", 
                                    parent=StockWin) 
            else:
                pass
        except:
            messagebox.showerror("Error",
                                 "Please Check Format",
                                 parent=StockWin)
    
    def deleteStock():
        selected = StockTreeView.selection()[0]
        PartCode = StockTreeView.item(selected)["values"][0]
        
        respDeleteStock = messagebox.askokcancel("Confirmation",
                                                 f"Delete {PartCode}?",
                                                 parent=StockWin)
        
        if respDeleteStock == True:
            sqlDelete = "DELETE FROM STOCK_LIST WHERE oid = %s"
            valDelete = (selected, )
            
            curStock = connStock.cursor()
            curStock.execute(sqlDelete, valDelete)
            connStock.commit()
            
            clearEntryStock()
            StockTreeView.delete(*StockTreeView.get_children())
            queryTreeStock()
            
            messagebox.showinfo("Delete Successful",
                                f"You Have Deleted {PartCode}", parent=StockWin) 
            curStock.close()
        else:
            pass
    
    def selectStock():
        selectedTup = StockTreeView.selection()
        
        if selectedTup == ():
            messagebox.showerror("Error",
                                 "Please Select an Item",
                                 parent=StockWin)
        
        else:
            selected = selectedTup[0]        
            sqlSelect = f"SELECT * FROM STOCK_LIST WHERE oid = %s"
            valSelect = (selected, )
            
            curStock = connStock.cursor()
            curStock.execute(sqlSelect, valSelect)
            recLst = curStock.fetchall()
            connStock.commit()
            curStock.close()
            
            clearEntryStock()
            
            PartNumStockBox.insert(0, recLst[0][1])
            PurOrderNumBox.insert(0, recLst[0][4])
            QtyStockBox.delete(0, END)
            QtyStockBox.insert(0, recLst[0][3])
            
            if recLst[0][5] == None:
                PurDateBox.config(state="normal")
                PurDateBox.insert(0, "")
                PurDateBox.config(state="readonly")
            else:
                PurDateBox.config(state="normal")
                PurDateBox.insert(0, recLst[0][5])
                PurDateBox.config(state="readonly")
                
            if recLst[0][6] == None:
                RcvDateBox.config(state="normal")
                RcvDateBox.insert(0, "")
                RcvDateBox.config(state="readonly")
            else:
                RcvDateBox.config(state="normal")
                RcvDateBox.insert(0, recLst[0][6])
                RcvDateBox.config(state="readonly")
            
            DescStockBox.insert(0, recLst[0][2])
            RemarkBox.insert(0, recLst[0][7])
            
            buttonUpdateStock.config(state=NORMAL)
            buttonCreateStock.config(state=DISABLED)
            buttonDeleteStock.config(state=NORMAL)
            StockTreeView.config(selectmode="none")
    
    def deselectStock():
        selected = StockTreeView.selection()
        if len(selected) > 0:
            StockTreeView.selection_remove(selected[0])
            clearEntryStock()
        else:
            clearEntryStock()
    
    def clearEntryStock():
        buttonUpdateStock.config(state=DISABLED)
        buttonDeleteStock.config(state=DISABLED)
        buttonCreateStock.config(state=NORMAL)
        StockTreeView.config(selectmode="browse")
        
        PartNumStockBox.delete(0, END)
        PurOrderNumBox.delete(0, END)
        
        QtyStockBox.delete(0, END)
        QtyStockBox.insert(0, 1)
        
        PurDateBox.config(state="normal")
        PurDateBox.delete(0, END)
        PurDateBox.config(state="readonly")
        RcvDateBox.config(state="normal")
        RcvDateBox.delete(0, END)
        RcvDateBox.config(state="readonly")
        
        DescStockBox.delete(0, END)
        RemarkBox.delete(0, END)

    def refreshStock():
        clearEntryStock()
        StockTreeView.delete(*StockTreeView.get_children())
        queryTreeStock()
    
    def closeStock():
        respCloseStock = messagebox.askokcancel("Confirmation",
                                                "Close Stock Window?",
                                                parent=StockWin)
        if respCloseStock == True:
            StockWin.destroy()
        else:
            pass




    
    StockDataFrame = LabelFrame(StockWin, text="Record")
    StockDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # StockDataFrame.pack(fill="x", expand="yes", padx=20)
    
    StockButtonFrame = LabelFrame(StockWin, text="Command")
    StockButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # StockButtonFrame.pack(fill="x", expand="yes", padx=20)
    

    
    def purCalPro():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("270x260")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            PurDateBox.config(state="normal")
            PurDateBox.delete(0, END)
            PurDateBox.insert(0, val)
            PurDateBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            PurDateBox.config(state="normal")
            PurDateBox.delete(0, END)
            PurDateBox.config(state="readonly")
            calWin.destroy()
        
        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)
        
    def rcvCalPro():
        calWin = Toplevel()
        calWin.title("Select the Date")
        calWin.geometry("270x260")
        
        cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
        cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
        
        def confirmDate():
            val = cal.get_date()
            RcvDateBox.config(state="normal")
            RcvDateBox.delete(0, END)
            RcvDateBox.insert(0, val)
            RcvDateBox.config(state="readonly")
            calWin.destroy()
        
        def emptyDate():
            RcvDateBox.config(state="normal")
            RcvDateBox.delete(0, END)
            RcvDateBox.config(state="readonly")
            calWin.destroy()
        
        buttonConfirm = Button(calWin, text="Confirm", command=confirmDate)
        buttonConfirm.grid(row=1, column=0, padx=5, pady=5)
    
        buttonEmpty = Button(calWin, text="Remove Date", command=emptyDate)
        buttonEmpty.grid(row=1, column=1, padx=5, pady=5)
    
        buttonClose = Button(calWin, text="Close", command=calWin.destroy)
        buttonClose.grid(row=1, column=2, padx=5, pady=5)




    
    PartNumStockLabel = Label(StockDataFrame, text="Part Number")
    PurOrderNumLabel = Label(StockDataFrame, text="Purchase Order No.")
    QtyStockLabel = Label(StockDataFrame, text="Qty. In Stock")
    PurDateLabel = Label(StockDataFrame, text="Purchase Date")
    RcvDateLabel = Label(StockDataFrame, text="Receive Date")
    DescStockLabel = Label(StockDataFrame, text="Description")
    RemarkLabel = Label(StockDataFrame, text="Remark")

    PartNumStockBox = Entry(StockDataFrame, width=20)
    PurOrderNumBox = Entry(StockDataFrame, width=20)
    
    QtyStockBox = Spinbox(StockDataFrame, from_=0, to=5000, width=10)
    PurDateBox = Entry(StockDataFrame, width=12, state="readonly")
    PurDateCalendarButton = Button(StockDataFrame, text="Cal", width=5, 
                                   command=purCalPro)
    RcvDateBox = Entry(StockDataFrame, width=12, state="readonly")
    RcvDateCalendarButton = Button(StockDataFrame, text="Cal", width=5, 
                                   command=rcvCalPro)
    
    DescStockBox = Entry(StockDataFrame, width=30)
    RemarkBox = Entry(StockDataFrame, width=30)



    PartNumStockLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    PurOrderNumLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    QtyStockLabel.grid(row=0, column=2, padx=10, pady=5, sticky=E)
    PurDateLabel.grid(row=1, column=2, padx=10, pady=5, sticky=E)
    RcvDateLabel.grid(row=2, column=2, padx=10, pady=5, sticky=E)
    DescStockLabel.grid(row=0, column=5, padx=10, pady=5, sticky=E)
    RemarkLabel.grid(row=1, column=5, padx=10, pady=5, sticky=E)

    PartNumStockBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    PurOrderNumBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    
    QtyStockBox.grid(row=0, column=3, padx=10, pady=5, sticky=W)
    PurDateBox.grid(row=1, column=3, padx=10, pady=5, sticky=W)
    PurDateCalendarButton.grid(row=1, column=4, padx=(0,10), pady=5, sticky=W)
    RcvDateBox.grid(row=2, column=3, padx=10, pady=5, sticky=W)
    RcvDateCalendarButton.grid(row=2, column=4, padx=(0,10), pady=5, sticky=W)
    
    DescStockBox.grid(row=0, column=6, padx=10, pady=5, sticky=W)
    RemarkBox.grid(row=1, column=6, padx=10, pady=5, sticky=W)
    
    QtyStockBox.delete(0, END)
    QtyStockBox.insert(0, 1)




    
    buttonUpdateStock = Button(StockButtonFrame, text="Update Part", command=updateStock, state=DISABLED)
    buttonUpdateStock.grid(row=0, column=0, padx=10, pady=10, sticky=W)
            
    buttonCreateStock = Button(StockButtonFrame, text="Create Part", command=createStock)
    buttonCreateStock.grid(row=0, column=1, padx=10, pady=10, sticky=W)
    
    buttonDeleteStock = Button(StockButtonFrame, text="Delete Part", command=deleteStock, state=DISABLED)
    buttonDeleteStock.grid(row=0, column=2, padx=10, pady=10, sticky=W)
    
    buttonSelectStock = Button(StockButtonFrame, text="Select Part", command=selectStock)
    buttonSelectStock.grid(row=0, column=3, padx=10, pady=10, sticky=W)
    
    buttonDeselectStock = Button(StockButtonFrame, text="Deselect Part", command=deselectStock)
    buttonDeselectStock.grid(row=0, column=4, padx=10, pady=10, sticky=W)
    
    buttonClearEntryStock = Button(StockButtonFrame, text="Clear Entry", command=clearEntryStock)
    buttonClearEntryStock.grid(row=0, column=5, padx=10, pady=10, sticky=W)
    
    buttonRefreshStock = Button(StockButtonFrame, text="Refresh", command=refreshStock)
    buttonRefreshStock.grid(row=0, column=6, padx=10, pady=10, sticky=W)
    
    buttonCloseStock = Button(StockButtonFrame, text="Close", command=closeStock)
    buttonCloseStock.grid(row=0, column=7, padx=10, pady=10, sticky=W)
    
    queryTreeStock()
    
    
    
    
    
    
    
    