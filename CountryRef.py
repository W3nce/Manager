from tkinter import *
from tkinter import ttk
from tkinter import messagebox 
from tkcalendar import *
from datetime import datetime
from mysql import *
import mysql.connector
import ConnConfig

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

USALst = ["Alaska", "Alabama", "Arkansas", "American Samoa", "Arizona", 
          "California", "Colorado", "Connecticut", "District of Columbia", 
          "Delaware", "Florida", "Georgia", "Guam", "Hawaii", "Iowa", "Idaho", 
          "Illinois", "Indiana", "Kansas", "Kentucky", "Louisiana", "Massachusetts", 
          "Maryland", "Maine", "Michigan", "Minnesota", "Missouri", "Mississippi", 
          "Montana", "North Carolina", "North Dakota", "Nebraska", "New Hampshire", 
          "New Jersey", "New Mexico", "Nevada", "New York", "Ohio", "Oklahoma", "Oregon", 
          "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", 
          "Tennessee", "Texas", "Utah", "Virginia", "Virgin Islands", "Vermont", 
          "Washington", "Wisconsin", "West Virginia", "Wyoming", "Others (Please Specify)"]

UKLst = ["South East", "Greater London", "North West", "East of England", "West Midlands",
         "South West", "Yorkshire & Humber", "East Midlands", "North East",
         "Scotland", "Wales", "Northern Ireland", "Others (Please Specify)"]

ChinaLst = ["Beijing", "Tianjin", "Hebei", "Shanxi", "Inner Mongolia AR", "Liaoning", 
            "Jilin", "Heilongjiang", "Shanghai", "Jiangsu", "Zhejiang", "Anhui", 
            "Fujian", "Jiangxi", "Shandong", "Henan", "Hubei", "Hunan", "Guangdong", 
            "Guangxi ZAR", "Hainan", "Chongqing", "Sichuan", "Guizhou", "Yunnan",
            "Tibet AR", "Shaanxi", "Gansu", "Qinghai", "Ningxia HAR", "Xinjiang UAR",
            "Hong Kong SAR", "Macau SAR", "Taiwan Region", "Others (Please Specify)"]

GermanyLst = ["Baden-Wurttemberg", "Bavaria", "Berlin", "Brandenburg", "Bremen",
              "Hamburg", "Hesse", "Lower Saxony", "Mecklenburg-Vorpommern", 
              "North Rhine-Westphalia", "Rhineland-Palatinate", "Saarland", 
              "Saxony", "Saxony-Anhalt", "Schleswig-Holstein", "Thuringia",
              "Others (Please Specify)"]

JapanLst = ["Aichi", "Akita", "Aomori", "Chiba", "Ehime", "Fukui", "Fukuoka", "Fukushima", 
            "Gifu", "Gunma", "Hiroshima", "Hokkaido", "Hyogo", "Ibaraki", "Ishikawa", 
            "Iwate", "Kagawa", "Kagoshima", "Kanagawa", "Kochi", "Kumamoto", "Kyoto", 
            "Mie", "Miyagi", "Miyazaki", "Nagano", "Nagasaki", "Nara", "Niigata", "Oita", 
            "Okayama", "Okinawa", "Osaka", "Saga", "Saitama", "Shiga", "Shimane", "Shizuoka",
            "Tochigi", "Tokushima", "Tokyo", "Tottori", "Toyama", "Wakayama", "Yamagata", 
            "Yamaguchi", "Yamanashi", "Others (Please Specify)"]

FranceLst = ["Auvergne-Rhone-Alps", "Burgundy-Free Country", "Brittany", 
             "Centre-Loire Valley", "Corsica", "Great East", "Upper France", 
             "Metropolitan Paris", "Normandy", "New Aquitaine", "Occitania", 
             "Loire Countries", "Provence-Alps-Azure Coast", 
             "French Guadeloupe", "French Guiana", "French Reunion", 
             "French Martinique", "French Mayotte", "French Polynesia", 
             "New Caledonia", "Others (Please Specify)"]

SpainLst = ["Andalusia", "Catalonia", "Metropolitan Madrid", "Valencia", "Galicia", 
            "Castile and Leon", "Basque", "Castilla-La Mancha", "Canary Islands", 
            "Murcia", "Aragon", "Extremadura", "Balearic Islands", "Asturias", 
            "Navarre", "Cantabria", "La Rioja", "Others (Please Specify)"]

CountryDic = {"United States": USALst, "United Kingdom": UKLst, "China": ChinaLst,
              "Germany": GermanyLst, "Japan": JapanLst, "France": FranceLst,
              "Spain": SpainLst}

# CurrencyDic = {"SGD": 1.00, "USD": 1.34, "CNY": 0.21, "HKD": 0.17, "EUR": 1.60, 
#                "JPY": 0.012, "GBP": 1.87, "AUD": 1.02, "CAD": 1.09, "CHF": 1.46, }

def getDivision(country):
    return CountryDic.get(country)

def getCcyLst():
    connCcy = mysql.connector.connect(host = logininfo[0],
                                  user = logininfo[1], 
                                  password =logininfo[2],
                                  database= "CURRENCY_MASTER")
    
    curCcy = connCcy.cursor()
    curCcy.execute("SELECT CcyAcro FROM `CCY_LIST`")
    ccyTup = curCcy.fetchall()
    ccyLst = []
    for item in ccyTup:
        ccyLst.append(item[0])
    return ccyLst

def getExRate(currency):
    connCcy = mysql.connector.connect(host = logininfo[0],
                                  user = logininfo[1], 
                                  password =logininfo[2],
                                  database= "CURRENCY_MASTER")
    
    curCcy = connCcy.cursor()
    curCcy.execute(f"SELECT ExRate FROM `CCY_LIST` WHERE CcyAcro = '{currency}'")
    exRate = curCcy.fetchall()
    if exRate == []:
        messagebox.showerror('Get Exchange Rate', 'No Exchange Rate for this Currency, 0.00 returned')
        return 0.00
    else:
        return exRate[0][0]

def openCurrency():
    CcyWin = Toplevel()
    CcyWin.title("Edit Currency & Exchange Rate")
    CcyWin.geometry("600x480")
    
    connInit = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2])
    
    SetupCommand = ["""CREATE SCHEMA IF NOT EXISTS `CURRENCY_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci""",
                    
                    """
                    CREATE TABLE IF NOT EXISTS `CURRENCY_MASTER`.`CCY_LIST`
                    (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                     `CcyAcro` VARCHAR(20),
                     `CcyFullName` VARCHAR(100),
                     `ExRate` FLOAT,
                     `DateUpdated` DATE,
                     `Remark` VARCHAR(100))
                    
                    ENGINE = InnoDB
                    DEFAULT CHARACTER SET = utf8mb4
                    COLLATE = utf8mb4_0900_ai_ci"""]

    curInit = connInit.cursor()
    for com in SetupCommand:
        curInit.execute(com)
        connInit.commit()
    
    curInit.close()
    
    connCcy = mysql.connector.connect(host = logininfo[0],
                                      user = logininfo[1], 
                                      password =logininfo[2],
                                      database= "CURRENCY_MASTER")
    
    
    
    # Insert Default Value
# =============================================================================
#     curCcy = connCcy.cursor()
#     curCcy.execute("SELECT * FROM CCY_LIST")
#     existLst = curCcy.fetchall()
#     
#     if existLst == []:
#         defaultCcy = f"""INSERT INTO `CCY_LIST` (
#         CcyAcro, CcyFullName, ExRate, DateUpdated, Remark)
#         VALUES (%s, %s, %s, %s, %s)"""
#                     
#         defaultLst = [("SGD", "Singapore Dollar", 1.00, None, ""),
#                       ("USD", "United States Dollar", 1.34, None, ""),
#                       ("CNY", "Chinese Yuan (Remminbi)", 0.21, None, ""),
#                       ("HKD", "Hong Kong Dollar", 0.17, None, ""),
#                       ("EUR", "European Euro", 1.60, None, ""),
#                       ("JPY", "Japanese Yen", 0.012, None, ""),
#                       ("GBP", "British Pound Sterling", 1.87, None, ""),
#                       ("AUD", "Australian Dollar", 1.02, None, ""),
#                       ("CAD", "Canadian Dollar", 1.09, None, ""),
#                       ("CHF", "Swiss Franc", 1.46, None, "")]
#         
#         for ccy in defaultLst:
#             curCcy.execute(defaultCcy, ccy)
#         connCcy.commit()
#     curCcy.close()
# =============================================================================

    
        
    
    
    CcyTitleLabel = Label(CcyWin, text="Select Currency", font=("Arial", 12))
    CcyTitleLabel.grid(row=0, column=0, columnspan=2, padx=0, pady=(10,0), ipadx=0, ipady=0, sticky=W+E)
    
    CcyTreeFrame = Frame(CcyWin)
    CcyTreeFrame.grid(row=1, column=0, columnspan=2, padx=10, pady=0, ipadx=10, ipady=5, sticky="EW")
    # CcyTreeFrame.pack(fill="x", expand=True)
    
    CcyTreeScroll = Scrollbar(CcyTreeFrame)
    CcyTreeScroll.pack(side=RIGHT, fill=Y)
    
    CcyTreeView = ttk.Treeview(CcyTreeFrame, yscrollcommand=CcyTreeScroll.set, selectmode="browse")
    CcyTreeScroll.config(command=CcyTreeView.yview)
    # CcyTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
    CcyTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
    
    CcyTreeView['columns'] = ("CcyAcro", 
                              "CcyFullName", 
                              "ExRate", 
                              "DateUpdated",
                              "Remark")
    
    # CcyTreeView.column("#0", anchor = CENTER, width =50, minwidth = 0)
    CcyTreeView.column("#0",  width=0, stretch=NO)
    CcyTreeView.column("CcyAcro", anchor = CENTER, width = 50, minwidth = 50)
    CcyTreeView.column("CcyFullName", anchor = CENTER, width= 150, minwidth = 50)
    CcyTreeView.column("ExRate", anchor = CENTER, width = 80, minwidth = 50)
    CcyTreeView.column("DateUpdated", anchor = CENTER, width = 80, minwidth = 50)
    CcyTreeView.column("Remark", anchor = CENTER, width = 120, minwidth = 50)
    
    CcyTreeView.heading("#0", text = "Index")
    CcyTreeView.heading("CcyAcro", text = "Code")
    CcyTreeView.heading("CcyFullName", text = "Currency Name")
    CcyTreeView.heading("ExRate", text = "Ex. Rate")
    CcyTreeView.heading("DateUpdated", text = "Latest Update")
    CcyTreeView.heading("Remark", text = "Remark")
    


    def queryTreeCcy():
        curCcy = connCcy.cursor()
        curCcy.execute("SELECT * FROM CCY_LIST")
        recLst = curCcy.fetchall()    

        def dateFormat(dateVal):
            if dateVal == None:
                return None
            else:
                return dateVal.strftime("%d-%b-%Y")

        for rec in recLst:
            CcyTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                               values=(rec[1], rec[2], rec[3], 
                                       dateFormat(rec[4]), rec[5]))
        curCcy.close()
    
    def updateCcy():
        timeNow = datetime.now()
        formatDate = timeNow.strftime("%Y-%m-%d")
        
        selected = CcyTreeView.selection()[0]
        CcyCode = CcyTreeView.item(selected)["values"][0]
        
        updateCcyCom = f"""UPDATE `CCY_LIST` SET
        CcyAcro = %s,
        CcyFullName = %s,
        ExRate = %s,
        DateUpdated = %s,
        Remark = %s
        
        WHERE oid = %s
        """
        
        Acro = CcyAcroBox.get()
        CcyFullName = CcyNameBox.get()
        ExRate = ExRateBox.get()
        Remark = RemarkBox.get()
        
        inputs = (Acro, CcyFullName, ExRate, formatDate, Remark, selected)
        
        respUpdateCcy = messagebox.askokcancel("Confirmation", 
                                               f"Confirm Update {CcyCode}?", 
                                               parent=CcyWin)
        if respUpdateCcy == True:
            curCcy = connCcy.cursor()
            curCcy.execute(updateCcyCom, inputs)
            connCcy.commit()
            curCcy.close()
            
            clearEntryCcy()
            CcyTreeView.delete(*CcyTreeView.get_children())
            queryTreeCcy()
            
            messagebox.showinfo("Update Successful", 
                                f"You Have Updated {CcyCode}", parent=CcyWin) 
        else:
            pass
    
    def createCcy():
        timeNow = datetime.now()
        formatDate = timeNow.strftime("%Y-%m-%d")
        
        Acro = CcyAcroBox.get()
        CcyFullName = CcyNameBox.get()
        ExRate = ExRateBox.get()
        Remark = RemarkBox.get()
        
        try:
            ExRateFloat = float(ExRate)
            
            respCreateCcy = messagebox.askokcancel("Confirmation",
                                                   "Add New Currency?",
                                                   parent=CcyWin)
            if respCreateCcy == True:
                createCcy = f"""INSERT INTO `CCY_LIST` (
                CcyAcro, CcyFullName, ExRate, DateUpdated, Remark)
                
                VALUES (%s, %s, %s, %s, %s)"""
                
                valueCcy = (Acro, CcyFullName, 
                            ExRateFloat, formatDate, Remark)
                
                curCcy = connCcy.cursor()
                curCcy.execute(createCcy, valueCcy)
                connCcy.commit()
                curCcy.close()
                
                clearEntryCcy()
                CcyTreeView.delete(*CcyTreeView.get_children())
                queryTreeCcy()
                
                messagebox.showinfo("Create Successful", 
                                    f"You Have Added New Currency {Acro}", 
                                    parent=CcyWin) 
            else:
                pass
        except:
            messagebox.showerror("Error",
                                "Please Enter a Proper Number for Exchange Rate",
                                parent=CcyWin)
        
    def deleteCcy():
        selected = CcyTreeView.selection()[0]
        CcyCode = CcyTreeView.item(selected)["values"][0]
        
        respDeleteCcy = messagebox.askokcancel("Confirmation",
                                               f"Delete Currency {CcyCode}?",
                                               parent=CcyWin)
        
        if respDeleteCcy == True:
            sqlDelete = "DELETE FROM CCY_LIST WHERE oid = %s"
            valDelete = (selected, )
            
            curCcy = connCcy.cursor()
            curCcy.execute(sqlDelete, valDelete)
            connCcy.commit()
            
            clearEntryCcy()
            CcyTreeView.delete(*CcyTreeView.get_children())
            queryTreeCcy()
            
            messagebox.showinfo("Delete Successful", 
                                f"You Have Removed Currency {CcyCode}", parent=CcyWin) 
            curCcy.close()
        else:
            pass
    
    def selectCcy():
        selectedTup = CcyTreeView.selection()
        
        if selectedTup == ():
            messagebox.showerror("Error",
                                 "Please Select an Item",
                                 parent=CcyWin)
        
        else:
            selected = selectedTup[0]        
            sqlSelect = f"SELECT * FROM CCY_LIST WHERE oid = %s"
            valSelect = (selected, )
            
            curCcy = connCcy.cursor()
            curCcy.execute(sqlSelect, valSelect)
            recLst = curCcy.fetchall()
            connCcy.commit()
            curCcy.close()
            
            clearEntryCcy()
            
            CcyAcroBox.insert(0, recLst[0][1])
            CcyNameBox.insert(0, recLst[0][2])
            ExRateBox.insert(0, recLst[0][3])
            RemarkBox.insert(0, recLst[0][5])
            
            buttonUpdateCcy.config(state=NORMAL)
            buttonCreateCcy.config(state=DISABLED)
            buttonDeleteCcy.config(state=NORMAL)
            CcyTreeView.config(selectmode="none")
    
    def deselectCcy():
        selected = CcyTreeView.selection()
        if len(selected) > 0:
            CcyTreeView.selection_remove(selected[0])
            clearEntryCcy()
        else:
            clearEntryCcy()
    
    def clearEntryCcy():
        buttonUpdateCcy.config(state=DISABLED)
        buttonDeleteCcy.config(state=DISABLED)
        buttonCreateCcy.config(state=NORMAL)
        CcyTreeView.config(selectmode="browse")
        
        CcyAcroBox.delete(0, END)
        CcyNameBox.delete(0, END)
        ExRateBox.delete(0, END)
        RemarkBox.delete(0, END)
    
    def refreshCcy():
        clearEntryCcy()
        CcyTreeView.delete(*CcyTreeView.get_children())
        queryTreeCcy()
    

    
    CcyDataFrame = LabelFrame(CcyWin, text="Record")
    CcyDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W)
    # CcyDataFrame.pack(fill="x", expand="yes", padx=20)
    
    CcyButtonFrame = LabelFrame(CcyWin, text="Command")
    CcyButtonFrame.grid(row=2, column=1, padx=20, pady=5, ipadx=5, ipady=5, sticky=W)
    # CcyButtonFrame.pack(fill="x", expand="yes", padx=20)
    
    CcyAcroLabel = Label(CcyDataFrame, text="Currency Code")
    CcyNameLabel = Label(CcyDataFrame, text="Currency Name")
    ExRateLabel = Label(CcyDataFrame, text="Exchange Rate")
    RemarkLabel = Label(CcyDataFrame, text="Remark")

    CcyAcroBox = Entry(CcyDataFrame, width=20)
    CcyNameBox = Entry(CcyDataFrame, width=20)
    ExRateBox = Spinbox(CcyDataFrame, from_=0.00, to=1000, increment=0.01, width=18)
    RemarkBox = Entry(CcyDataFrame, width=20)



    CcyAcroLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    CcyNameLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    ExRateLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    RemarkLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)

    CcyAcroBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    CcyNameBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    ExRateBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    RemarkBox.grid(row=3, column=1, padx=10, pady=5, sticky=W)


    
    buttonUpdateCcy = Button(CcyButtonFrame, text="Update", width=10, command=updateCcy, state=DISABLED)
    buttonCreateCcy = Button(CcyButtonFrame, text="Add New", width=10, command=createCcy)
    buttonDeleteCcy = Button(CcyButtonFrame, text="Delete", width=10, command=deleteCcy, state=DISABLED)
    buttonSelectCcy = Button(CcyButtonFrame, text="Select", width=10, command=selectCcy)
    buttonDeselectCcy = Button(CcyButtonFrame, text="Deselect", width=10, command=deselectCcy)
    buttonClearEntryCcy = Button(CcyButtonFrame, text="Clear Entry", width=10, command=clearEntryCcy)
    buttonRefreshCcy = Button(CcyButtonFrame, text="Refresh", width=10, command=refreshCcy)

    buttonUpdateCcy.grid(row=0, column=0, padx=10, pady=1)
    buttonCreateCcy.grid(row=1, column=0, padx=10, pady=1)
    buttonDeleteCcy.grid(row=2, column=0, padx=10, pady=1)
    buttonSelectCcy.grid(row=0, column=1, padx=10, pady=1)
    buttonDeselectCcy.grid(row=1, column=1, padx=10, pady=1)
    buttonClearEntryCcy.grid(row=2, column=1, padx=10, pady=1)
    buttonRefreshCcy.grid(row=3, column=1, padx=10, pady=1)
    
    queryTreeCcy()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    









