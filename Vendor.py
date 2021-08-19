from tkinter import *
from tkinter import ttk
from tkinter import filedialog 
from PIL import ImageTk, Image
from mysql import *
from mysql.connector import connect, Error
from datetime import datetime
import csv
import textwrap

from VendorModel import VendorClass
import CountryRef
import ConnConfig

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)



def RunVEND():
    def front(window):
        window.attributes('-topmost', 1)
        window.attributes('-topmost', 0)
        
    root = Toplevel()
    root.iconbitmap("MWA_Icon.ico")
    root.title("Vendor Manager")
    root.geometry("1000x600")
    root.state("zoomed")
    root.columnconfigure(0, weight=1)
    front(root)
    
    

# =============================================================================
#     connVendInit = connect(host = logininfo[0],
#                            user = logininfo[1], 
#                            password =logininfo[2])
#     
#     
#     
#     CreateVendDatabase = ["""
#                           CREATE SCHEMA IF NOT EXISTS `INDEX_VEND_MASTER` 
#                           DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci """,
#                           
#                           """
#                           CREATE TABLE IF NOT EXISTS `INDEX_VEND_MASTER`.`VENDOR_LIST`
#                           (`oid` INT AUTO_INCREMENT PRIMARY KEY,
#                            `VENDOR_NAME` VARCHAR(20),
#                            `CLASS` VARCHAR(10),
#                            `COMPANY_NAME` VARCHAR(100),
#                            `COMPANY_COUNTRY` VARCHAR(50),
#                            `COMPANY_STATE` VARCHAR(50),
#                            `COMPANY_CITY` VARCHAR(50),
#                            `POSTAL_CODE` VARCHAR(30),
#                            `COMPANY_ADDRESS_A` VARCHAR(100),
#                            `COMPANY_ADDRESS_B` VARCHAR(100),
#                            `VENDOR_DESCRIPTION` VARCHAR(100),
#                            `MAIN_CONTACT_PERSON` VARCHAR(100),
#                            `MAIN_CONTACT_NUMBER` VARCHAR(20),
#                            `MAIN_CONTACT_EMAIL` VARCHAR(100),
#                            `SECONDARY_CONTACT_PERSON` VARCHAR(100),
#                            `SECONDARY_CONTACT_NUMBER` VARCHAR(20),
#                            `SECONDARY_CONTACT_EMAIL` VARCHAR(100),
#                            `DATE_OF_ENTRY` DATETIME NOT NULL,
#                            `STATUS` INT)
#                           
#                           ENGINE = InnoDB
#                           DEFAULT CHARACTER SET = utf8mb4
#                           COLLATE = utf8mb4_0900_ai_ci""",
#                           
#                           """
#                           CREATE TABLE IF NOT EXISTS `INDEX_VEND_MASTER`.`CLASS_LIST`
#                           (`oid` INT AUTO_INCREMENT PRIMARY KEY,
#                            `CLASS` VARCHAR(10),
#                            `DESCRIPTION` VARCHAR(100))
#                           
#                           ENGINE = InnoDB
#                           DEFAULT CHARACTER SET = utf8mb4
#                           COLLATE = utf8mb4_0900_ai_ci
#                           """]
# 
#     curVendInit = connVendInit.cursor()
# 
#     for command in CreateVendDatabase:
#         curVendInit.execute(command)
#         connVendInit.commit()
#     
#     curVendInit.close()
# =============================================================================
    
    
    
    connVend = connect(host = logininfo[0],
                       user = logininfo[1], 
                       password =logininfo[2],
                       database= "INDEX_VEND_MASTER")



    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background="silver",
                    rowheight=20,
                    fieldbackground="light grey")
    
    style.map("Treeview")
    
    TopLabel = Label(root, text = "Vendor List", font=("Arial", 12))
    TopLabel.grid(row=0, column=0, padx=0, pady=0, ipadx=0, ipady=10, sticky=W+E)
    
    VendTreeFrame = Frame(root)
    VendTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
    # VendTreeFrame.pack()
    
    VendTreeScroll = Scrollbar(VendTreeFrame)
    VendTreeScroll.pack(side=RIGHT, fill=Y)
    
    VendTreeView = ttk.Treeview(VendTreeFrame, yscrollcommand=VendTreeScroll.set, selectmode="browse")
    VendTreeScroll.config(command=VendTreeView.yview)
    # VendTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
    VendTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5, fill="x", expand=True)
    
    VendTreeView['columns'] = ("Vendor",
                               "Class",
                               "Country",
                               "Address",
                               "VendorDescription",
                               "POC",
                               "POCNum",
                               "POCEmail",
                               "DateOfEntry",
                               "Status")
    
    VendTreeView.column("#0",width = 0, stretch=NO)
    VendTreeView.column("Vendor",width = 100, minwidth = 50, anchor=CENTER)
    VendTreeView.column("Class",width = 60, minwidth = 50, anchor=CENTER)
    VendTreeView.column("Country",width = 80, minwidth = 50, anchor=CENTER)
    VendTreeView.column("Address",width = 310, minwidth = 50)
    VendTreeView.column("VendorDescription",width = 145, minwidth = 55)
    VendTreeView.column("POC",width = 100, minwidth = 50, anchor=CENTER)
    VendTreeView.column("POCNum",width = 100, minwidth = 50)
    VendTreeView.column("POCEmail",width = 100, minwidth = 50)
    VendTreeView.column("DateOfEntry",width = 125, minwidth = 50, anchor=CENTER)
    VendTreeView.column("Status",width = 80, minwidth = 50, anchor=CENTER)
    
    VendTreeView.heading("#0",text = "Label", anchor = CENTER)
    VendTreeView.heading("Vendor",text = "Vendor", anchor = CENTER)   
    VendTreeView.heading("Class",text = "Class", anchor = CENTER)
    VendTreeView.heading("Country",text = "Country", anchor = CENTER)
    VendTreeView.heading("Address",text = "Address", anchor = CENTER)
    VendTreeView.heading("VendorDescription",text = "Description", anchor = CENTER)
    VendTreeView.heading("POC",text = "POC", anchor = CENTER)
    VendTreeView.heading("POCNum",text = "POC No.", anchor = CENTER)
    VendTreeView.heading("POCEmail",text = "POC Email", anchor = CENTER)
    VendTreeView.heading("DateOfEntry",text = "Date of Entry", anchor = CENTER)
    VendTreeView.heading("Status",text = "Status", anchor = CENTER)  
    




    def deselectVendClick(e):
        deselectVend()
    
    def selectVendClick(e):
        selectVend()
    
    def updateVendReturn(e):
        if buttonUpdateVend["state"] == "disabled":
            messagebox.showerror("Unable to Update", 
                                 "Please Select a Vendor", parent=root) 
        else:
            updateVend()
    
    def deleteVendDel(e):
        if buttonDeleteVend["state"] == "disabled":
            messagebox.showerror("Unable to Delete", 
                                 "Please Select a Vendor", parent=root) 
        else:
            deleteVend()

    VendTreeView.bind("<Button-3>", deselectVendClick)
    VendTreeView.bind("<Double-Button-1>", selectVendClick)
    VendTreeView.bind("<Return>", updateVendReturn)
    VendTreeView.bind("<Delete>", deleteVendDel) 
    




    def countrySelected(e):
        countryVal = VendComCountryBox.get()
        if countryVal == "Singapore":
            VendComStateBox.config(state="normal")
            VendComStateBox.delete(0, END)
            VendComStateBox.config(value=[])
            VendComStateBox.config(state="readonly")
            VendComCityBox.delete(0, END)
            VendComCityBox.config(state="readonly")
            
        elif countryVal == "Others (Please Specify)":
            VendComCountryBox.config(state="normal")
            VendComCountryBox.delete(0, END)
            VendComCountryBox.config(value=[])
            VendComStateBox.config(state="normal")
            VendComStateBox.delete(0, END)
            VendComStateBox.config(value=[])
            VendComCityBox.config(state="normal")
            VendComCityBox.delete(0, END)
            
        else:
            countryDiv = CountryRef.getDivision(countryVal)
            if countryDiv == None:
                VendComStateBox.config(state="normal")
                VendComStateBox.delete(0, END)
                VendComStateBox.config(value=[])
                VendComCityBox.config(state="normal")
                VendComCityBox.delete(0, END)
            else:
                VendComStateBox.config(value=countryDiv)
                VendComStateBox.current(0)
                VendComStateBox.config(state="readonly")
                VendComCityBox.config(state="normal")
                VendComCityBox.delete(0, END)
            
    def stateSelected(e):
        stateVal = VendComStateBox.get()
        if stateVal == "Others (Please Specify)":
            VendComStateBox.config(state="normal")
            VendComStateBox.delete(0, END)
            VendComStateBox.config(value=[])
            VendComCityBox.config(state="normal")
            VendComCityBox.delete(0, END)
        else:
            pass
                
    def queryTreeVend():
        curVend = connVend.cursor()
        curVend.execute("SELECT * FROM VENDOR_LIST")
        recLst = curVend.fetchall()    

        statLst = ["Inactive", "Active"]
        
        for rec in recLst:
            addressVal = f"{rec[8]} {rec[9]} {rec[7]} {rec[6]} {rec[5]} {rec[4]}"
            VendTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                values=(rec[1], rec[2], rec[4], addressVal,
                                        rec[10], rec[11], rec[12], rec[13],
                                        rec[17], statLst[rec[18]]))
        connVend.commit()
        curVend.close()
        
        iidLstVend = VendTreeView.get_children()
        oidLstVend = list(iidLstVend)

        indexLstVend = []
        for rec in recLst:
            if str(rec[0]) in iidLstVend:
                indexLstVend.append(rec[1])
        
        numLstVend = list(range(len(indexLstVend)))
        sortIndexVend = sorted(indexLstVend)
        
        dictIndexAndOidVend = dict(zip(indexLstVend, oidLstVend))
        dictNumAndIndexVend = dict(zip(numLstVend, sortIndexVend))
        
        for i in numLstVend:
            IndexVal = dictNumAndIndexVend.get(i)
            oidVal = dictIndexAndOidVend.get(IndexVal)             
            VendTreeView.move(oidVal, "", i)
    
    def fetchClass():
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

    def updateVend():
        sqlCommand = f"""UPDATE VENDOR_LIST SET
        `VENDOR_NAME` = %s,
        `CLASS` = %s,
        `COMPANY_NAME` = %s,
        `COMPANY_COUNTRY` = %s,
        `COMPANY_STATE` = %s,
        `COMPANY_CITY` = %s,
        `POSTAL_CODE` = %s,
        `COMPANY_ADDRESS_A` = %s,
        `COMPANY_ADDRESS_B` = %s,
        `VENDOR_DESCRIPTION` = %s,
        `MAIN_CONTACT_PERSON` = %s,
        `MAIN_CONTACT_NUMBER` = %s,
        `MAIN_CONTACT_EMAIL` = %s,
        `SECONDARY_CONTACT_PERSON` = %s,
        `SECONDARY_CONTACT_NUMBER` = %s,
        `SECONDARY_CONTACT_EMAIL` = %s,
        `STATUS` = %s
        
        WHERE `oid` = %s"""
        
        selected = VendTreeView.selection()[0]
        # venName = VendVenNameBox.get()
        
        inputs = (VendVenNameBox.get(), VendClassBox.get(), VendComNameBox.get(),
                  VendComCountryBox.get(), VendComStateBox.get(), VendComCityBox.get(),
                  VendPostalCodeBox.get(), VendAddABox.get(), VendAddBBox.get(),
                  VendDescBox.get(), VendMainPOCBox.get(), VendMainPOCNumBox.get(),
                  VendMainPOCEmailBox.get(), VendSecPOCBox.get(), VendSecPOCNumBox.get(),
                  VendSecPOCEmailBox.get(), VendStatusBox.current(), selected)
        
        respUpdateVend = messagebox.askokcancel("Confirmation",
                                                "Update This Vendor?",
                                                parent=root)
        if respUpdateVend == True:
            curVend = connVend.cursor()
            curVend.execute(sqlCommand, inputs)
            connVend.commit()
            curVend.close()
            
            clearEntryVend()
            VendTreeView.delete(*VendTreeView.get_children())
            queryTreeVend()
            
            messagebox.showinfo("Update Successful", 
                                "You Have Updated This Vendor", parent=root) 
    
    def createVend():
        timeNow = datetime.now()
        formatDate = timeNow.strftime('%Y-%m-%d %H:%M:%S')

        createCom = """INSERT INTO VENDOR_LIST (
        VENDOR_NAME, CLASS, COMPANY_NAME, COMPANY_COUNTRY, COMPANY_STATE,
        COMPANY_CITY, POSTAL_CODE, COMPANY_ADDRESS_A, COMPANY_ADDRESS_B, 
        VENDOR_DESCRIPTION, MAIN_CONTACT_PERSON, MAIN_CONTACT_NUMBER,
        MAIN_CONTACT_EMAIL, SECONDARY_CONTACT_PERSON, SECONDARY_CONTACT_NUMBER,
        SECONDARY_CONTACT_EMAIL, DATE_OF_ENTRY, STATUS)
    
        VALUES (%s, %s ,%s, %s, %s, %s, 
                %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s)"""
        
        venVal = VendVenNameBox.get()
        
        values = (VendVenNameBox.get(), VendClassBox.get(), VendComNameBox.get(),
                  VendComCountryBox.get(), VendComStateBox.get(), VendComCityBox.get(),
                  VendPostalCodeBox.get(), VendAddABox.get(), VendAddBBox.get(),
                  VendDescBox.get(), VendMainPOCBox.get(), VendMainPOCNumBox.get(),
                  VendMainPOCEmailBox.get(), VendSecPOCBox.get(), VendSecPOCNumBox.get(),
                  VendSecPOCEmailBox.get(), formatDate, VendStatusBox.current())
        
        respCreateVend = messagebox.askokcancel("Confirmation",
                                                "Create Vendor?",
                                                parent=root)
        if respCreateVend == True:
            curVend = connVend.cursor()  
            try:  
                curVend.execute(createCom, values)
                connVend.commit()
                
                curVend.close()  
                clearEntryVend()
                VendTreeView.delete(*VendTreeView.get_children())
                queryTreeVend()
                
                messagebox.showinfo("Create Successful", 
                                    f"You Have Added Vendor {venVal}", parent=root) 
            except Error as e:
                messagebox.showinfo("Create Vendor", 
                                f"SQL Error while creating vendor,\n {e}", parent=root)
                curVend.close() 
             
            
    
    def deleteVend():
        selected = VendTreeView.selection()[0]
        sqlDelete = "DELETE FROM VENDOR_LIST WHERE oid = %s"
        valDelete = (selected, )
        
        respDeleteVend = messagebox.askokcancel("Confirmation",
                                                "Delete Vendor?",
                                                parent=root)
        if respDeleteVend == True:
            curVend = connVend.cursor()
            curVend.execute(sqlDelete, valDelete)
            connVend.commit()
            
            venVal = VendVenNameBox.get()
            
            clearEntryVend()
            VendTreeView.delete(*VendTreeView.get_children())
            queryTreeVend()
            
            messagebox.showinfo("Delete Successful", 
                                f"You Have Deleted Vendor {venVal}", parent=root) 
            curVend.close()
    
    def selectVend():
        selectVal = VendTreeView.selection()
        if selectVal == ():
            messagebox.showerror("No Item Selected",
                                 "Please Select an Item",
                                 parent=root)
        else:
            try:
                selected = selectVal[0]
                sqlSelect = "SELECT * FROM VENDOR_LIST WHERE oid = %s"
                valSelect = (selected, )
                curVend = connVend.cursor()
                curVend.execute(sqlSelect, valSelect)
                recLst = curVend.fetchall()
                connVend.commit()
                curVend.close()
                
                clearEntryVend()
                
                VendClassBox.config(state="normal")
                VendClassBox.delete(0, END)
                VendClassBox.insert(0, recLst[0][2])
                VendClassBox.config(state="readonly")
                
                VendStatusBox.current(recLst[0][18])
                VendVenNameBox.insert(0, recLst[0][1])
                VendComNameBox.insert(0, recLst[0][3])
                VendDescBox.insert(0, recLst[0][10])
                
                countryLst = ["Singapore", "United States", "China", "Germany", 
                              "Japan", "United Kingdom", "France", "Spain"]
                
                if recLst[0][4] not in countryLst:    
                    VendComCountryBox.config(state="normal")
                    VendComCountryBox.delete(0, END)
                    VendComCountryBox.insert(0, recLst[0][4])
                    VendComStateBox.config(state="normal")
                    VendComStateBox.delete(0, END)
                    VendComStateBox.insert(0, recLst[0][5])
                    VendComCityBox.config(state="normal")
                    VendComCityBox.delete(0, END)
                    VendComCityBox.insert(0, recLst[0][6])
                
                elif recLst[0][4] == "Singapore":
                    VendComCountryBox.config(state="normal")
                    VendComCountryBox.delete(0, END)
                    VendComCountryBox.insert(0, recLst[0][4])
                    VendComCountryBox.config(state="readonly")
                    VendComStateBox.config(state="normal")
                    VendComStateBox.config(value=[])
                    VendComStateBox.delete(0, END)
                    VendComStateBox.insert(0, recLst[0][5])
                    VendComStateBox.config(state="readonly")
                    VendComCityBox.config(state="normal")
                    VendComCityBox.delete(0, END)
                    VendComCityBox.insert(0, recLst[0][6])
                    VendComCityBox.config(state="readonly")
                
                else:
                    VendComCountryBox.config(state="normal")
                    VendComCountryBox.delete(0, END)
                    VendComCountryBox.insert(0, recLst[0][4])
                    VendComCountryBox.config(state="readonly")
                    VendComStateBox.config(state="normal")
                    VendComStateBox.delete(0, END)
                    VendComStateBox.insert(0, recLst[0][5])
                    VendComStateBox.config(state="readonly")
                    VendComCityBox.config(state="normal")
                    VendComCityBox.delete(0, END)
                    VendComCityBox.insert(0, recLst[0][6])
                
                VendPostalCodeBox.insert(0, recLst[0][7])
                VendAddABox.insert(0, recLst[0][8])
                VendAddBBox.insert(0, recLst[0][9])
                VendMainPOCBox.insert(0, recLst[0][11])
                VendMainPOCNumBox.insert(0, recLst[0][12])
                VendMainPOCEmailBox.insert(0, recLst[0][13])
                VendSecPOCBox.insert(0, recLst[0][14])
                VendSecPOCNumBox.insert(0, recLst[0][15])
                VendSecPOCEmailBox.insert(0, recLst[0][16])
                
                buttonUpdateVend.config(state=NORMAL)
                buttonDeleteVend.config(state=NORMAL)
                buttonCreateVend.config(state=DISABLED)
                VendTreeView.config(selectmode="none")
            
            except:
                VendClassBox.current(0)
                VendClassBox.config(state="readonly")
                messagebox.showerror("Error", "Please Check", parent=root)
    
    def deselectVend():
        selected = VendTreeView.selection()
        if len(selected) > 0:
            VendTreeView.selection_remove(selected[0])
            clearEntryVend()
        else:
            clearEntryVend()
    
    def clearEntryVend():
        VendClassBox.current(0)
        VendClassBox.config(state="readonly")
        VendStatusBox.current(1)
        VendVenNameBox.delete(0, END)
        VendComNameBox.delete(0, END)
        VendDescBox.delete(0, END)
        
        VendComCountryBox.config(state="normal")
        VendComCountryBox.config(value=["Singapore", "United States", "China",
                                        "Germany", "Japan", "United Kingdom",
                                        "France", "Spain", "Others (Please Specify)"])
        VendComCountryBox.delete(0, END)
        VendComCountryBox.insert(0, "Select")
        VendComCountryBox.config(state="readonly")
        VendComStateBox.config(value=["Input Country"])
        VendComStateBox.current(0)
        VendComStateBox.config(state="readonly")   
        VendComCityBox.config(state="normal")
        VendComCityBox.delete(0, END)
        
        VendPostalCodeBox.delete(0, END)
        VendAddABox.delete(0, END)
        VendAddBBox.delete(0, END)
        VendMainPOCBox.delete(0, END)
        VendMainPOCNumBox.delete(0, END)
        VendMainPOCEmailBox.delete(0, END)
        VendSecPOCBox.delete(0, END)
        VendSecPOCNumBox.delete(0, END)
        VendSecPOCEmailBox.delete(0, END)
        
        buttonUpdateVend.config(state=DISABLED)
        buttonDeleteVend.config(state=DISABLED)
        buttonCreateVend.config(state=NORMAL)
        VendTreeView.config(selectmode="browse")
        
    def ImportVendorCSV():
        fileDir = filedialog.askopenfilename(initialdir="~/Desktop",
                                              title="Select A File",
                                              filetypes=(("CSV Files", "*.csv"),
                                                        ("Any Files", "*.*")),parent = root)
        fullLst = []
        try:
            with open(f"{fileDir}",encoding = 'utf-8-sig') as f: 
                csvFile = csv.reader(f, delimiter=",")
                
                rawLst = []
                for row in csvFile:
                    rawLst.append(row)
                
                #print(rawLst)
                
                fline = rawLst[0]
                VendorAttributeDict = VendorClass.attribute_map
                VendAttrNames = [i[1] for i in VendorAttributeDict.items()]
                
                try: 
                    if 'ï»¿' in fline[0]:
                        fline[0] = fline[0].replace('ï»¿','')
                        
                    else:
                        pass
                    valid = [fieldname in VendAttrNames for fieldname in fline]
                    assert all(valid)
    
                except Exception as e:
                    print(e)
    
                for i in range(1, len(rawLst)):
                    SingleVendor = VendorClass()
                    
                    for j in range(len(rawLst[i])):
                        
                        for attr in VendorClass.attribute_map:
                            if rawLst[0][j] == VendorClass.attribute_map[attr]:
                                try: 
                                    SingleVendor[attr] = rawLst[i][j] if rawLst[i][j] else None
                                    
                                except AssertionError as e:
                                    print(e)
                                    messagebox.showerror("Vendor Import", 
                                                        f"Incorrect Vendor Imput at \n{SingleVendor['_VENDOR_NAME']} : {VendorClass.attribute_map[attr]}\n{rawLst[i][j]}", 
                                                        parent=root)
                                    return
                        
                        singleLst  = [SingleVendor[attr] for attr in VendorClass.attribute_map]
                    if singleLst != ["" for attr in VendorClass.attribute_map]:
                        fullLst.append(singleLst)
        except FileNotFoundError as e:
            messagebox.showinfo("Vendor Import", 
                    "No file selected", parent=root) 
            return
                
        sqlfields = ', '.join([VendorAttributeDict[attr] for attr in VendorClass.attribute_map])
        sqlvalues = ', '.join(['%s' for attr in VendorClass.attribute_map])

        sqlCommand = f""" REPLACE INTO `index_vend_master`.`vendor_list` ({sqlfields})
    
        VALUES ({sqlvalues})"""
        curVend = connVend.cursor()
        for singleLst in fullLst:
            curVend.execute(sqlCommand, singleLst)
            connVend.commit()

        curVend.close()
        clearEntryVend()
        VendTreeView.delete(*VendTreeView.get_children())
        queryTreeVend()
        
        messagebox.showinfo("Import Successful", 
                            "You Have Imported a Vendor List", parent=root) 
        
    def ExportVendorCSV():
        pass
        respExportUnit = messagebox.askokcancel("Confirmation", "Confirm Export as CSV?",parent = root)
        if respExportUnit == True:
            curVend = connVend.cursor()
            curVend.execute(f"SELECT * FROM `index_vend_master`.`vendor_list`")
            result = curVend.fetchall()
            connVend.commit()
            colnames = [coldetail[0] for coldetail in curVend.description]
            connVend.commit()
            curVend.close()
        
            
            VendorAttributeDict = VendorClass.attribute_map
            VendAttrNames = [i[1] for i in VendorAttributeDict.items()]
            
            if VendAttrNames == colnames[1:]:
                pass
            else:
                messagebox.showerror("Column Error", "Column name in Database have changed",parent = root)
                # Attribute in Class different from Database field name
                return
            fullLst = []
            for i in range(len(result)):
                SingleVendor = VendorClass()
                
                for j in range(len(result[i])):
                    
                    for attr in VendorClass.attribute_map:
                        if colnames[j] == VendorClass.attribute_map[attr]:
                            SingleVendor[attr] = result[i][j] if result[i][j] else None

                    
                    singleLst  = [SingleVendor[attr] for attr in VendorClass.attribute_map]
                fullLst.append(singleLst)

        f =  filedialog.asksaveasfile(mode='w',
                                      defaultextension=".csv",
                                      parent = root,
                                      initialdir="~/Desktop",
                                      initialfile = 'VendorList.csv',
                                      title = 'Save as',
                                      filetypes=[("CSV files", "*.csv"),
                                                 ("all files", "*.*")]
                                      ) 
        if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
            return
        theWriter = csv.writer(f,lineterminator='\n')
        theWriter.writerow(VendAttrNames)
        for rec in fullLst:
            theWriter.writerow(rec)
        f.close()
        
        messagebox.showinfo("Export Successful", 
                            f"You Have Successfuly Exported a Vendor List", parent=root) 
                


    def refreshVend():
        VendClassBox.config(value=fetchClass())
        clearEntryVend()
        VendTreeView.delete(*VendTreeView.get_children())
        queryTreeVend()
    
    def closeVend():
        respCloseVend = messagebox.askokcancel("Confirmation",
                                               "Close Vendor?",
                                               parent=root)
        if respCloseVend == True:
            root.destroy()

    

    VendDataFrame = LabelFrame(root, text="Record")
    VendDataFrame.grid(row=2, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # VendDataFrame.pack(fill="x", expand="yes", padx=20)
    
    VendButtonFrame = LabelFrame(root, text="Command")
    VendButtonFrame.grid(row=3, column=0, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
    # VendButtonFrame.pack(fill="x", expand="yes", padx=20)
    
    VendClassLabel = Label(VendDataFrame, text="Class", font=("Arial", 8))
    VendStatusLabel = Label(VendDataFrame, text="Status", font=("Arial", 8))
    VendVenNameLabel = Label(VendDataFrame, text="Vendor Name", font=("Arial", 8))
    VendComNameLabel = Label(VendDataFrame, text="Company Full Name", font=("Arial", 8))
    VendDescLabel = Label(VendDataFrame, text="Description", font=("Arial", 8))    
    VendComCountryLabel = Label(VendDataFrame, text="Country", font=("Arial", 8))
    VendComStateLabel = Label(VendDataFrame, text="State", font=("Arial", 8))
    VendComCityLabel = Label(VendDataFrame, text="City", font=("Arial", 8))
    VendPostalCodeLabel = Label(VendDataFrame, text="Postal Code", font=("Arial", 8))
    VendAddALabel = Label(VendDataFrame, text="Blk. & St.", font=("Arial", 8))
    VendAddBLabel = Label(VendDataFrame, text="Unit & Building", font=("Arial", 8))
    VendMainPOCLabel = Label(VendDataFrame, text="Main POC", font=("Arial", 8))
    VendMainPOCNumLabel = Label(VendDataFrame, text="Main POC No.", font=("Arial", 8))
    VendMainPOCEmailLabel = Label(VendDataFrame, text="Main POC Email", font=("Arial", 8))
    VendSecPOCLabel = Label(VendDataFrame, text="Sec. POC", font=("Arial", 8))
    VendSecPOCNumLabel = Label(VendDataFrame, text="Sec. POC No.", font=("Arial", 8))
    VendSecPOCEmailLabel = Label(VendDataFrame, text="Sec. POC Email", font=("Arial", 8))
    
    VendClassBox = ttk.Combobox(VendDataFrame, width=20, 
                                value=fetchClass(), font=("Arial", 8), state="readonly")
    VendStatusBox = ttk.Combobox(VendDataFrame, width=20, 
                                 value=["Inactive", "Active"], font=("Arial", 8), state="readonly")
    VendVenNameBox = Entry(VendDataFrame, width=35, font=("Arial", 8))
    VendComNameBox = Entry(VendDataFrame, width=35, font=("Arial", 8))
    VendDescBox = Entry(VendDataFrame, width=35, font=("Arial", 8))
    VendComCountryBox = ttk.Combobox(VendDataFrame, width=25, 
                                     value=["Singapore", "United States", "China",
                                            "Germany", "Japan", "United Kingdom",
                                            "France", "Spain", "Others (Please Specify)"], font=("Arial", 8), state="readonly")
    VendComStateBox = ttk.Combobox(VendDataFrame, width=25, 
                                     value=["Input Country"], font=("Arial", 8), state="readonly")
    VendComCityBox = Entry(VendDataFrame, width=27, font=("Arial", 8))
    VendPostalCodeBox = Entry(VendDataFrame, width=27, font=("Arial", 8))
    VendAddABox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendAddBBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendMainPOCBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendMainPOCNumBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendMainPOCEmailBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendSecPOCBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendSecPOCNumBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    VendSecPOCEmailBox = Entry(VendDataFrame, width=30, font=("Arial", 8))
    
    VendClassBox.current(0)
    VendStatusBox.current(1)
    VendComCountryBox.config(state="normal")
    VendComCountryBox.insert(0, "Select")
    VendComCountryBox.config(state="readonly")
    VendComStateBox.current(0)
    
    VendClassLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
    VendStatusLabel.grid(row=1, column=0, padx=10, pady=5, sticky=E)
    VendVenNameLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    VendComNameLabel.grid(row=3, column=0, padx=10, pady=5, sticky=E)
    VendDescLabel.grid(row=4, column=0, padx=10, pady=5, sticky=E)
    VendComCountryLabel.grid(row=0, column=2, padx=10, pady=5, sticky=E)
    VendComStateLabel.grid(row=1, column=2, padx=10, pady=5, sticky=E)
    VendComCityLabel.grid(row=2, column=2, padx=10, pady=5, sticky=E)
    VendPostalCodeLabel.grid(row=3, column=2, padx=10, pady=5, sticky=E)
    VendAddALabel.grid(row=4, column=2, padx=10, pady=5, sticky=E)
    VendAddBLabel.grid(row=5, column=2, padx=10, pady=5, sticky=E)
    VendMainPOCLabel.grid(row=0, column=4, padx=10, pady=5, sticky=E)
    VendMainPOCNumLabel.grid(row=1, column=4, padx=10, pady=5, sticky=E)
    VendMainPOCEmailLabel.grid(row=2, column=4, padx=10, pady=5, sticky=E)
    VendSecPOCLabel.grid(row=3, column=4, padx=10, pady=5, sticky=E)
    VendSecPOCNumLabel.grid(row=4, column=4, padx=10, pady=5, sticky=E)
    VendSecPOCEmailLabel.grid(row=5, column=4, padx=10, pady=5, sticky=E)

    VendClassBox.grid(row=0, column=1, padx=10, pady=5, sticky=W)
    VendStatusBox.grid(row=1, column=1, padx=10, pady=5, sticky=W)
    VendVenNameBox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
    VendComNameBox.grid(row=3, column=1, padx=10, pady=5, sticky=W)
    VendDescBox.grid(row=4, column=1, padx=10, pady=5, sticky=W)
    VendComCountryBox.grid(row=0, column=3, padx=10, pady=5, sticky=W)
    VendComStateBox.grid(row=1, column=3, padx=10, pady=5, sticky=W)
    VendComCityBox.grid(row=2, column=3, padx=10, pady=5, sticky=W)
    VendPostalCodeBox.grid(row=3, column=3, padx=10, pady=5, sticky=W)
    VendAddABox.grid(row=4, column=3, padx=10, pady=5, sticky=W)
    VendAddBBox.grid(row=5, column=3, padx=10, pady=5, sticky=W)
    VendMainPOCBox.grid(row=0, column=5, padx=10, pady=5, sticky=W)
    VendMainPOCNumBox.grid(row=1, column=5, padx=10, pady=5, sticky=W)
    VendMainPOCEmailBox.grid(row=2, column=5, padx=10, pady=5, sticky=W)
    VendSecPOCBox.grid(row=3, column=5, padx=10, pady=5, sticky=W)
    VendSecPOCNumBox.grid(row=4, column=5, padx=10, pady=5, sticky=W)
    VendSecPOCEmailBox.grid(row=5, column=5, padx=10, pady=5, sticky=W)
    




    VendComCountryBox.bind("<<ComboboxSelected>>", countrySelected)
    VendComStateBox.bind("<<ComboboxSelected>>", stateSelected)
    




    
    buttonUpdateVend = Button(VendButtonFrame, text="Update Vendor Info", command=updateVend, state=DISABLED)
    buttonUpdateVend.grid(row=0, column=0, padx=10, pady=10, sticky=W)
            
    buttonCreateVend = Button(VendButtonFrame, text="Add New Vendor", command=createVend)
    buttonCreateVend.grid(row=0, column=1, padx=10, pady=10, sticky=W)
    
    buttonDeleteVend = Button(VendButtonFrame, text="Delete Vendor", command=deleteVend, state=DISABLED)
    buttonDeleteVend.grid(row=0, column=2, padx=10, pady=10, sticky=W)
    
    buttonSelectVend = Button(VendButtonFrame, text="Select Vendor", command=selectVend)
    buttonSelectVend.grid(row=0, column=3, padx=10, pady=10, sticky=W)
    
    buttonDeselectVend = Button(VendButtonFrame, text="Deselect Vendor", command=deselectVend)
    buttonDeselectVend.grid(row=0, column=4, padx=10, pady=10, sticky=W)
    
    buttonClearEntryVend = Button(VendButtonFrame, text="Clear Entry", command=clearEntryVend)
    buttonClearEntryVend.grid(row=0, column=5, padx=10, pady=10, sticky=W)
    
    buttonImportVend = Button(VendButtonFrame, text="Import CSV", command=ImportVendorCSV)
    buttonImportVend.grid(row=0, column=5, padx=10, pady=10, sticky=W)
    
    buttonExportVend = Button(VendButtonFrame, text="Export CSV", command=ExportVendorCSV)
    buttonExportVend.grid(row=0, column=6, padx=10, pady=10, sticky=W)
    
    buttonRefreshVend = Button(VendButtonFrame, text="Refresh", command=refreshVend)
    buttonRefreshVend.grid(row=0, column=7, padx=10, pady=10, sticky=W)
    
    buttonCloseVend = Button(VendButtonFrame, text="Close", command=closeVend)
    buttonCloseVend.grid(row=0, column=8, padx=10, pady=10, sticky=W)   
    
    queryTreeVend()





    def editClass():
        classWin = Toplevel()
        classWin.title("Class Manager")
        classWin.geometry("800x420")
            
        ClassTopLabel = Label(classWin, text = "Class Management", font=("Arial", 12))
        ClassTopLabel.grid(row=0, column=0, padx=0, pady=0, ipadx=0, ipady=10, sticky=W+E)
        
        ClassTreeFrame = Frame(classWin)
        ClassTreeFrame.rowconfigure(1, weight=1)
        ClassTreeFrame.grid(row=1, column=0, padx=10, pady=0, ipadx=10, ipady=5, sticky=W+E)
        # ClassTreeFrame.pack()
        
        ClassTreeScroll = Scrollbar(ClassTreeFrame)
        ClassTreeScroll.pack(side=RIGHT, fill=Y)
        
        ClassTreeView = ttk.Treeview(ClassTreeFrame, yscrollcommand=ClassTreeScroll.set, 
                                     height=15, selectmode="browse")
        ClassTreeScroll.config(command=ClassTreeView.yview)
        # ClassTreeView.grid(row=0, column=0, columnspan=1, padx=5, pady=5, ipadx=5, ipady=5)
        ClassTreeView.pack(padx=5, pady=5, ipadx=5, ipady=5)
        
        ClassTreeView['columns'] = ("ClassName",
                                   "Description")
        
        ClassTreeView.column("#0",width = 0, stretch=NO)
        ClassTreeView.column("ClassName",width = 80, minwidth = 50)
        ClassTreeView.column("Description",width = 200, minwidth = 50)

        ClassTreeView.heading("#0",text = "Label", anchor = CENTER)
        ClassTreeView.heading("ClassName",text = "Class", anchor = CENTER)   
        ClassTreeView.heading("Description",text = "Description", anchor = CENTER)
    




        def deselectClassClick(e):
            deselectClass()
        
        def selectClassClick(e):
            selectClass()
        
        def updateClassReturn(e):
            if buttonUpdateClass["state"] == "disabled":
                messagebox.showerror("Unable to Update", 
                                     "Please Select a Class", parent=classWin) 
            else:
                updateClass()
        
        def deleteClassDel(e):
            if buttonDeleteClass["state"] == "disabled":
                messagebox.showerror("Unable to Delete", 
                                     "Please Select a Class", parent=classWin) 
            else:
                deleteClass()
    
        ClassTreeView.bind("<Button-3>", deselectClassClick)
        ClassTreeView.bind("<Double-Button-1>", selectClassClick)
        ClassTreeView.bind("<Return>", updateClassReturn)
        ClassTreeView.bind("<Delete>", deleteClassDel) 





        def queryTreeClass():
            curVend = connVend.cursor()
            curVend.execute("SELECT * FROM CLASS_LIST")
            classRec = curVend.fetchall()    
    
            for rec in classRec:
                ClassTreeView.insert(parent="", index=END, iid=rec[0], text=rec[0], 
                                    values=(rec[1], rec[2]))
            connVend.commit()
            curVend.close()
    
        def updateClass():
            sqlComUpdateCls = f"""UPDATE CLASS_LIST SET
            `CLASS` = %s,
            `DESCRIPTION` = %s
            
            WHERE `oid` = %s"""
            
            selected = ClassTreeView.selection()[0]
            
            inputs = (ClassNameBox.get(), ClassDescBox.get(), selected)
            
            respUpdateCls = messagebox.askokcancel("Confirmation",
                                                   "Update Class?",
                                                   parent=classWin)
            if respUpdateCls == True:
                curVend = connVend.cursor()        
                curVend.execute(sqlComUpdateCls, inputs)
                connVend.commit()
                
                curVend.close()
                clearEntryClass()
                ClassTreeView.delete(*ClassTreeView.get_children())
                queryTreeClass()
            
                VendClassBox.config(value=fetchClass())
                VendClassBox.current(0)
                
                messagebox.showinfo("Update Successful", 
                                    f"You Have Updated Class {ClassNameBox.get()}", 
                                    parent=classWin) 
        
        def createClass():
            createComClass = """INSERT INTO CLASS_LIST (
            CLASS, DESCRIPTION)
        
            VALUES (%s, %s)"""
            
            values = (ClassNameBox.get(), ClassDescBox.get())
            
            respCreateCls = messagebox.askokcancel("Confirmation",
                                                   "Create Class?",
                                                   parent=classWin)
            if respCreateCls == True:            
                curVend = connVend.cursor()
                curVend.execute(createComClass, values)
                connVend.commit()
                curVend.close()
                
                clearEntryClass()
                ClassTreeView.delete(*ClassTreeView.get_children())
                queryTreeClass()
                
                VendClassBox.config(value=fetchClass())
                VendClassBox.current(0)
                
                messagebox.showinfo("Create Successful", 
                                    f"You Have Created New Class {ClassNameBox.get()}", 
                                    parent=classWin)
        
        def deleteClass():
            className = ClassNameBox.get()
            selected = ClassTreeView.selection()[0]
            sqlDelete = "DELETE FROM CLASS_LIST WHERE oid = %s"
            valDelete = (selected, )
            
            respDeleteCls = messagebox.askokcancel("Confirmation",
                                                   "Delete Class?",
                                                   parent=classWin)
            if respDeleteCls == True:            
                curVend = connVend.cursor()
                curVend.execute(sqlDelete, valDelete)
                connVend.commit()
                curVend.close()
                
                clearEntryClass()
                ClassTreeView.delete(*ClassTreeView.get_children())
                queryTreeClass()
                
                VendClassBox.config(value=fetchClass())
                VendClassBox.current(0)
                
                messagebox.showinfo("Delete Successful", 
                                    f"You Have Removed Class {className}", parent=classWin) 
        
        def selectClass():
            selectVal = ClassTreeView.selection()
            if selectVal == ():
                messagebox.showerror("No Item Selected",
                                     "Please Select an Item",
                                     parent=classWin)
            else:
                try:
                    selected = selectVal[0]
                    sqlSelect = "SELECT * FROM CLASS_LIST WHERE oid = %s"
                    valSelect = (selected, )
                    curVend = connVend.cursor()
                    curVend.execute(sqlSelect, valSelect)
                    recLst = curVend.fetchall()
                    connVend.commit()
                    curVend.close()
                    clearEntryClass()
                    
                    ClassNameBox.insert(0, recLst[0][1])
                    ClassDescBox.insert(0, recLst[0][2])
                
                    buttonUpdateClass.config(state=NORMAL)
                    buttonDeleteClass.config(state=NORMAL)
                    buttonCreateClass.config(state=DISABLED)
                    ClassTreeView.config(selectmode="none")
                except:
                    messagebox.showerror("Error", "Please Check", parent=classWin)
        
        def deselectClass():
            selected = ClassTreeView.selection()
            if len(selected) > 0:
                ClassTreeView.selection_remove(selected[0])
                clearEntryClass()
            else:
                clearEntryClass()
        
        def clearEntryClass():
            ClassNameBox.delete(0, END)
            ClassDescBox.delete(0, END)
            
            buttonUpdateClass.config(state=DISABLED)
            buttonDeleteClass.config(state=DISABLED)
            buttonCreateClass.config(state=NORMAL)
            ClassTreeView.config(selectmode="browse")
        
        def closeClass():
            respCloseCls = messagebox.askokcancel("Confirmation",
                                                  "Close Class Window?",
                                                  parent=classWin)
            if respCloseCls == True:   
                classWin.destroy()

        



        ClassDataFrame = LabelFrame(classWin, text="Record")
        ClassDataFrame.grid(row=1, column=1, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
        # ClassDataFrame.pack(fill="x", expand="yes", padx=20)
        
        ClassButtonFrame = LabelFrame(classWin, text="Command")
        ClassButtonFrame.grid(row=1, column=2, padx=20, pady=5, ipadx=5, ipady=5, sticky=W+E)
        # ClassButtonFrame.pack(fill="x", expand="yes", padx=20)
    
    
        ClassNameLabel = Label(ClassDataFrame, text="Class (Short Form)")
        ClassDescLabel = Label(ClassDataFrame, text="Description")
    
        ClassNameBox = Entry(ClassDataFrame, width=20)
        ClassDescBox = Entry(ClassDataFrame, width=20)
    
        ClassNameLabel.grid(row=0, column=0, padx=10, pady=5, sticky=E)
        ClassDescLabel.grid(row=2, column=0, padx=10, pady=5, sticky=E)
    
        ClassNameBox.grid(row=1, column=0, padx=10, pady=5, sticky=W)
        ClassDescBox.grid(row=3, column=0, padx=10, pady=5, sticky=W)        

        
    
        buttonUpdateClass = Button(ClassButtonFrame, text="Update Class Info", width=20, command=updateClass, state=DISABLED)
        buttonUpdateClass.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky=W)
                
        buttonCreateClass = Button(ClassButtonFrame, text="Add New Class", width=20, command=createClass)
        buttonCreateClass.grid(row=1, column=0, padx=(10,5), pady=5, sticky=W)
        
        buttonDeleteClass = Button(ClassButtonFrame, text="Delete Class", width=20, command=deleteClass, state=DISABLED)
        buttonDeleteClass.grid(row=2, column=0, padx=(10,5), pady=5, sticky=W)
        
        buttonSelectClass = Button(ClassButtonFrame, text="Select Class", width=20, command=selectClass)
        buttonSelectClass.grid(row=3, column=0, padx=(10,5), pady=5, sticky=W)
        
        buttonDeselectClass = Button(ClassButtonFrame, text="Deselect Class", width=20, command=deselectClass)
        buttonDeselectClass.grid(row=4, column=0, padx=(10,5), pady=5, sticky=W)
        
        buttonClearEntryClass = Button(ClassButtonFrame, text="Clear Entry", width=20, command=clearEntryClass)
        buttonClearEntryClass.grid(row=5, column=0, padx=(10,5), pady=5, sticky=W)
        
        buttonCloseClass = Button(ClassButtonFrame, text="Close", width=20, command=closeClass)
        buttonCloseClass.grid(row=6, column=0, padx=(10,5), pady=(5,10), sticky=W)   
    
        queryTreeClass()
    
    
    
    
    
    menuVend = Menu(root)
    root.config(menu=menuVend)
    fileMenu = Menu(menuVend, tearoff=0)
    menuVend.add_cascade(label="File", menu=fileMenu)
    fileMenu.add_command(label="Edit Class Info", command=editClass)
    fileMenu.add_separator()
    fileMenu.add_command(label="Exit", command=root.destroy)
    

    
















