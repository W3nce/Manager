from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox 
from tkcalendar import *
from datetime import datetime
from PIL import ImageTk, Image
from mysql import *
from mysql.connector import Error
import mysql.connector
import csv
import re 
from fpdf import FPDF
import ConnConfig,CountryRef
from CBTreeView import CbTreeview
from RFQemail import EmailRFQWindow as IssueEmail

from AutoCombo import AutocompleteCombobox,AutocompleteEntry 
import os

import pprint as pp
logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

connRFQ = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password = logininfo[2])

ProjList,MachineList,MakerList,EmployeeDict = [],[],[],{}

CurrentRFQ = None

curRFQ = connRFQ.cursor()
curRFQ.execute("""
               SELECT `STATUS_ID`,`STATUS_NAME` FROM `rfq_master`.`status_master`
               """)
               
StatusDict = {Status[0]:Status[1] for Status in curRFQ.fetchall()}
StatusList = list(StatusDict)


def openRFQ(tabNote,RFQ = None):
    global ProjList,MachineList,MakerList,img,EmployeeDict,CurrentRFQ,StatusDict,photo
        
    if CurrentRFQ and RFQ != CurrentRFQ:
        messagebox.showwarning("Open RFQ Error", "Please close current RFQ")
                
        tabObjDict = tabNote.children.copy()
        tab_names = ["Request for Quotation"]
        
        for Tab in tabObjDict:
            
            if tabNote.tab('.!notebook.' + Tab, "text") in tab_names:
                tabNote.select(tabObjDict[Tab])
        return
    else:
        tabObjDict = tabNote.children.copy()
        tab_names = ["Request for Quotation"]
        
        for Tab in tabObjDict:
            
            if tabNote.tab('.!notebook.' + Tab, "text") in tab_names:
                
                tabObjDict[Tab].destroy()
            
        pass
        
    CurrentRFQ = RFQ
    
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
    
    def Refresh(): 
        global CurrentRFQ
        RFQNo = CurrentRFQ
        if CurrentRFQ:
            if not UnitTreeView.TreeViewSAVED :
                if messagebox.askyesno('Refresh RFQ Entry','Are you sure to refresh this RFQ?\nAny unsaved changes will be lost',parent = RFQFrame):
                    
                    OpenCurrent(RFQNo)
                else: 
                    OpenCurrent(RFQNo)
            else:
                OpenCurrent(RFQNo)
        else:
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
            NewRFQ()

    
    def NewRFQ():   
        global ProjList,MachineList,MakerList
        _QueryMachNo()
        _QueryMaker()
        
        ProjectCombo.configure(state='active')
        MachineCombo.configure(state='active')
        CompletedCheckbutton.configure(state='disabled')
# =============================================================================
#         
#         if MakerList:
#             SelectMakerPartsButton.configure(state = 'active')
#             CreateRFQButton.configure(state = 'active')
#             IssueRFQButton.configure(state = 'active')
#         else:
#             SelectMakerPartsButton.configure(state = 'disabled')   
#             CreateRFQButton.configure(state = 'disabled') 
#             IssueRFQButton.configure(state = 'disabled') 
# =============================================================================
            
        _ShowPurchaserName()
    
    def _QueryMachNo():
        global ProjList,MachineList,MachineDict,MakerList
        MachineList = []
        MachineDict = {}
        
        curRFQ = connRFQ.cursor()
        if ProjList:
            
            curRFQ.execute(f"""
                           SELECT MACH_ID,MACH_NAME,ORDER_QTY FROM {ProjectCombo.get()}.mach_index
                           """)
                           
            MachineDict = {Machid[0]:(Machid[1],Machid[2]) for Machid in curRFQ.fetchall()}
            MachineList = list(MachineDict)
            
            
            MachineCombo['value'] = MachineList 
        
        if MachineList :
            MachineCombo.current(0)
            
        else:
            MachineCombo.set('Please create a Machine')
            MakerList = []
            
        curRFQ.close()
        
        ProjectNameLabel['text'] = ProjNameDict[ProjectCombo.get()]

    
    def _QueryMaker():
        global MakerList,MachineDict,AsmDict
        AsmDict = {}
        curRFQ = connRFQ.cursor()
        if MachineDict: 
            
            curRFQ.execute(f"""
                           SELECT `ASSEM_FULL`,`DES_QTY` FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}`
                           """)
                           
            AsmDict = {Asm[0]:Asm[1] for Asm in curRFQ.fetchall()}
        AsmList = list(AsmDict)
        MakerList = []
        if AsmList:
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
                 
        else:
            MakerCombo.set('No Assembly for this Machine')                
            
        curRFQ.close()
        MakerCombo['value'] = MakerList
                    
        MacQtyBox.configure(state="normal")
        MacQtyBox.delete(0,END)
        
        
        MacQtyBox.insert(END,MachineDict[MachineCombo.get()][1]) if MachineCombo.get() in MachineDict else None
        MachineNameLabel['text'] = MachineDict[MachineCombo.get()][0] if MachineCombo.get() in MachineDict else ''
        
        if MakerList:
            MakerCombo.current(0)
            
            SelectMakerPartsButton.configure(state = 'active')
            CreateRFQButton.configure(state = 'active')
            IssueRFQButton.configure(state = 'active')
            
        else :
            MakerCombo.set('No parts in Unit Assembly')
            
            SelectMakerPartsButton.configure(state = 'disabled')
            CreateRFQButton.configure(state = 'disabled') 
            IssueRFQButton.configure(state = 'disabled') 
            
        MacQtyBox.configure(state="readonly")
        clearTreeUnit()
        

        
    def QueryMachNo(event) :
        _QueryMachNo()   
        _QueryMaker()

    def QueryMaker(event):
        _QueryMaker()
        
            
    def SelectVendor():
        _SelectVendor(VendorBox,VendorAddressBox,VendorNameLabel)     
        
    def ChangeCurrency(event):
        _ChangeCurrency()
    
    def _ChangeCurrency():
        Ccy = CurrencyCombo.get()
        if  messagebox.askyesno('Change Currency', f'Confirm Change Currency to {Ccy}'):
            for item in UnitTreeView.get_children():
                newrow = UnitTreeView.item(item)['values'].copy()
                newrow[13] = Ccy
                UnitTreeView.item(item, values=newrow)
            UnitTreeView.CalLineTotalPrice()
            InsertReadonly(TotalSGDBox,"{:.2f}".format(round(float(UnitTreeView.TOTAL_TCURR)*float(CountryRef.getExRate(CurrencyCombo.get()))),2))
        else:
            pass
    
    def ShowPurchaserName(e):
        _ShowPurchaserName()
        
    def _ShowPurchaserName():
        if PurchaserCombo['value']:
            PurchaserNameBox.configure(state="normal")
            PurchaserNameBox.delete(0,END)
            PurchaserNameBox.insert(END,EmployeeDict[PurchaserCombo.get()])
            PurchaserNameBox.configure(state="readonly")
        else: 
            PurchaserNameBox.configure(state="normal")
            PurchaserNameBox.delete(0,END)
            PurchaserNameBox.insert(END,'Please Create an Employee')
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
                               UnitCost, Currency, TotalUnitCost,Vendor
                               FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}_{Asm}`
                               where Maker = "{_Maker}"
                               """)
            else:
                
                curRFQ.execute(f"""
                               SELECT oid, 
                               PartNum, Description, CLS,
                               Maker, Spec, V,  DES, SPA,
                               OH, REQ, PCH,  
                               UnitCost, Currency, TotalUnitCost,Vendor
                               FROM `{ProjectCombo.get()}`.`{MachineCombo.get()}_{Asm}`
                               """)
              
            
                        
            PartsList = curRFQ.fetchall()
            PartsListRFQed = []
            for rec in PartsList:
                FullPartNum = f'{ProjectCombo.get()}-{MachineCombo.get()}-{Asm}-' + rec[1] 
                if not rec[15] or rec[15] == 'None':
                    UnitTreeView.insert(parent="", index=END, iid=FullPartNum, 
                                    values=(FullPartNum, rec[2], rec[3],
                                            rec[4], rec[5], rec[6], rec[7], rec[8],
                                            rec[9], AsmDict[Asm] , rec[10], rec[11],
                                            None if rec[12] == None else "{:.2f}".format(round(float(rec[12]), 2)),
                                            rec[13], 
                                            None if rec[14] == None else "{:.2f}".format(round(float(rec[14]), 2)),''))
                else:
                    PartsListRFQed += [(FullPartNum, rec[2], rec[3],
                                            rec[4], rec[5], rec[6], rec[7], rec[8],
                                            rec[9], AsmDict[Asm] , rec[10], rec[11],
                                            None if rec[12] == None else "{:.2f}".format(round(float(rec[12]), 2)),
                                            rec[13], 
                                            None if rec[14] == None else "{:.2f}".format(round(float(rec[14]), 2)),'')]
            if PartsListRFQed:
                if messagebox.askyesno('Select Maker Parts',
                                    'Some Parts have already been RFQed\nDo you still want to add to Current RFQ?',
                                    parent = RFQFrame):
                    for rec in PartsListRFQed:
                        UnitTreeView.insert(parent="", index=END, iid=rec[0], 
                                        values=rec,checked = True)
                    
                    
            curRFQ.close()
            UnitTreeView.CalLineTotalPrice()
            UnitTreeView.TreeViewSAVED = 0
            SumAllSGD()
            
    def clearTreeUnit():
        
        UnitTreeView.delete(*UnitTreeView.get_children())
        UnitTreeView.CalLineTotalPrice()
        InsertReadonly(TotalSGDBox,'')
        UnitTreeView.TreeViewSAVED = 0
        
    def SaveCurrentParts():
        global CurrentRFQ
        if CurrentRFQ:
            curRFQAddParts = connRFQ.cursor()
            curRFQAddParts.execute(f"""TRUNCATE `rfq_master`.`{CurrentRFQ}`""")     
            connRFQ.commit()
            if UnitTreeView.get_children():
                MySqlCommand = f"""
                                INSERT INTO `rfq_master`.`{CurrentRFQ}`(
                                    `FullPartNum`,
                                    `UnitCost`,
                                    `PartNum`,
                                    `Assem_Full`,
                                    `PO`)
                                    VALUES """
                for iid in UnitTreeView.get_children():
                    FullPartNum = UnitTreeView.item(iid, 'values')[0]
                    PO = 1 if 'checked' in UnitTreeView.item(iid, 'tags') else 0
                    RFQUnitCost = 'NULL' if UnitTreeView.item(iid, 'values')[12] == 'None' else "{:.2f}".format(round(float(UnitTreeView.item(iid, 'values')[12]),2)) 
                    
                    PartInfo = FullPartNum.split('-',3)
                    Assem_Full = PartInfo[2]+'_'+PartInfo[3]
                
                    MySqlCommand += f"""
                                ('{FullPartNum}',
                                 {RFQUnitCost},
                                 '{PartInfo[3]}',
                                 '{Assem_Full}',
                                 {PO})""" + ',' 
                    
                    
                try: 
                    curRFQAddParts.execute(MySqlCommand[:-1])
                            
                    connRFQ.commit()
                    UnitTreeView.TreeViewSAVED = 1
                    #OpenCurrent(CurrentRFQ)
                    
        
                except Error as e:
                    messagebox.showerror('Save Current Parts in RFQ',
                                         f'Failed to save Part in RFQ\n({CurrentRFQ}){e}',
                                         parent = RFQFrame)
                    
                    connRFQ.rollback()
                    
            curRFQAddParts.close()
        
        else:
            messagebox.showerror('Save Current Parts in RFQ',
                                'No RFQ Created, Please Create a New RFQ Entry',
                                parent = RFQFrame)
        
    def SaveCurrentRFQ():
        global CurrentRFQ
        
        
        curRFQ = connRFQ.cursor()
        if CurrentRFQ:
            SaveCurrentParts()
            RFQ_REF_NO = CurrentRFQ
            PARTS_COMP = 0
            TOTAL_PARTS = len(UnitTreeView.get_children())
            VENDOR_NAME = VendorBox.get() if  VendorBox.get() else None
            STATUS = 0
            ISSUE_DATE = IssuedBox.get() if IssuedBox.get() else None
            REPLY_DATE = ReplyDueCalEnt.get() if  ReplyDueCalEnt.get() else None
            DELIVER_DATE = DeliverBCalEnt.get() if  DeliverBCalEnt.get() else None
            DATE_OF_ENTRY = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            CURRENCY = CurrencyCombo.get() if  CurrencyCombo.get() else None
            PURCHASER = PurchaserCombo.get()
            TOTAL_TCURR = UnitTreeView.TOTAL_TCURR
            TOTAL_SGD = float(TotalSGDBox.get()) if TotalSGDBox.get() != '' else None
            try:    
                curRFQ.execute("""UPDATE `rfq_master`.`rfq_list` SET
                            `PARTS_COMP`= %s,
                             `TOTAL_PARTS`= %s,
                             `VENDOR_NAME`= %s,
                             `STATUS`= %s,
                             `ISSUE_DATE`= %s,
                             `REPLY_DATE`= %s,
                             `DELIVER_DATE`= %s,
                             `DATE_OF_ENTRY`= %s,
                             `CURRENCY`= %s, 
                             `PURCHASER`= %s,
                             `TOTAL_TCURR`= %s,
                             `TOTAL_SGD` = %s
                             WHERE `RFQ_REF_NO` = %s
                                """, 
                            (PARTS_COMP,TOTAL_PARTS, VENDOR_NAME,
                             STATUS, ISSUE_DATE, REPLY_DATE, 
                             DELIVER_DATE, DATE_OF_ENTRY,CURRENCY,
                             PURCHASER, TOTAL_TCURR, TOTAL_SGD,RFQ_REF_NO))
                messagebox.showinfo('Save Current Parts in RFQ',
                                 f'Parts Saved in {CurrentRFQ}',
                                 parent = RFQFrame)
            except Error as e:
                print(e)
        else:
            messagebox.showerror('Save Current Parts in RFQ',
                                'No RFQ Created, Please Create a New RFQ Entry',
                                parent = RFQFrame)
        connRFQ.commit()
        curRFQ.close()
        
        
    def OpenCurrent(RFQNo):
        global ProjList,MachineList,MakerList,AsmDict,CurrentRFQ
        
        curRFQ = connRFQ.cursor()
        curRFQ.execute(f"""
                           SELECT * FROM `rfq_master`.`rfq_list` where `RFQ_REF_NO` = '{RFQNo}'
                           """)
        rfq_list = curRFQ.fetchall()
        
        if rfq_list:
            CurrentRFQInfo = rfq_list[0] 
            
            CurrentRFQ = RFQNo
            curRFQ.execute(f"""
                           SELECT `ASSEM_FULL`,`DES_QTY` FROM `{CurrentRFQInfo[2]}`.`{CurrentRFQInfo[3]}`
                           """)
                           
            AsmDict = {Asm[0]:Asm[1] for Asm in curRFQ.fetchall()}
            
            

            



            #print(CurrentRFQ)
            
            InsertReadonly(RFQRefBox,CurrentRFQInfo[1])
            
            ProjectCombo.configure(state='active')
            ProjectCombo.current(ProjList.index(CurrentRFQInfo[2]))
            ProjectCombo.configure(state='disabled') 
            _QueryMachNo()   
            MachineCombo.configure(state='active')
            MachineCombo.current(MachineList.index(CurrentRFQInfo[3]))
            MachineCombo.configure(state='disabled')
            _QueryMaker()
            InsertReadonly(MacQtyBox,CurrentRFQInfo[4])
            InsertReadonly(VendorBox,CurrentRFQInfo[7])
            InsertReadonly(StatusBox,StatusDict[int(CurrentRFQInfo[8])])
            InsertReadonly(IssuedBox,CurrentRFQInfo[9])
            InsertReadonly(ReplyDueCalEnt,CurrentRFQInfo[10])
            InsertReadonly(DeliverBCalEnt,CurrentRFQInfo[11])
            CurrencyCombo.current(CountryRef.getCcyLst().index(CurrentRFQInfo[13]))
            PurchaserCombo.current(list(EmployeeDict).index(CurrentRFQInfo[14]))
            InsertReadonly(TotalSGDBox,CurrentRFQInfo[16])
            
            CompletedCheckbutton.configure(state=ACTIVE)
            if CurrentRFQInfo[8] == 3:
                CompletedCheck.set(1)
            else:
                CompletedCheck.set(0)

            curRFQ.execute(f"""
                           SELECT * FROM `rfq_master`.`{RFQNo}`
                           """)
                           
            PartsList = curRFQ.fetchall()
            TreeViewParts = []
            for Parts in PartsList:
                FullPartNum = Parts[0]
                PartKeyInfo = FullPartNum.split('-')
                ProjID = PartKeyInfo[0]
                MachID = PartKeyInfo[1]
                AsmID = PartKeyInfo[2]
                PartNo = PartKeyInfo[3]
                RFQUnitCost = Parts[1]
                PO = Parts[4]
                try : 
                    assert ProjID == CurrentRFQInfo[2] and MachID == CurrentRFQInfo[3]
                except AssertionError as e:
                    messagebox.showerror('Load RFQ from Database',
                                         f'Unmatched Proj ID and Machine No. for Part in RFQ\n({CurrentRFQ[1]} : {FullPartNum})',
                                         parent = RFQFrame)
                    print(e)
                    
                Assem_full = MachID + '_' + AsmID
                
                curRFQ.execute(f"""SELECT oid, 
                                   PartNum, Description, CLS,
                                   Maker, Spec, V,  DES, SPA,
                                   OH, REQ, PCH,  
                                   UnitCost, Currency, TotalUnitCost
                                   FROM `{ProjID}`.`{Assem_full}`
                                   where PartNum = {PartNo}
                                   """)
                PartInfo = curRFQ.fetchall()[0] + (FullPartNum, AsmID,RFQUnitCost, PO)
            
                TreeViewParts.append(PartInfo)
                
            
            #print(TreeViewParts)
            clearTreeUnit()
            for rec in TreeViewParts:
                UnitTreeView.insert(parent="", index=END, iid=rec[15], 
                                values=(rec[15], rec[2], rec[3],
                                        rec[4], rec[5], rec[6], rec[7], rec[8],
                                        rec[9], AsmDict[rec[16]] , rec[10], rec[11],
                                        None if rec[17] == None else "{:.2f}".format(round(float(rec[17]), 2)),
                                        CurrentRFQInfo[13], 
                                        None if rec[14] == None else "{:.2f}".format(round(float(rec[14]), 2)),''))
                if rec[18]:
                    UnitTreeView.tag_remove(rec[15], 'unchecked')
                    UnitTreeView.tag_add(rec[15], ('checked',))
                
            UnitTreeView.TreeViewSAVED = 1
            UnitTreeView.RFQIssued = 1 if CurrentRFQInfo[9] else 0
            
            if UnitTreeView.get_children():
                UnitTreeView.CalLineTotalPrice()
                SumAllSGD()
            
        else:
            if messagebox.askyesno('Load RFQ from Database',f'RFQ No. {RFQNo} do not exist in Database\nDo you wish to create a RFQ',parent = RFQFrame):
                NewRFQ()
            else:
                RFQFrame.destroy()
                
        curRFQ.close()
            
    def CreateRFQEntry(NewVendor = False):
        global CurrentRFQ,MachineList   
        
        if MachineList:
            pass
        else:
            messagebox.showerror('Create New RFQ Entry','Please choose/create a valid machine',parent = RFQFrame)
        
        curRFQ = connRFQ.cursor()
        
        
        curRFQ.execute(f"""
                       SELECT RFQ_REF_NO FROM `rfq_master`.`rfq_list` 
                       where INSTR(RFQ_REF_NO, '{ProjectCombo.get() + MachineCombo.get()}') > 0
                       
                       """ )
        
        
                
        RFQList = [RFQ[0] for RFQ in curRFQ.fetchall()]
        ProjMachID = str(ProjectCombo.get()) + str(MachineCombo.get())
        connRFQ.commit()
        
        if RFQList:
            NextRFQNo = 1
            for ExistingRFQNo in RFQList:
                CurrRFQNo = int(ExistingRFQNo[ExistingRFQNo.find(ProjMachID)+len(ProjMachID): ])
                if  CurrRFQNo<= NextRFQNo:
                    NextRFQNo = CurrRFQNo+1
                    continue
                else:
                    pass
         
            CurrentRFQ = 'RFQ' + ProjMachID + str(NextRFQNo).rjust(4,'0')
        else:
            CurrentRFQ = 'RFQ' + ProjMachID + str(1).rjust(4,'0')
        
        RFQ_REF_NO = CurrentRFQ
        PROJECT_CLS_ID = ProjectCombo.get()
        MACH_ID = MachineCombo.get() 
        MACH_DES_QTY = MacQtyBox.get() if MacQtyBox.get() else None
        PARTS_COMP = 0
        TOTAL_PARTS = len(UnitTreeView.get_children())
        if NewVendor:
            VENDOR_NAME = None
        else:
            VENDOR_NAME = VendorBox.get() if  VendorBox.get() else None
        STATUS = 0
        ISSUE_DATE = None
        REPLY_DATE = ReplyDueCalEnt.get() if  ReplyDueCalEnt.get() else None
        DELIVER_DATE = DeliverBCalEnt.get() if  DeliverBCalEnt.get() else None
        DATE_OF_ENTRY = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        CURRENCY = CurrencyCombo.get() if  CurrencyCombo.get() else None
        PURCHASER = PurchaserCombo.get()
        TOTAL_TCURR = round(float(UnitTreeView.TOTAL_TCURR), 2)
        TOTAL_SGD = float(TotalSGDBox.get()) if TotalSGDBox.get() != '' else None
        try:    
            curRFQ.execute("""INSERT INTO `rfq_master`.`rfq_list`(
                        `RFQ_REF_NO`,`PROJECT_CLS_ID`,`MACH_ID`,`MACH_DES_QTY`,`PARTS_COMP`,
                         `TOTAL_PARTS`,`VENDOR_NAME`,`STATUS`,`ISSUE_DATE`,`REPLY_DATE`,
                         `DELIVER_DATE`,`DATE_OF_ENTRY`,`CURRENCY`, `PURCHASER`,`TOTAL_TCURR`,
                         `TOTAL_SGD`)
                            VALUES 
                        (%s ,%s, %s, %s, %s, %s, 
                         %s, %s, %s, %s, %s, %s, 
                         %s, %s, %s, %s)""", 
                        (RFQ_REF_NO, PROJECT_CLS_ID, MACH_ID, MACH_DES_QTY, PARTS_COMP,
                         TOTAL_PARTS, VENDOR_NAME, STATUS, ISSUE_DATE, REPLY_DATE, 
                         DELIVER_DATE, DATE_OF_ENTRY,CURRENCY, PURCHASER, TOTAL_TCURR, TOTAL_SGD))
                    
            connRFQ.commit()
            
            curRFQ.execute(f"""CREATE TABLE IF NOT EXISTS `rfq_master`.`{CurrentRFQ}`
                           (`FullPartNum` VARCHAR(20) PRIMARY KEY,
                            `UnitCost` FLOAT DEFAULT 0.00,
                            `PartNum` VARCHAR(10),
                            `Assem_Full` VARCHAR(10),
                            `PO` tinyint default 0
                            )
                        
                            ENGINE = InnoDB
                            DEFAULT CHARACTER SET = utf8mb4
                            COLLATE = utf8mb4_0900_ai_ci""")
            connRFQ.commit()
        except Error as e:
            messagebox.showerror('Create New RFQ Entry',
                                 f'Failed to Create RFQ\n({CurrentRFQ})\n{e}',
                                 parent = RFQFrame)
            connRFQ.rollback()
            return
            
        curRFQ.close()
        ProjectCombo.configure(state='disabled') 
        MachineCombo.configure(state='disabled')
        SaveCurrentParts()
        messagebox.showinfo('Create New RFQ Entry',f'{CurrentRFQ} created successful',parent = RFQFrame)
        OpenCurrent(CurrentRFQ)
    
    def ConfirmCreateRFQ():
        global CurrentRFQ
        RFQExist = ''
        if CurrentRFQ:
            RFQExist = f'Current RFQ already exist ({CurrentRFQ})\n'
        
        if messagebox.askyesno('Create New RFQ Entry',f'{RFQExist}Confirm Create?',parent = RFQFrame):
            CreateRFQEntry()

        
    
    def SaveVendorUnitCostToPM():
        global CurrentRFQ
        curRFQ = connRFQ.cursor()
        if CurrentRFQ:
        
            if messagebox.askyesno('Save Current Parts and Export to Project Manager',
                                   f'Do you want to export parts\nfrom ({CurrentRFQ}) to Project Manager?'
                                   ,parent = RFQFrame):
                try:
                    SaveCurrentRFQ()
                
                    
                    
                    assert any(['checked' in UnitTreeView.item(iid, 'tags') for iid in UnitTreeView.get_children()]) and UnitTreeView.TreeViewSAVED
    
                    
                    if UnitTreeView.get_children():
                        for iid in UnitTreeView.get_children(): 
                            PartInfo = UnitTreeView.item(iid, 'values')
                            FullPartNum = PartInfo[0]
                            Vendor = 'NULL'
                            KeyPartInfo = FullPartNum.split('-',3)
                            ProjID = KeyPartInfo[0]
                            MachID = KeyPartInfo[1]
                            AsmID = KeyPartInfo[2]
                            PartNo = KeyPartInfo[3]
                            Assem_Full = KeyPartInfo[1]+'_'+KeyPartInfo[2]
                            if 'checked' in UnitTreeView.item(iid, 'tags'):
                                Vendor = f"'{VendorBox.get()}'" if  VendorBox.get() else 'NULL'
                            
                            curRFQ.execute(f"""UPDATE `{ProjID}`.`{Assem_Full}` PM
                                               INNER JOIN `rfq_master`.`{CurrentRFQ}` RFQ
                                               ON `PM`.`PartNum` = `RFQ`.`PartNum`
                                               SET `PM`.`UnitCost` = `RFQ`.`UnitCost`,
                                                       `PM`.`TotalUnitCost` =  {PartInfo[14]},
                                                       `PM`.`Currency` = '{PartInfo[13]}',
                                                       `PM`.`ExchangeRate` =  {CountryRef.getExRate(PartInfo[13])},
                                                       `PM`.`TotalSGD` = {float(CountryRef.getExRate(PartInfo[13]) )* float(PartInfo[14])} ,
                                                        `PM`.`Vendor` = 
                                                         """ + Vendor)
                       
                    
                    connRFQ.commit()
                    
                    
                    OpenCurrent(CurrentRFQ)
                except Error as e:
                    print(e)
                    
                except AssertionError as e:
                    messagebox.showerror('Save Current Parts and Export to Project Manager','Please ensure RFQ is Issued and Parts are selected',parent = RFQFrame)
                 
                
            else:
                messagebox.showinfo('Save Changes to RFQ',f'No Changes made to RFQ ({CurrentRFQ})\n and Machine {ProjectCombo.get()}-{MachineCombo.get()} ',parent = RFQFrame)
                
        
        curRFQ.close()            
        #OpenCurrent(CurrentRFQ)
        
    def SelectPuchaseAll():
        UnitTreeView.CheckUncheckAll()
        

        
        
    def SumAllSGD():
        global CurrentRFQ
        UnitTreeView.CalLineTotalPrice()
        TotalSGD = 0
        if UnitTreeView.get_children():
            for iid in UnitTreeView.get_children():
                PartInfo = UnitTreeView.item(iid,'values')
                TotalSGD += float(CountryRef.getExRate(PartInfo[13])) * float(PartInfo[14])
            
            InsertReadonly(TotalSGDBox,"{:.2f}".format(round(float(TotalSGD), 2)))
            
        else:
            InsertReadonly(TotalSGDBox,'')
            
    def SaveRFQasPDF():
        
        PartsList = []
        for iid in UnitTreeView.get_children():
            PartInfo = UnitTreeView.item(iid,'values')
            PartsList+=[[PartInfo[0],
                        PartInfo[1],
                        PartInfo[3],
                        PartInfo[4],
                        PartInfo[10]
                        ]]
            
            
        FileDir  = os.path.join(os.path.join(os.getcwd(),'temp'), CurrentRFQ +'.csv') 
        with open(f"{FileDir}","w", newline="") as f:
            theWriter = csv.writer(f,lineterminator='\n')
            theWriter.writerow(['PartNum','Description','Maker','Maker Specification','Quatitiy'])
            for rec in PartsList:
                print(rec)
                theWriter.writerow(rec)
            f.close()
            
    def IssueRFQ():
        global CurrentRFQ
            
        if CurrentRFQ:
            try:
                assert ReplyDueCalEnt.get() and VendorBox.get()
                
            except:
                messagebox.showerror('Issue RFQ','Please ensure you choose a Vendor and Reply Date',parent = RFQFrame)
                return
            
            SaveCurrentRFQ()
            curRFQ = connRFQ.cursor()
            
            if IssuedBox.get() and UnitTreeView.RFQIssued:
                if messagebox.askyesno('Issue RFQ',f'This RFQ has been issued, Do you want to issue again?\n({CurrentRFQ})',parent = RFQFrame):
                    SaveRFQasPDF()
            
                    IssueEmail(CurrentRFQ = CurrentRFQ,Purchaser = PurchaserNameBox.get(),Vendor = VendorBox.get())
                    InsertReadonly(IssuedBox,datetime.now().strftime('%Y-%m-%d'))
                    
                    curRFQ.execute("""UPDATE `rfq_master`.`rfq_list` SET
                             `STATUS`= %s,
                             `ISSUE_DATE`= %s
                             WHERE `RFQ_REF_NO` = %s
                                """, 
                            (1, datetime.now().strftime('%Y-%m-%d'),CurrentRFQ))
                    connRFQ.commit()
                    
                    UnitTreeView.RFQIssued = 1
            else:
                SaveRFQasPDF()
            
                IssueEmail(CurrentRFQ = CurrentRFQ,Purchaser = PurchaserNameBox.get(),Vendor = VendorBox.get())
                InsertReadonly(IssuedBox,datetime.now().strftime('%Y-%m-%d'))
                curRFQ.execute("""UPDATE `rfq_master`.`rfq_list` SET
                             `STATUS`= %s,
                             `ISSUE_DATE`= %s
                             WHERE `RFQ_REF_NO` = %s
                                """, 
                            (1, datetime.now().strftime('%Y-%m-%d'),CurrentRFQ))
                connRFQ.commit()
                UnitTreeView.RFQIssued = 1
            curRFQ.close()    
        else:
            messagebox.showerror('Issue RFQ','Please Create/Select RFQ to Issue',parent = RFQFrame)
        
    def RFQNewVendor():
        if messagebox.askyesno('Create New RFQ','This RFQ has been issued, Do you want duplicate this RFQ for another Vendor?)',parent = RFQFrame):
                    
            CreateRFQEntry(NewVendor=True)
        else:
            pass
            
    def CloseTab():
        global CurrentRFQ
        if messagebox.askyesno('Close Tab',f'Do you want to save and before closing this RFQ tab?\n({CurrentRFQ if CurrentRFQ else None})',parent = RFQFrame):
            SaveCurrentRFQ()
            RFQFrame.destroy()
            CurrentRFQ = None
            
        else:
            RFQFrame.destroy()
            CurrentRFQ = None
             
            
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
    
    MachineCombo = ttk.Combobox(Frame1, width=30, height = 6,value = MachineList, state="readonly")
    MachineCombo.grid(row = 1, column = 1, padx = 10,pady =4,sticky = W)
    MachineCombo.bind("<<ComboboxSelected>>", QueryMaker)
    
    MachineNameLabel = Label(Frame1, text = '')
    MachineNameLabel.grid(row = 1, column = 2,pady =4,sticky  = W)
    
    MakerLabel = Label(Frame1, text = 'Maker')
    MakerLabel.grid(row = 2, column = 0, padx = 10,pady =4)
    
    SelectMakerPartsFrame = Frame(Frame1)
    SelectMakerPartsFrame.grid(row = 2, column = 1,columnspan = 2, padx = 10,pady =4,sticky = EW)

    MakerCombo = ttk.Combobox(SelectMakerPartsFrame, width=30, height = 6,value = MakerList, state="readonly")      
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
    
    CompletedCheck = IntVar()
    CompletedCheck.set(0)
    
    CompletedCheckbutton  = Checkbutton(Frame21,text = 'Completed',  variable= CompletedCheck)
    CompletedCheckbutton.grid(row = 0, column = 6,padx = (10,4),pady =4,sticky = W)
    CompletedCheckbutton.configure(state=DISABLED)

    Frame22 = Frame(Frame2)
    Frame22.grid(row = 0, column = 1,sticky = E)

    CurrencyLabel = Label(Frame22,text = 'Currency')
    CurrencyLabel.grid(row = 0, column = 0,padx = (10,4),pady =4,sticky = W)
    
    CurrencyCombo = ttk.Combobox(Frame22,width=6, value=CountryRef.getCcyLst(), state="readonly")
    CurrencyCombo.grid(row = 0, column = 1,padx = 4,pady =4,sticky = W)
    CurrencyCombo.bind("<<ComboboxSelected>>", ChangeCurrency)
    CurrencyCombo.current(0) if CurrencyCombo['value'] else None

    PurchaserLabel = Label(Frame22,text = 'Purchaser')
    PurchaserLabel.grid(row = 0, column = 2,padx = (10,4),pady =4,sticky = W)

    PurchaserCombo = ttk.Combobox(Frame22,width=8, value=list(EmployeeDict), state="readonly")
    PurchaserCombo.grid(row = 0, column = 3,padx = 4,pady =4,sticky = W)
    PurchaserCombo.bind("<<ComboboxSelected>>", ShowPurchaserName)
    PurchaserCombo.current(0) if PurchaserCombo['value'] else None
    
    PurchaserNameBox = Entry(Frame22,width=28,state = "readonly")
    PurchaserNameBox.grid(row = 0, column = 4,padx = 4,pady =4,sticky = W)
    
    Frame3 = LabelFrame(RFQFrame)
    Frame3.columnconfigure(1, weight=1)
    Frame3.grid(row = 2, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    TopFrame = Frame(Frame3)
    TopFrame.pack(side = TOP, fill=BOTH,expand = True)
    TopFrame.columnconfigure(0, weight=1)

    BottomTreeFrame = Frame(Frame3)
    BottomTreeFrame.pack(side= BOTTOM, fill=BOTH,expand = True)
    BottomTreeFrame.columnconfigure(0, weight=1)
    
    TotalSGDLabel = Label(BottomTreeFrame,text = 'Total SGD',font = ('Arial 10 bold'))
    TotalSGDLabel.grid(row = 0, column = 1,pady =4,padx = 10,sticky = E)
    
    TotalSGDBox = Entry(BottomTreeFrame,width=24,state = "readonly",font = ('Arial 10 bold'),justify=RIGHT)
    TotalSGDBox.grid(row = 0, column = 2, padx = 4,pady =4,sticky = W) 
    
    Button(TopFrame,text = 'Select All',command = SelectPuchaseAll).pack(side = RIGHT,padx = (4,8) ,pady =4)
    
    UnitTreeScroll = Scrollbar(Frame3)
    UnitTreeScroll.pack(side=RIGHT, fill=Y)
    
    UnitTreeView = CbTreeview(Frame3, yscrollcommand=UnitTreeScroll.set, 
                                selectmode="browse", TotalSGDBox = TotalSGDBox)
    UnitTreeScroll.config(command=UnitTreeView.yview)
    
    UnitTreeView.pack(padx=2, pady=2,fill="x", expand=True)
    
    
    
    
    UnitTreeView["columns"] = ("Part No", "Description", "CLS",  
                                "Maker", "Spec","V", "DES", "SPA", 
                                "OH","UA", "REQ", "PCH", 
                                "UnitCost", "Curr","Total","PO")
    
    UnitTreeView.column("#0", width=0, stretch=NO)
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
    

    BlankBox = Entry(BottomTreeFrame,state = "readonly" )
    BlankBox.grid(row = 0, column = 0,pady =4,sticky = EW)
    
    ClearPartsButton = Button(BottomTreeFrame,text = 'Clear Parts', width=14,command = clearTreeUnit)
    ClearPartsButton.grid(row = 0, column = 3, padx = (4,8) ,pady =4,sticky = W) 
    
    Frame4 = Frame(RFQFrame)
    Frame4.columnconfigure(3, weight=1)
    Frame4.grid(row = 3, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    CreateRFQButton = Button(Frame4,text = 'Create RFQ',width = 14,command = ConfirmCreateRFQ)
    CreateRFQButton.grid(row = 0, column = 0)
    
    IssueRFQButton = Button(Frame4,text = 'Issue RFQ' ,width = 14,command = IssueRFQ)
    IssueRFQButton.grid(row = 0, column = 1,padx = 8, pady = 6)
    
    DuplicateRFQButton = Button(Frame4,text = 'Duplicate RFQ' ,width = 14,command = RFQNewVendor)
    DuplicateRFQButton.grid(row = 0, column = 2,padx = 8, pady = 6)

    SaveRFQPartsButton = Button(Frame4,text = 'Save RFQ' ,width = 14, command = SaveCurrentRFQ)
    SaveRFQPartsButton.grid(row = 0, column = 4,padx = 8, pady = 6)

    SavePMButton = Button(Frame4,text = 'Save to PM' ,width = 14, command = SaveVendorUnitCostToPM)
    SavePMButton.grid(row = 0, column = 5,padx = 8, pady = 6)
        
    RefreshButton = Button(Frame4,text = 'Refresh' ,width = 14,command = Refresh)
    RefreshButton.grid(row = 0, column = 6,padx = (8,0), pady = 6)

    CloseButton = Button(Frame4,text = 'Close Tab' ,width = 14,command = CloseTab)
    CloseButton.grid(row = 0, column = 7,padx = (8,0), pady = 6)
    
    if CurrentRFQ:
        UnitTreeView.TreeViewSAVED = 1
    
    Refresh()
    


def reviewRFQ(tabNote):
    ReviewRFQFrame = Frame(tabNote)
    tabNote.add(ReviewRFQFrame, text="Review All RFQ")
    ReviewRFQFrame.columnconfigure(0, weight=1)
    tabNote.select(ReviewRFQFrame)
   
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background="silver",
                    rowheight=20,
                    fieldbackground="light grey")
    
    style.map("Treeview")
    
    def RefreshTreeView(SearchCrit= None,Search = None):
        clearTreeUnit()
        curRFQ = connRFQ.cursor()
        
        if SearchCrit and Search: 
            curRFQ.execute(f"""SELECT `RFQ`.`oid`,`RFQ_REF_NO`,`PROJECT_CLS_ID`,`MACH_ID`,`MACH_DES_QTY`,`PARTS_COMP`,
                             `TOTAL_PARTS`,`VENDOR_NAME`,`RFQ`.`STATUS`,`ISSUE_DATE`,`REPLY_DATE`,
                             `DATE_OF_ENTRY`,`TOTAL_TCURR`,`CURRENCY`,`TOTAL_SGD`,`EM`.`EMPLOYEE_NAME`
                              FROM `rfq_master`.`rfq_list` `RFQ`
                              LEFT JOIN `index_emp_master`.`emp_data` `EM`
                              ON `EM`.`EMPLOYEE_ID` = `RFQ`.`PURCHASER`
                              WHERE INSTR(`{SearchCrit}`,'{Search}') > 0
                              ORDER BY `PROJECT_CLS_ID` ASC,`MACH_ID` ASC,`DATE_OF_ENTRY` ASC
                              """)
                              
        else:
            
            curRFQ.execute("""
                           SELECT `RFQ`.`oid`,`RFQ_REF_NO`,`PROJECT_CLS_ID`,`MACH_ID`,`MACH_DES_QTY`,`PARTS_COMP`,
                             `TOTAL_PARTS`,`VENDOR_NAME`,`RFQ`.`STATUS`,`ISSUE_DATE`,`REPLY_DATE`,
                             `DATE_OF_ENTRY`,`TOTAL_TCURR`,`CURRENCY`,`TOTAL_SGD`,`EM`.`EMPLOYEE_NAME`
                          FROM `rfq_master`.`rfq_list` `RFQ`
                          LEFT JOIN `index_emp_master`.`emp_data` `EM`
                          ON `EM`.`EMPLOYEE_ID` = `RFQ`.`PURCHASER`
                          ORDER BY `PROJECT_CLS_ID` ASC,`MACH_ID` ASC,`DATE_OF_ENTRY` ASC
                           """)

            
                        
        RFQList = curRFQ.fetchall()
        for rec in RFQList:
            RFQTreeView.insert(parent="", index=END, iid=rec[0], 
                            values=(rec[1], rec[2], rec[3],
                                    rec[4], rec[5], rec[6], rec[7], rec[8],
                                    rec[9], rec[10], rec[11], rec[12],rec[13],
                                    rec[14],rec[15]))
                
 
        curRFQ.close()
        
    def SearchCombCompList(e):
        _SearchCombCompList()
    def _SearchCombCompList():
        CritIndex = RFQTreeView["columns"].index(SearchCombo.get())
        CompList = []
        for iid in RFQTreeView.get_children():
            CritName = RFQTreeView.item(iid,'values')[CritIndex]
            if CritName not in CompList:
                
                CompList+=[CritName]
            
        SearchEntry.set_completion_list(CompList)
        SearchEntry.delete(0,END)
        
    def clearTreeUnit():
        
        RFQTreeView.delete(*RFQTreeView.get_children())
        
            
    def openRFQTab():
        selectVal= RFQTreeView.selection()
        if selectVal == ():
            messagebox.showwarning("Unable to Load", "Please Select a RFQ",
                                   parent=ReviewRFQFrame)
        else:
            openRFQ(tabNote,RFQ = RFQTreeView.item(selectVal[0],'values')[0])
    
    def SearchwCrit():
        RefreshTreeView(SearchCrit=SearchDict[SearchCombo.get()], Search=SearchEntry.get())
        
    def DeleteRFQ():
        selectVal= RFQTreeView.selection()
        if selectVal == ():
            messagebox.showwarning("Unable to Delete", "Please Select a RFQ",
                                   parent=ReviewRFQFrame)
        else:
            oid = selectVal[0]
            RFQNo = RFQTreeView.item(oid,'values')[0]
            if messagebox.askyesno('Delete RFQ',f'Confirm Delete {RFQNo}?',parent = ReviewRFQFrame):
                curRFQ = connRFQ.cursor()
                try:
                    curRFQ.execute(f"""
                                   DELETE FROM `rfq_master`.`rfq_list`
                                   WHERE `oid` = {oid}
                                   """)
                    connRFQ.commit()
                    
                    curRFQ.execute(f"""
                                   DROP TABLE  IF EXISTS `rfq_master`.`{RFQNo}`
                                   """)
                    connRFQ.commit()
                except Error as e:
                    connRFQ.rollback()
                    print(e)
        
                curRFQ.close()
                RefreshTreeView()
                
    def CreateNewRFQ():
        openRFQ(tabNote)
        
        
    def CloseTab():
        ReviewRFQFrame.destroy()
    
    Frame1 = LabelFrame(ReviewRFQFrame)
    Frame1.columnconfigure(1, weight=1)
    Frame1.grid(row = 0, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    Label(Frame1, text = 'RFQ List', font="Arial 12 bold").pack(side = LEFT,padx = 6)
    
    Frame2 = LabelFrame(ReviewRFQFrame,relief = 'ridge')
    Frame2.columnconfigure(1, weight=1)
    Frame2.grid(row = 1, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    TopTreeFrame = Frame(Frame2)
    TopTreeFrame.pack(side = TOP, fill=BOTH,expand = True)
    TopTreeFrame.columnconfigure(0, weight=1)

    BottomTreeFrame = Frame(Frame2)
    BottomTreeFrame.pack(side= BOTTOM, fill=BOTH,expand = True)
    BottomTreeFrame.columnconfigure(0, weight=1)

    
    RFQTreeScroll = Scrollbar(Frame2)
    RFQTreeScroll.pack(side=RIGHT, fill=Y)
    
    RFQTreeView = ttk.Treeview(Frame2, yscrollcommand=RFQTreeScroll.set, 
                                selectmode="browse",height = 14)
    RFQTreeScroll.config(command=RFQTreeView.yview)
    
    RFQTreeView.pack(padx=2, pady=2,fill="x", expand=True)
    
    RFQTreeView['show'] = 'headings'
    
    RFQTreeView["columns"] = ("RFQ Number", "Project ID", "Machine",  
                                "Mach QTY", "Parts Comp","Total Parts", "Vendor", "Status", 
                                "Issue Date","Reply Date", "Entry Date", "Total Cost" , "Currency",
                                "Total SGD","Purchaser")
    
        
    RFQTreeView.column("#0", width=0, stretch=NO)
    RFQTreeView.column("RFQ Number", anchor=W, width=100)
    RFQTreeView.column("Project ID", anchor=W, width=80)
    RFQTreeView.column("Machine", anchor=W, width=80)
    RFQTreeView.column("Mach QTY", anchor=CENTER, width=30)
    RFQTreeView.column("Parts Comp", anchor=CENTER, width=30)
    RFQTreeView.column("Total Parts", anchor=CENTER, width=30)
    RFQTreeView.column("Vendor", anchor=W, width=80)
    RFQTreeView.column("Status", anchor=CENTER, width=40)
    RFQTreeView.column("Issue Date", anchor=CENTER, width=80)
    RFQTreeView.column("Reply Date", anchor=CENTER, width=80)
    RFQTreeView.column("Entry Date", anchor=CENTER, width=80)
    RFQTreeView.column("Total Cost", anchor=E, width=80)
    RFQTreeView.column("Currency", anchor=CENTER, width=30)
    RFQTreeView.column("Total SGD", anchor=E, width=80)
    RFQTreeView.column("Purchaser", anchor=W, width=80)
    
    
    RFQTreeView.heading("#0", text="Index", anchor=W)
    RFQTreeView.heading("RFQ Number", text="RFQ Number", anchor=W)
    RFQTreeView.heading("Project ID", text="Project ID", anchor=W)
    RFQTreeView.heading("Machine", text="Machine", anchor=W)
    RFQTreeView.heading("Mach QTY", text="Mach QTY", anchor=CENTER)
    RFQTreeView.heading("Parts Comp", text="Parts Comp", anchor=CENTER) 
    RFQTreeView.heading("Total Parts", text="Total Parts", anchor=CENTER)
    RFQTreeView.heading("Vendor", text="Vendor", anchor=W)
    RFQTreeView.heading("Status", text="Status", anchor=CENTER)
    RFQTreeView.heading("Issue Date", text="Issue Date", anchor=CENTER)
    RFQTreeView.heading("Reply Date", text="Reply Date", anchor=CENTER)
    RFQTreeView.heading("Entry Date", text="Entry Date", anchor=CENTER)
    RFQTreeView.heading("Total Cost", text="Total Cost", anchor=E)
    RFQTreeView.heading("Currency", text="Curr", anchor=CENTER)
    RFQTreeView.heading("Total SGD", text="Total SGD", anchor=E)
    RFQTreeView.heading("Purchaser", text="Purchaser", anchor=W)
    
    RFQTreeView.bind('<Double-Button-1>',lambda e: openRFQTab())
    
    RefreshTreeView()
    
    Frame3 = Frame(ReviewRFQFrame)
    Frame3.columnconfigure(1, weight=1)
    Frame3.grid(row = 2, column = 0, ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky =EW)
    
    SearchDict = {'RFQ Number': 'RFQ_REF_NO', 
                  'Project ID': 'PROJECT_CLS_ID',
                  'Vendor': 'VENDOR_NAME',
                  'Currency': 'CURRENCY'}
    
    SearchLabel = Label(Frame3, text = 'Search :')
    SearchLabel.pack(side = LEFT,pady = 6)
    
    SearchCombo = ttk.Combobox(Frame3, value = list(SearchDict) ,width = 16,state = 'readonly')
    SearchCombo.pack(side = LEFT,pady = 6)
    SearchCombo.bind("<<ComboboxSelected>>", SearchCombCompList)
    SearchCombo.current(0)


    SearchEntry = AutocompleteEntry(Frame3)
    SearchEntry.set_completion_list(['abc','cde'])
    SearchEntry.pack(side = LEFT,padx = 10,pady = 6)
    SearchEntry.bind("<Return>", SearchwCrit)
        
    _SearchCombCompList()
    
    SearchButton = Button(Frame3, text= 'Search', command = SearchwCrit,width = 14)
    SearchButton.pack(side = LEFT,padx = 10,pady = 6)
    
    ShowAllButton = Button(Frame3, text= 'Show All', command = RefreshTreeView,width = 14)
    ShowAllButton.pack(side = LEFT,padx = 10,pady = 6)
    
    CloseTabButton  = Button(Frame3, text= 'Close Tab', command = CloseTab,width = 14)
    CloseTabButton.pack(side = RIGHT,pady = 6,padx = (0,10))
    
    LoadRFQButton  = Button(Frame3, text= 'Load RFQ', command = openRFQTab,width = 14)
    LoadRFQButton.pack(side = RIGHT,pady = 6,padx = 10)

    DeleteRFQButton  = Button(Frame3, text= 'Delete RFQ', command = DeleteRFQ,width = 14)
    DeleteRFQButton.pack(side = RIGHT,pady = 6,padx = 10)

    CreateRFQButton = Button(Frame3, text= 'Create RFQ', command = CreateNewRFQ,width = 14)
    CreateRFQButton.pack(side = RIGHT,pady = 6,padx = 10)

def _SelectVendor(VendorBox,VendorAddressBox,VendorNameLabel):
    global img
    connVend = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password =logininfo[2],
                                       database = "INDEX_VEND_MASTER")
    
    curVend = connVend.cursor()
    curVend.execute(f"""SELECT * FROM index_vend_master.vendor_list
                    """)
    VendList = curVend.fetchall()
    VendDict = {Vend[1]:Vend for Vend in VendList}
    curVend.close()
    
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

    global currVEND
    currVEND = None
    def selectItem(event):
        global currVEND
        currVEND =VendorSelectionTree.item(VendorSelectionTree.focus())['values']

        
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

    def SelectedVend(VendorBox,VendorAddressBox,VendorNameLabel):
        global currVEND
        Vend = VendDict[currVEND[0]]
        try:
            VendAddress = Vend[8]+','+Vend[9]+','+Vend[6]+','+Vend[5]+','+Vend[7]+','+Vend[4]
            InsertReadonly(VendorBox,Vend[1])
            InsertReadonly(VendorAddressBox,VendAddress)
            
            VendorNameLabel['text'] = Vend[3]
            
            
            
            VendWin.destroy()
        except:
            if currVEND:
                messagebox.showwarning("No Vendor Selected", "Please Select a Vendor", 
                                   parent=VendWin)
            else:
                messagebox.showwarning("Select Vendor", "Error in Selecting Vendor", 
                                   parent=VendWin)
    
    def ClearVend(VendorBox,VendorAddressBox,VendorNameLabel):
        InsertReadonly(VendorBox)
        InsertReadonly(VendorAddressBox)
        VendWin.destroy()
        
        VendorNameLabel['text'] = ''
    
    def BackVend():
        VendWin.destroy()

    SelButton = Button(VendWin, text ="Select", command = lambda : SelectedVend(VendorBox,VendorAddressBox,VendorNameLabel), width = 10)
    SelButton.grid(row = 2, column = 0)
    
    ClearButton = Button(VendWin, text ="Clear", command = lambda : ClearVend(VendorBox,VendorAddressBox,VendorNameLabel), width = 10)
    ClearButton.grid(row = 2, column = 1)
    
    ExitButton = Button(VendWin, text ="Exit", command = BackVend, width = 10)
    ExitButton.grid(row = 2, column = 2)

def InsertReadonly(Entry,Input = None):
    Entry.config(state="normal")
    Entry.delete(0, END)
    
    if Input:
        Entry.insert(0, Input)
    Entry.config(state="readonly")

def CalWin(Entry):
    calWin = Toplevel()
    calWin.title("Select the Date")
    calWin.columnconfigure( 1,weight = 1)
    
    cal = Calendar(calWin, selectmode="day", date_pattern="y-mm-dd")
    cal.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky=EW)
    
    def confirmDate():
        val = cal.get_date()
        InsertReadonly(Entry,val)
        calWin.destroy()
    
    def emptyDate():
        InsertReadonly(Entry)
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
    reviewRFQ(tabNote)
    openRFQ(tabNote)
    root.mainloop()
    
    
    
    
    
    
