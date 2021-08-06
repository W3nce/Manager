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
from fpdf import FPDF
import ConnConfig,CountryRef

import os
import win32com.client as client

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

connPO = mysql.connector.connect(host = logininfo[0],
                                       user = logininfo[1], 
                                       password = logininfo[2])

Cross = os.path.join(os.path.join(os.getcwd(),'icons'), "cross.png") 




class AttachmentFile(LabelFrame):
    def  __init__(self,master,filename = None, FileDict = None,PositionArry = None,Close = True,**kw):
        self.filename = filename
        self.FileDict = FileDict
        self.PositionArry = PositionArry
        self.root = master
        self.cross = PhotoImage('cross',file = Cross)
       
       
        LabelFrame.__init__(self, master,relief  = RAISED, **kw)
        self.Label = Label(self,text =filename if filename else '')
        self.Label.pack(side = LEFT)
        if Close:
            self.Button = Button(self,image =self.cross ,command = lambda : self.Remove())
            self.Button.pack(side = RIGHT)
           
       
    def Remove(self):
        def FileIn(name,myarray):
            for n, i in enumerate(myarray):
                for sublist in myarray:
                    if i == name:
                        myarray[n]=None
                        
        if self.FileDict:
            self.FileDict.pop(self.filename)
        if self.PositionArry:
            FileIn(self.filename,self.PositionArry)
            
        self.destroy()
       
       
    

def EmailPOWindow(PONumber = None,Purchaser = None,Vendor = None):
    root = Toplevel()
    root.iconbitmap("MWA_Icon.ico")
    root.title("Send PO Email")
    root.columnconfigure(0, weight=1)
    curPO = connPO.cursor()
    VendorInfo = []
        
    if PONumber:
        _PONumber = PONumber
    else :
        _PONumber = ''
    
    if Purchaser:
        _Purchaser = Purchaser
    else :
        _Purchaser = ''
        
        
    
    if Vendor:
        _Vendor = Vendor
        curPO.execute(f"""
                       SELECT `MAIN_CONTACT_PERSON`, 
                       `MAIN_CONTACT_EMAIL`, 
                       `SECONDARY_CONTACT_PERSON`, 
                       `SECONDARY_CONTACT_EMAIL` FROM `index_vend_master`.`vendor_list`
                       WHERE `VENDOR_NAME` = '{_Vendor}'
                       """)
        VendorInfo = curPO.fetchall()[0]
    if VendorInfo:
        VendorEmailDict = {VendorInfo[0]: VendorInfo[1], VendorInfo[2]: VendorInfo[3]}
    else:
        messagebox.showerror('Email PO','Failed to get Vendor' ,parent = root)
        return
    
    curPO.close()
    
    AttachmentDict = {}
    
    AttachmentArry = [[None,None,None,None,None]]
    
    
    def ChangeTo(event):
        _ChangeTo()
    def _ChangeTo():
        ToEntry.delete(0,END)
        ToEntry.insert(END,VendorEmailDict[ToEntryCombo.get()])
        
    def ShowEmail():
        
        PartTypeList = ['STANDARD PART', 'FABRICATION PART','GENERAL SUPPLY']
        PartType = PartTypeList[0]
        
        CCEntry.delete(0,END)
        CCEntry.insert(END,'Jeff.zhao@motionwell.com.sg')
        
        SubjectEntry.delete(0,END)
        SubjectEntry.insert(END,f"""MOTIONWELL - PO : {_PONumber} - {PartType} """)
        
        POPDFfileName  = _PONumber+'.pdf'
        POPDFfiledir = os.path.join(os.getcwd(), POPDFfileName) 
        
        AttachmentArry[0][0] = POPDFfileName
        
        AttachmentDict.update({POPDFfileName:(POPDFfiledir,AttachmentFile(AttachmentFileFrame,
                                                                    filename = POPDFfileName,
                                                                    FileDict= AttachmentDict,
                                                                    PositionArry = AttachmentArry,
                                                                    Close= False))})
        
        AttachmentDict[POPDFfileName][1].grid(row = 0, column = 0,padx = 6, pady = 4)
        
        BodyText.delete('1.0',END)
        BodyText.insert(END, f"""Dear {_Vendor},
 
I have attached our purchase order, {_PONumber} for your reference.
 
Please help us to expedite the order and update me on this order.
 
Have a nice day!



Kind regards,   
      
{_Purchaser}   
           
Motionwell Automation Pte. Ltd.
Website: www.motionwell.com.sg
Tel: 65502218
20 Woodlands Link
#09-08(Design & Assembly Center)
#07-26(Manufacturing Shop)
Woodlands Industrial Estate
Singapore 738733

Website: www.motionwell.com.sg| HP: +65 87684297

The contents of this e-mail message and any attachments are confidential and are intended solely for addressee. The information may also be legally privileged. This transmission is sent in trust, for the sole purpose of delivery to the intended recipient. If you have received this transmission in error, any use, reproduction or dissemination of this transmission is strictly prohibited. If you are not the intended recipient, please immediately notify the sender by reply e-mail or phone and delete this message and its attachments, if any.""")
 

    def AddAttach():
        fileDir = filedialog.askopenfilename(initialdir="~/Desktop",
                                              title="Select A File",
                                              filetypes=(("CSV Files", "*.csv"),
                                                        ("Any Files", "*.*")),parent = root)
        
        filename, file_extension = os.path.splitext(fileDir)
        BaseName = os.path.basename(fileDir)
        try:
            assert os.path.exists(fileDir) and BaseName not in AttachmentDict
            AttachmentDict.update({BaseName:(fileDir,AttachmentFile(AttachmentFileFrame,
                                                                    filename = BaseName,
                                                                    FileDict= AttachmentDict,
                                                                    PositionArry = AttachmentArry))})
            
            if len(AttachmentDict)> len(AttachmentArry) *5:
                AttachmentArry.append([None,None,None,None,None])
                
            for Row in range(len(AttachmentArry)):
                for Attachment in range(len(AttachmentArry[Row])):
                    if AttachmentArry[Row][Attachment]:
                        pass
                    else:
                        AttachmentArry[Row][Attachment] = BaseName
                        
                        AttachmentDict[BaseName][1].grid(row = Row, column = Attachment,padx = 6, pady = 4)
                        break
   
                
                
        except AssertionError as e:
            pass
        
    def SendPOEmail():
        print(AttachmentDict)
        _SendPOEmail(AttachmentDict)
        
        
    def _SendPOEmail(AttachmentDict):
        AttachmentLst =list(AttachmentDict)
        AttachmentDict2 = {}
        for key in AttachmentLst:
            AttachmentDict2.update({key:AttachmentDict[key][0]})
        
        EmailPO(ToEntry.get(),
             CC = CCEntry.get(),
             BCC = BCCEntry.get(),
             root = root,
             CurrentPO = _PONumber,
             Vendor = _Vendor,
             Name = _Purchaser,
             Attachment = AttachmentDict2,
             Subject = SubjectEntry.get(),
             Body = BodyText.get('0.0',END))
        
        
        
    

    AddressFrame = LabelFrame(root,)
    AddressFrame.grid(ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky = EW)
    AddressFrame.columnconfigure(3, weight=1)
    
    ToLabel = Label(AddressFrame, text = 'To')
    ToLabel.grid(row = 0, column = 0, pady = 10, padx = 4, sticky = E)
    Label(AddressFrame,text = ':').grid(row = 0, column = 1 )  
    
    ToEntryCombo = ttk.Combobox(AddressFrame,value =list(VendorEmailDict) ,width = 20)
    ToEntryCombo.grid(row = 0, column = 2, pady = 10, padx = 4,sticky = EW)
    ToEntryCombo.bind("<<ComboboxSelected>>", ChangeTo)
    ToEntryCombo.current(0) if ToEntryCombo['value'] else None
    
    ToEntry = Entry(AddressFrame)
    ToEntry.grid(row = 0, column = 3, pady = 10, padx = 4,sticky = EW)
    
    _ChangeTo()
    CCLabel = Label(AddressFrame, text = 'CC')
    CCLabel.grid(row = 1, column = 0, pady = 10, padx = 4, sticky = E)
    Label(AddressFrame,text = ':').grid(row = 1, column = 1 )
    
    CCEntry = Entry(AddressFrame,width = 50)
    CCEntry.grid(row = 1, column = 2,columnspan =2, pady = 10, padx = 4, sticky = EW)
    
    BCCLabel = Label(AddressFrame,text =  'BCC')
    BCCLabel.grid(row = 2, column = 0, pady = 10, padx = 4, sticky = E)
    Label(AddressFrame,text = ':').grid(row = 2, column = 1 )
    
    BCCEntry = Entry(AddressFrame,width = 50)
    BCCEntry.grid(row = 2, column = 2,columnspan =2, pady = 10, padx = 4, sticky = EW)
    
    SendButton = Button(AddressFrame,text = 'Send', command = SendPOEmail,height = 3,width = 14)
    
    SendButton.grid(row = 0, column = 4,rowspan = 2,pady = 10, padx = 4, sticky = E)
    
    CancelButton =  Button(AddressFrame,text = 'Cancel', command = lambda : root.destroy() ,width = 14,height = 1 )
    CancelButton.grid(row = 2, column = 4,pady = 10, padx = 4, sticky = E)
    
    SubjectFrame = LabelFrame(root,)
    SubjectFrame.grid(ipadx = 10, ipady = 6,pady = (6,0), padx = 6,sticky = EW)
    SubjectFrame.columnconfigure(1, weight=1)
    
    SubjectLabel = Label(SubjectFrame, text = 'Subject  :')
    SubjectLabel.grid(row = 0, column = 0, pady = 10, padx = 4, sticky = W)
    
    SubjectEntry =Entry(SubjectFrame,width = 100)
    SubjectEntry.grid(row = 0, column = 1, pady = 10, padx = 4, sticky = EW)
    
    AttachmentFrame = Frame(root,)
    AttachmentFrame.columnconfigure(1, weight=1)
    AttachmentFrame.grid(ipadx = 10, ipady = 2,pady = (6,0), padx = 6,sticky = EW)
    
    AttachmentLabel = Label(AttachmentFrame,text = 'Attachment :')
    AttachmentLabel.grid(row = 0,column = 0)
    
    AttachmentFileFrame = Frame(AttachmentFrame)
    AttachmentFileFrame.grid(row = 0,column = 1,rowspan = 2,sticky =EW)
    AttachmentFileFrame.columnconfigure(4, weight=1)
    
    AddAttachment = Button(AttachmentFrame,text = 'Add',width = 8,command = AddAttach)
    AddAttachment.grid(pady = 4,row = 0,column = 2,rowspan = 2, sticky = NSEW)
    
    BodyFrame = Frame(root)
    BodyFrame.grid(ipadx = 10, padx = 6,sticky = EW)
    BodyFrame.columnconfigure(1, weight=1)
    
    BodyText = Text(BodyFrame)
    BodyText.grid(row = 0, column = 1, pady = 10, padx = 4, sticky = NSEW)
                 
        
    BodyTextScroll = ttk.Scrollbar(BodyFrame,command=BodyText.yview)
    BodyTextScroll.grid(row=0, column=1, pady = 10,padx = (1,4),sticky='nsew')
    BodyText['yscrollcommand'] = BodyTextScroll.set
    
    ShowEmail()   


def EmailPO(TO,
             CC = None,
             BCC = None,
             root = None,
             Vendor = None,
             Name = None,
             CurrentPO = None,
             Attachment = None,
             Subject = None,
             Body = None):
        
        Outlook = client.Dispatch('Outlook.Application')
        NewPOMail = Outlook.CreateItem(0)
        NewPOMail.To = TO
        if CC:
            NewPOMail.CC = CC
        if BCC:
            NewPOMail.BCC =  BCC
        if Vendor:
            _Vendor = Vendor
        else:
            _Vendor = ''
        if Name:
            _Name = Name
        else:
            _Name = ''
        if CurrentPO:
            _CurrentPO = CurrentPO
        else:
            _CurrentPO = ''
            
        if Attachment:
            _AttachmentDict = Attachment
        else:
            _AttachmentDict = {}
            
        PartTypeList = ['STANDARD PART', 'FABRICATION PART','GENERAL SUPPLY']
        PartType = PartTypeList[0]
        if Subject:
            NewPOMail.Subject = Subject
        else: NewPOMail.Subject = f"""MOTIONWELL - PO : {_CurrentPO} - {PartType} """
        if Body:
            NewPOMail.Body = Body
        else: NewPOMail.Body = f"""Dear {_Vendor}

I have attached our purchase order, {_CurrentPO} for your reference.
 
Please help us to expedite the order and update me on this order.
 
Have a nice day!



Kind regards,  
      
{_Name}   
           
Motionwell Automation Pte. Ltd.
Website: www.motionwell.com.sg
Tel: 65502218
20 Woodlands Link
#09-08(Design & Assembly Center)
#07-26(Manufacturing Shop)
Woodlands Industrial Estate
Singapore 738733

Website: www.motionwell.com.sg| HP: +65 87684297

The contents of this e-mail message and any attachments are confidential and are intended solely for addressee. The information may also be legally privileged. This transmission is sent in trust, for the sole purpose of delivery to the intended recipient. If you have received this transmission in error, any use, reproduction or dissemination of this transmission is strictly prohibited. If you are not the intended recipient, please immediately notify the sender by reply e-mail or phone and delete this message and its attachments, if any."""
        try:
            
            for Attachment in _AttachmentDict:
                assert os.path.exists(_AttachmentDict[Attachment])
                NewPOMail.Attachments.Add(Source=_AttachmentDict[Attachment])
            #message.Display(True)
            NewPOMail.send
            messagebox.showinfo('Issue PO',f'{_CurrentPO} has been issued')
        except AssertionError as e:
            messagebox.showerror('PO Email', 'Error : {e} \n{_AttachmentDict[Attachment]} does not exist')
        root.destroy()

def CloseTab():
    root.destroy()
    pass

if __name__ == "__main__":
    root = Tk()
    root.title("PO") 
    root.state("zoomed")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    global tabNote
    def a():
        EmailPOWindow(PONumber = 'MWAPOP21001010001', Purchaser = 'Zi Wei' ,Vendor = 'Airtac')
    tabNote = ttk.Notebook(root)
    tabNote.grid(row=0, column=0, sticky="NSEW") 
    POEFrame = Frame(tabNote)
    tabNote.add(POEFrame, text="issue PO Email")
    POEFrame.columnconfigure(0, weight=1)
    tabNote.select(POEFrame)
    Button(POEFrame,text = 'press',command = a).pack()
    root.mainloop()
    