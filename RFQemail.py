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

connRFQ = mysql.connector.connect(host = logininfo[0],
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
       
       
    

def EmailRFQWindow(CurrentRFQ = None,Purchaser = None,Vendor = None):
    root = Toplevel()
    root.iconbitmap("MWA_Icon.ico")
    root.title("Send RFQ Email")
    root.columnconfigure(0, weight=1)
    curRFQ = connRFQ.cursor()
    VendorInfo = []
        
    if CurrentRFQ:
        _CurrentRFQ = CurrentRFQ
    else :
        _CurrentRFQ = ''
    
    if Purchaser:
        _Purchaser = Purchaser
    else :
        _Purchaser = ''
        
        
    
    if Vendor:
        _Vendor = Vendor
        curRFQ.execute(f"""
                       SELECT `MAIN_CONTACT_PERSON`, 
                       `MAIN_CONTACT_EMAIL`, 
                       `SECONDARY_CONTACT_PERSON`, 
                       `SECONDARY_CONTACT_EMAIL` FROM `index_vend_master`.`vendor_list`
                       WHERE `VENDOR_NAME` = '{_Vendor}'
                       """)
        VendorInfo = curRFQ.fetchall()[0]
    if VendorInfo:
        VendorEmailDict = {VendorInfo[0]: VendorInfo[1], VendorInfo[2]: VendorInfo[3]}
    else:
        messagebox.showerror('Email RFQ','Failed to get Vendor' ,parent = root)
        return
    
    curRFQ.close()
    
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
        SubjectEntry.insert(END,f"""MOTIONWELL - RFQ : {_CurrentRFQ} - {PartType} """)
        
        RFQCSVfileName  = CurrentRFQ+'.csv'
        RFQCSVfiledir = os.path.join(os.path.join(os.getcwd(),'temp'), RFQCSVfileName) 
        
        AttachmentArry[0][0] = RFQCSVfileName
        
        AttachmentDict.update({RFQCSVfileName:(RFQCSVfiledir,AttachmentFile(AttachmentFileFrame,
                                                                    filename = RFQCSVfileName,
                                                                    FileDict= AttachmentDict,
                                                                    PositionArry = AttachmentArry,
                                                                    Close= False))})
        
        AttachmentDict[RFQCSVfileName][1].grid(row = 0, column = 0,padx = 6, pady = 4)
        
        BodyText.delete('1.0',END)
        BodyText.insert(END, f"""Dear {_Vendor},

Please help to quote with your best price for the item attached in the CSV File provided

Thank you!



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
        
    def SendRFQEmail():
        print(AttachmentDict)
        _SendRFQEmail(AttachmentDict)
        
        
    def _SendRFQEmail(AttachmentDict):
        AttachmentLst =list(AttachmentDict)
        AttachmentDict2 = {}
        for key in AttachmentLst:
            AttachmentDict2.update({key:AttachmentDict[key][0]})
        
        EmailRFQ(ToEntry.get(),
             CC = CCEntry.get(),
             BCC = BCCEntry.get(),
             root = root,
             CurrentRFQ = _CurrentRFQ,
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
    
    SendButton = Button(AddressFrame,text = 'Send', command = SendRFQEmail,height = 3,width = 14)
    
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
    BodyFrame.columnconfigure(0, weight=1)
    
    BodyText = Text(BodyFrame)
    BodyText.grid(row = 0, column = 0, pady = 10, padx = (4,0), sticky = NSEW)
    
    BodyTextScroll = ttk.Scrollbar(BodyFrame,command=BodyText.yview)
    BodyTextScroll.grid(row=0, column=1, pady = 10,padx = (1,4),sticky='nsew')
    BodyText['yscrollcommand'] = BodyTextScroll.set
                 
    ShowEmail()   


def EmailRFQ(TO,
             CC = None,
             BCC = None,
             root = None,
             Vendor = None,
             Name = None,
             CurrentRFQ = None,
             Attachment = None,
             Subject = None,
             Body = None):
        
        Outlook = client.Dispatch('Outlook.Application')
        NewRFQMail = Outlook.CreateItem(0)
        NewRFQMail.To = TO
        if CC:
            NewRFQMail.CC = CC
        if BCC:
            NewRFQMail.BCC =  BCC
        if Vendor:
            _Vendor = Vendor
        else:
            _Vendor = ''
        if Name:
            _Name = Name
        else:
            _Name = ''
        if CurrentRFQ:
            _CurrentRFQ = CurrentRFQ
        else:
            _CurrentRFQ = ''
            
        if Attachment:
            _AttachmentDict = Attachment
        else:
            _AttachmentDict = {}
            
        PartTypeList = ['STANDARD PART', 'FABRICATION PART','GENERAL SUPPLY']
        PartType = PartTypeList[0]
        if Subject:
            NewRFQMail.Subject = Subject
        else: 
            NewRFQMail.Subject = f"""MOTIONWELL - RFQ : {_CurrentRFQ} - {PartType} """
        if Body:
            NewRFQMail.Body = Body
        else: NewRFQMail.Body = f"""
Dear {_Vendor},

Please help to quote with your best price for the item attached in the CSV File provided

Thank you!



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

The contents of this e-mail message and any attachments are confidential and are intended solely for addressee. The information may also be legally privileged. This transmission is sent in trust, for the sole purpose of delivery to the intended recipient. If you have received this transmission in error, any use, reproduction or dissemination of this transmission is strictly prohibited. If you are not the intended recipient, please immediately notify the sender by reply e-mail or phone and delete this message and its attachments, if any."""
        try:
            
            for Attachment in _AttachmentDict:
                assert os.path.exists(_AttachmentDict[Attachment])
                NewRFQMail.Attachments.Add(Source=_AttachmentDict[Attachment])
            #message.Display(True)
            NewRFQMail.send
            messagebox.showinfo('Issue RFQ',f'{_CurrentRFQ} has been issued')
        except AssertionError as e:
            messagebox.showerror('RFQ Email','Error : {e} \n{_AttachmentDict[Attachment]} does not exist')
        root.destroy()

def CloseTab():
    root.destroy()
    pass

if __name__ == "__main__":
    root = Tk()
    root.title("Request for Quotation") 
    root.state("zoomed")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    global tabNote
    def a():
        EmailRFQWindow(CurrentRFQ = 'RFQP21001010001', Purchaser = 'Zi Wei' ,Vendor = 'Vendor 1 BD')
    tabNote = ttk.Notebook(root)
    tabNote.grid(row=0, column=0, sticky="NSEW") 
    RFQEFrame = Frame(tabNote)
    tabNote.add(RFQEFrame, text="Request for Quotation Email")
    RFQEFrame.columnconfigure(0, weight=1)
    tabNote.select(RFQEFrame)
    Button(RFQEFrame,text = 'press',command = a).pack()
    #EmailRFQWindow()
    root.mainloop()
    
