from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox 
from tkcalendar import *
from datetime import datetime
from PIL import ImageTk, Image
from mysql import *
from mysql.connector import connect, Error
import csv
import ConnConfig

logininfo = (ConnConfig.host,ConnConfig.username,ConnConfig.password)

connInit = connect(host = logininfo[0],
                    user = logininfo[1], 
                    password =logininfo[2])
 
curInit = connInit.cursor()

SetupCommand = ["""
                CREATE SCHEMA IF NOT EXISTS `INDEX_PRO_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci""",
                    
                """
                CREATE TABLE IF NOT EXISTS `INDEX_PRO_MASTER`.`PROJECT_INFO`
                (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                 `PROJECT_CLS_ID` VARCHAR(10),
                 `PROJECT_NAME` VARCHAR(50),
                 `PROJECT_MANAGER` INT,
                 `PROJECT_SUPPORT` INT,
                 `START_DATE` DATE,
                 `DELIVERY_DATE` DATE,
                 `BASE_PLAN_LOCK` BOOL,
                 `PROJECT_LOCK` BOOL,
                 `STATUS` INT,
                 `NUM_DELETED` INT DEFAULT 0,
                 `NUM_ONHOLD` INT DEFAULT 0,
                 `EST_PART` INT DEFAULT 0,
                 `EST_COST` FLOAT DEFAULT 0.0,
                 `DATE_OF_ENTRY` DATE NOT NULL)
                
                ENGINE = InnoDB
                DEFAULT CHARACTER SET = utf8mb4
                COLLATE = utf8mb4_0900_ai_ci""",
                
                """
                CREATE USER if not exists  
                'Login'@'%' IDENTIFIED BY 'password' """,
                
                """
                GRANT SELECT,CREATE, UPDATE,INSERT,DELETE ON 
                `index_emp_master`.* To 'Login'@'%'"""]



for com in SetupCommand:
    
    curInit.execute(com)
    connInit.commit()
    
curInit.close()


connLogin = connect(host = logininfo[0],
                    user = logininfo[1], 
                    password =logininfo[2])

SetupCommand = ["""
                CREATE SCHEMA IF NOT EXISTS `INDEX_EMP_MASTER` 
                DEFAULT CHARACTER SET utf8mb4 
                COLLATE utf8mb4_0900_ai_ci""",
                """
                CREATE TABLE IF NOT EXISTS `INDEX_EMP_MASTER`.`EMP_DATA` 
                (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                 `EMPLOYEE_ID` VARCHAR(10),
                 `EMPLOYEE_NAME` VARCHAR(60),
                 `EMPLOYEE_CLASS` INT,
                 `NATIONALITY` VARCHAR(30),
                 `DOB` DATE,
                 `ADDRESS` VARCHAR(120),
                 `JOIN_DATE` DATE,
                 `STATUS` INT)

                 ENGINE = InnoDB
                 DEFAULT CHARACTER SET = utf8mb4
                 COLLATE = utf8mb4_0900_ai_ci""",
                 
                 """
                 CREATE TABLE IF NOT EXISTS `index_emp_master`.`login_data` (
                 `oid` INT NOT NULL AUTO_INCREMENT,
                 `EMPLOYEE_ID` VARCHAR(10) NOT NULL,
                 `EMPLOYEE_CLASS` INT NOT NULL,
                 `EMPLOYEE_PW` VARCHAR(255) NOT NULL DEFAULT  ( SHA1('MWA2021')),
                 PRIMARY KEY (`oid`))
                ENGINE = InnoDB
                AUTO_INCREMENT = 1
                DEFAULT CHARACTER SET = utf8mb4
                COLLATE = utf8mb4_0900_ai_ci"""
                ,
                """
                INSERT IGNORE INTO `index_emp_master`.`emp_data`
                (oid, EMPLOYEE_ID, EMPLOYEE_NAME, EMPLOYEE_CLASS, 
                 NATIONALITY, DOB, ADDRESS, JOIN_DATE, STATUS) 
                VALUES(1,'ADM01', 'ADMIN 1', 3, 'NIL', 
                 '2021-06-17', 'NIL', '2021-06-17', 1) """,
                 
                 """
                 INSERT IGNORE INTO `index_emp_master`.`login_data`
                 (oid, EMPLOYEE_ID, EMPLOYEE_CLASS,EMPLOYEE_PW) 
                 VALUES(1,'ADM01',3,SHA1('MWA2021'))
                 """
                 ,
                 """CREATE SCHEMA IF NOT EXISTS `PUR_ORDER_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci"""
                    ,
                    
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
                    COLLATE = utf8mb4_0900_ai_ci"""
                    ,
                    
                    """CREATE SCHEMA IF NOT EXISTS `COMPANY_INFO` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci"""
                    ,
                    
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
                    ,
                 """CREATE SCHEMA IF NOT EXISTS `CURRENCY_MASTER` 
                    DEFAULT CHARACTER SET utf8mb4 
                    COLLATE utf8mb4_0900_ai_ci"""
                    ,
                    
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
             
                
             
curSetup = connLogin.cursor()
for command in SetupCommand:
    curSetup.execute(command)
    connLogin.commit()
curSetup.close()
                
curCom = connLogin.cursor()
curCom.execute("SELECT * FROM `COMPANY_INFO`.`COMPANY_MWA`")
existLst = curCom.fetchall()

if existLst == []:
    defaultComSql = f"""INSERT INTO `COMPANY_INFO`.`COMPANY_MWA` (
    ComName, ComRegNum, GSTRegNum, Address, PosCode, 
    CenterA, CenterB, ContactNum, Email)

    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                
    defaultComInfo = ("Motionwell Automation Pte. Ltd.", "201435019E", "201435019E",
                      "20 Woodlands Link", "738733", "#09-08 (Design & Assembly Center)", 
                      "#07-26 (Manufacturing Shop)", "98506102", "info@motionwell.com.sg")
    
    curCom.execute(defaultComSql, defaultComInfo)
    connLogin.commit()
curCom.close()
                     
curCcy = connLogin.cursor()
curCcy.execute("SELECT * FROM `CURRENCY_MASTER`.`CCY_LIST`")
existLst = curCcy.fetchall()

if existLst == []:
    defaultCcy = f"""INSERT INTO `CURRENCY_MASTER`.`CCY_LIST` (
    CcyAcro, CcyFullName, ExRate, DateUpdated, Remark)
    VALUES (%s, %s, %s, %s, %s)"""
                
    defaultLst = [("SGD", "Singapore Dollar", 1.00, None, ""),
                  ("USD", "United States Dollar", 1.34, None, ""),
                  ("CNY", "Chinese Yuan (Remminbi)", 0.21, None, ""),
                  ("HKD", "Hong Kong Dollar", 0.17, None, ""),
                  ("EUR", "European Euro", 1.60, None, ""),
                  ("JPY", "Japanese Yen", 0.012, None, ""),
                  ("GBP", "British Pound Sterling", 1.87, None, ""),
                  ("AUD", "Australian Dollar", 1.02, None, ""),
                  ("CAD", "Canadian Dollar", 1.09, None, ""),
                  ("CHF", "Swiss Franc", 1.46, None, "")]
    
    for ccy in defaultLst:
        curCcy.execute(defaultCcy, ccy)
    connLogin.commit()
curCcy.close()



curCls = connLogin.cursor()
curCls.execute("SELECT * FROM `INDEX_VEND_MASTER`.`CLASS_LIST`")
existLst = curCls.fetchall()

if existLst == []:
    defaultCcy = f"""INSERT INTO `INDEX_VEND_MASTER`.`CLASS_LIST` (
    CLASS, DESCRIPTION)
    VALUES (%s, %s)"""
                
    defaultLst = [("N", "Normal"),
                  ("FG", "Fabrication - General Tolerance"),
                  ("FOT", "Fabrication - Others"),
                  ("FP", "Fabrication - Precision"),
                  ("FPF", "Fabrication - Alu Profile Work"),
                  ("FPL", "Fabrication - Polymer"),
                  ("FPW", "Fabrication - Plastic Sheet Work"),
                  ("FSC", "Fabrication - Signcraft"),
                  ("FSH", "Fabrication - Sheet Metal Work"),
                  ("FSP", "Fabrication - Semi Precision"),
                  ("FWD", "Fabrication - Welded Structures"),
                  ("FWW", "Fabrication - Woodwork"),
                  ("GSCE", "General Supply - Computer Hardware"),
                  ("GSHR", "General Supply - Human Resources"),
                  ("GSOM", "General Supply - Office Maintenance"),
                  ("GSOS", "General Supply - Office Supply"),
                  ("GSPE", "General Supply - Project Equipment"),
                  ("GSPM", "General Supply - Project Material"),
                  ("GSPS", "General Supply - Project Services"),
                  ("GSTP", "General Supply - Transportation"),
                  ("MAT", "Fabrication - Meterial Treatment")]
    
    for clsVal in defaultLst:
        curCls.execute(defaultCcy, clsVal)
    connLogin.commit()
curCls.close()


             
connVendInit = connect(host = logininfo[0],
                           user = logininfo[1], 
                           password =logininfo[2])
    
    
    
CreateVendDatabase = ["""
                      CREATE SCHEMA IF NOT EXISTS `INDEX_VEND_MASTER` 
                      DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci """,
                      
                      """
                      CREATE TABLE IF NOT EXISTS `INDEX_VEND_MASTER`.`VENDOR_LIST`
                      (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                       `VENDOR_NAME` VARCHAR(20),
                       `CLASS` VARCHAR(10),
                       `COMPANY_NAME` VARCHAR(100),
                       `COMPANY_COUNTRY` VARCHAR(50),
                       `COMPANY_STATE` VARCHAR(50),
                       `COMPANY_CITY` VARCHAR(50),
                       `POSTAL_CODE` VARCHAR(30),
                       `COMPANY_ADDRESS_A` VARCHAR(100),
                       `COMPANY_ADDRESS_B` VARCHAR(100),
                       `VENDOR_DESCRIPTION` VARCHAR(100),
                       `MAIN_CONTACT_PERSON` VARCHAR(100),
                       `MAIN_CONTACT_NUMBER` VARCHAR(20),
                       `MAIN_CONTACT_EMAIL` VARCHAR(100),
                       `SECONDARY_CONTACT_PERSON` VARCHAR(100),
                       `SECONDARY_CONTACT_NUMBER` VARCHAR(20),
                       `SECONDARY_CONTACT_EMAIL` VARCHAR(100),
                       `DATE_OF_ENTRY` DATETIME NOT NULL,
                       `STATUS` INT)
                      
                      ENGINE = InnoDB
                      DEFAULT CHARACTER SET = utf8mb4
                      COLLATE = utf8mb4_0900_ai_ci""",
                      
                      """
                      CREATE TABLE IF NOT EXISTS `INDEX_VEND_MASTER`.`CLASS_LIST`
                      (`oid` INT AUTO_INCREMENT PRIMARY KEY,
                       `CLASS` VARCHAR(10),
                       `DESCRIPTION` VARCHAR(100))
                      
                      ENGINE = InnoDB
                      DEFAULT CHARACTER SET = utf8mb4
                      COLLATE = utf8mb4_0900_ai_ci
                      """]

curVendInit = connVendInit.cursor()

for command in CreateVendDatabase:
    curVendInit.execute(command)
    connVendInit.commit()

curVendInit.close()                 
                 


AUTH = 0
AUTHLVL = 0

root = Tk()
root.title("Management")
root.state("zoomed")
root.iconbitmap("MWA_Icon.ico")
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)
root.resizable(height = None, width = None)

OuterFrame = LabelFrame(root)
OuterFrame.place(relx=0.5, rely=0.5, anchor=CENTER)

LoginFrame = LabelFrame(OuterFrame)
LoginFrame.pack(anchor = CENTER,pady = 10,padx = 10,ipady = 5,ipadx = 5)

LoginLabel = Label(LoginFrame, text = "Database Login", font=("Arial", 12))
LoginLabel.grid(row=0, column=0, padx=10, pady=(10,5),columnspan = 3,sticky = W)
 
User = Entry(LoginFrame, width=18)
User.grid(row=1, column=1, padx=20, pady=(10,0),sticky = W)
UserLabel = Label(LoginFrame, text="User")
UserLabel.grid(row=1, column=0, padx=10, pady=(10,0),sticky = W)
Label(LoginFrame, text=":", font=("Arial", 10)).grid(row=1, column=1,pady=(10,0),sticky = W)

Password = Entry(LoginFrame, show="*", width=18)
Password.grid(row=2, column=1, padx=20, pady=(10,0),sticky = W)
PasswordLabel = Label(LoginFrame, text="Password")
PasswordLabel.grid(row=2, column=0, padx=10, pady=(10,0),sticky = W)
Label(LoginFrame, text=":", font=("Arial", 10)).grid(row=2, column=1,pady=(10,0),sticky = W)

ResetPasswordLabel = Label(LoginFrame, text="Reset Password", cursor="hand2",fg="blue")
ResetPasswordLabel.grid(row=3, column=0, padx=10, pady=(10,0),sticky = W,columnspan = 2)
ResetPasswordLabel.bind("<Button-1>", lambda e: print('Contact Admin'))

global LOCKEDUSER 
LOCKEDUSER = None

def Login():
    global AUTH,AUTHLVL,LOCKEDUSER
    LOCKEDUSER = (User.get(),Password.get())
    def ChangePW():
        ChangePWWin = Tk()
        ChangePWWin.title("Create New Password")
        ChangePWWin.geometry("460x300")
        ChangePWWin.attributes('-topmost', 1)
        
        TitleLabel = Label(ChangePWWin, text = "Reset Password", font=("Arial", 10))
        TitleLabel.grid(row=0, column=0,padx=10, pady=10,columnspan  = 3 , sticky=W) 
        
        CurrentPassword = Entry(ChangePWWin, show="*", width=18)
        CurrentPassword.grid(row=1, column=2, padx=20, pady=(10,0),sticky = W)
        CurrentPasswordLabel = Label(ChangePWWin, text="Current Password")
        CurrentPasswordLabel.grid(row=1, column=0, padx=10, pady=(10,0),sticky = W)
        Label(ChangePWWin, text=":", font=("Arial", 10)).grid(row=1, column=1,pady=(10,0),sticky = W)
        
        ChangePassword = Entry(ChangePWWin, show="*", width=18)
        ChangePassword.grid(row=2, column=2, padx=20, pady=(10,0),sticky = W)
        ChangePasswordLabel = Label(ChangePWWin, text="New Password")
        ChangePasswordLabel.grid(row=2, column=0, padx=10, pady=(10,0),sticky = W)
        Label(ChangePWWin, text=":", font=("Arial", 10)).grid(row=2, column=1,pady=(10,0),sticky = W)
        
        ConfirmPassword = Entry(ChangePWWin, show="*", width=18)
        ConfirmPassword.grid(row=3, column=2, padx=20, pady=(10,0),sticky = W)
        ConfirmPasswordLabel = Label(ChangePWWin, text="Confirm Password")
        ConfirmPasswordLabel.grid(row=3, column=0, padx=10, pady=(10,0),sticky = W)
        Label(ChangePWWin, text=":", font=("Arial", 10)).grid(row=3, column=1,pady=(10,0),sticky = W)
        
        ErrorFrame = Frame(ChangePWWin)
        ErrorFrame.grid(row=4, column=0, padx=10, pady=(10,0),sticky = W,columnspan = 2)
        
        MismatchLabel = Label(ErrorFrame, text = 'Password Mismatch/Wrong Password', font = 'Helvetica 8 bold', fg = 'red')
        
        def ConfirmChangePW():
            global LOCKEDUSER
            Cursor = connLogin.cursor()
            Cursor.execute(f"""SELECT * FROM `index_emp_master`.`login_data` 
                             WHERE `EMPLOYEE_ID` = '{LOCKEDUSER[0]}' """)
            res = Cursor.fetchall()[0]
            Cursor.execute(f"""SELECT SHA1('{CurrentPassword.get()}')""")
            CurrPW = Cursor.fetchall()[0]
            if res[3] == CurrPW[0] and ChangePassword.get() == ConfirmPassword.get():
                try:
                    Cursor.execute(f"""
                                   UPDATE `index_emp_master`.`login_data`
                                   SET
                                   EMPLOYEE_PW = SHA1('{ConfirmPassword.get()}')
                                   WHERE EMPLOYEE_ID = '{LOCKEDUSER[0]}'
                                   """)
                    connLogin.commit()
                    Popup = messagebox.showinfo("Password saved", "Password changed successfully", parent = ChangePWWin)
                    ChangePWWin.destroy()
                except Error as e:
                    connLogin.rollback()
                    Popup = messagebox.showerror("Database Error", "Failed to reset password", parent = ChangePWWin)
            else:
                MismatchLabel.pack()
             
                
            Cursor.close()
            
        ConfirmChangePWButton = Button(ChangePWWin,text = 'Change Password',command = ConfirmChangePW)
        ConfirmChangePWButton.grid(row=4, column=2, padx=10, pady=(10,0),sticky = W)
        
        ChangePWWin.mainloop() 
        
        
    try:
        Cursor = connLogin.cursor()
                
        Cursor.execute(f"""SELECT * FROM `index_emp_master`.`login_data` 
                             WHERE EMPLOYEE_ID = '{User.get()}'
                                     """)
        res = Cursor.fetchall()[0]
        Cursor.execute(f"""SELECT SHA1('{Password.get()}')""")
        PW = Cursor.fetchall()[0]
        #print(PW[0] ,res[0][3] )
        
        if res:
            if res[3] == PW[0]:
                Popup = messagebox.showinfo("Authentication Success", "Database Connected",parent = root)
                AUTH = 1
                AUTHLVL = res[2]
                root.destroy()
                if LOCKEDUSER[1] == 'MWA2021':
                    ChangePW()
            else:
                Popup = messagebox.showerror("Authentication Failed", "Please enter a valid password",parent = root)
            
        else:
            Popup = showerror.showinfo("Authentication Failed", "Please enter a valid username",parent = root)

    
    except Error as e:
        
        Popup = messagebox.showerror("Database Error", "Please enter the valid credentials",parent = root)
        
    
    Cursor.close()
     


LoginButton = Button(LoginFrame, text = "LOGIN", command = Login , width = 8)
LoginButton.grid(row=1, column=2, rowspan =2, columnspan = 2, padx=10, pady=(10,0),sticky = NSEW)

root.mainloop()

connLogin.close()