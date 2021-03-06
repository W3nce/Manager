import os
import pathlib
import win32com.client as client
from tkinter import *

XeroConfigKey = ['XERO_EMAIL','CHROME_DRIVER_LOCATION','CLIENT_ID','CLIENT_SECRET','CLIENT_SECRET2','CONFIGURED']
XERO_EMAIL=None
CHROME_DRIVER_LOCATION=None

CLIENT_ID = "3B806419C88F4E638B2ED6F7876B037E"
CLIENT_SECRET = "J_3AsYU9ZAesIrSBvUrzBtvylcpnrs9LoZIflIZM-kOJCPhK"
CLIENT_SECRET2 = '0Hpkg89SeEG1gEBOlEzYVe23dYRnSVqP7dWrsLjVkAi7ULml'

CONFIGURED = 0

def ConfigureChrome(email = '',driver = ''):
    global output,XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED
    output = '','',''
    
    root = Tk()
    root.iconbitmap("MWA_Icon.ico")
    root.title("Chrome/Xero Configuration")
    root.state("zoomed")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    root.resizable(height = None, width = None)        
    
    def SaveConfiguration():
        global output,XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED
        XERO_EMAIL = emailEnt.get()
        CHROME_DRIVER_LOCATION = driverEnt.get()
        
        if os.path.exists(CHROME_DRIVER_LOCATION):
            output = XERO_EMAIL,CHROME_DRIVER_LOCATION,1
            root.destroy()
        else:
            ErrorLabel['text'] = "Please make sure the Chrome Driver is in specificed FilePath\nPath: "+CHROME_DRIVER_LOCATION
        
        
        
            
    def Close():
        global output,XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED
        if messagebox.askokcancel("Quit", "Do you want to quit?",parent = root):
            root.destroy()
            output = email,driver,0
            
    root.protocol("WM_DELETE_WINDOW", Close)
    
    
    OuterFrame = LabelFrame(root)
    OuterFrame.place(relx=0.5, rely=0.5, anchor=CENTER)
    
    ConfigFrame = LabelFrame(OuterFrame)
    ConfigFrame.pack(anchor = CENTER,pady = 10,padx = 10,ipady = 5,ipadx = 5)
    
    TitleLabel = Label(ConfigFrame, text = "Configure Chrome/Xero", font=("Arial", 12))
    TitleLabel.grid(row=0, column=0, padx=10, pady=(10,5),columnspan = 3,sticky = W)
     
    emailEnt = Entry(ConfigFrame, width=35)
    emailEnt.grid(row=1, column=2, padx=20, pady=(10,0),sticky = W)
    emailEnt.insert(END,email if email else '')
    emailLabel = Label(ConfigFrame, text="XERO_EMAIL")
    emailLabel.grid(row=1, column=0, padx=10, pady=(10,0),sticky = W)
    Label(ConfigFrame, text=":", font=("Arial", 10)).grid(row=1, column=1,pady=(10,0),sticky = W)
    
    driverEnt = Entry(ConfigFrame, width=35)
    driverEnt.grid(row=2, column=2, padx=20, pady=(10,0),sticky = W)
    driverEnt.insert(END,driver if driver else '')
    driverLabel = Label(ConfigFrame, text="CHROME_DRIVER_LOCATION")
    driverLabel.grid(row=2, column=0, padx=10, pady=(10,0),sticky = W)
    Label(ConfigFrame, text=":", font=("Arial", 10)).grid(row=3, column=1,pady=(10,0),sticky = W)
    
    SaveButton = Button(ConfigFrame, text = "SAVE", command = SaveConfiguration , width = 8)
    SaveButton.grid(row=2, column=3, columnspan = 2, padx=10, pady=(10,0),sticky = NSEW)
    
    CancelButton = Button(ConfigFrame, text = "CANCEL", command = Close , width = 8)
    CancelButton.grid(row=1, column=3, columnspan = 2, padx=10, pady=(10,0),sticky = NSEW)
    
    ErrorFrame = Frame(ConfigFrame)
    ErrorFrame.grid(row=3, column=0, columnspan = 4, padx=10, pady=(10,0),sticky = NSEW)
    
    ErrorLabel = Label(ErrorFrame, text = '',font = 'Helvetica 9 bold',fg = 'red')
    ErrorLabel.pack()
    
    root.mainloop()
    return output
   
def OverWrite(Reconfigure = False): 
    global XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED,CLIENT_ID,CLIENT_SECRET,CLIENT_SECRET2

    try:
        if 'ChromeConfig.txt' in os.listdir():
            
            with open('ChromeConfig.txt', 'r+') as ChromeConfig:
                ChromeConfig.seek(0)
                text = ChromeConfig.readlines()
                for line in text:
                    
                    if XeroConfigKey[0] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        XERO_EMAIL = line[start:end]
                        continue
                    if XeroConfigKey[1] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        CHROME_DRIVER_LOCATION = line[start:end]
                        continue
                        
                    if XeroConfigKey[2] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        CLIENT_ID = line[start:end]
                        continue
                    
                                        
                    if XeroConfigKey[4] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        CLIENT_SECRET2 = line[start:end]
                        continue
                        
                    if XeroConfigKey[3] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        CLIENT_SECRET = line[start:end]
                        continue
                        

                                        
                    if XeroConfigKey[5] in line:
                        start = line.find('<<') + len('<<')
                        end = line.find('>>')
                        CONFIGURED = line[start:end]
                        
                        if int(line[start:end]) == 0 or Reconfigure:
                            XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED = ConfigureChrome(email = XERO_EMAIL,driver = os.path.abspath("chromedriverv91.exe") if any([CHROME_DRIVER_LOCATION == 'None', CHROME_DRIVER_LOCATION == None]) else CHROME_DRIVER_LOCATION)
                            text = [f"{XeroConfigKey[0]} = <<{XERO_EMAIL}>>", 
                                    f"\n{XeroConfigKey[1]} = <<{CHROME_DRIVER_LOCATION}>>", 
                                    f"\n{XeroConfigKey[2]} = <<{CLIENT_ID}>>", 
                                    f"\n{XeroConfigKey[3]} = <<{CLIENT_SECRET}>>", 
                                    f"\n{XeroConfigKey[4]} = <<{CLIENT_SECRET2}>>",
                                    f"\n{XeroConfigKey[5]} = <<{CONFIGURED}>>"]
                            
                            ChromeConfig.seek(0)
                            ChromeConfig.writelines(text)
                            break
                        else:
                            break
                ChromeConfig.close()
                print(f"{XeroConfigKey[0]} = <<{XERO_EMAIL}>>", 
                        f"\n{XeroConfigKey[1]} = <<{CHROME_DRIVER_LOCATION}>>", 
                        f"\n{XeroConfigKey[2]} = <<{CLIENT_ID}>>", 
                        f"\n{XeroConfigKey[3]} = <<{CLIENT_SECRET}>>", 
                        f"\n{XeroConfigKey[4]} = <<{CLIENT_SECRET2}>>",
                        f"\n{XeroConfigKey[5]} = <<{CONFIGURED}>>")
                
        else:
            with open('ChromeConfig.txt', 'w+') as ChromeConfig:
                XERO_EMAIL,CHROME_DRIVER_LOCATION,CONFIGURED = ConfigureChrome(email = XERO_EMAIL, driver = os.path.abspath("chromedriverv91.exe") if any([CHROME_DRIVER_LOCATION == 'None', CHROME_DRIVER_LOCATION == None])  else CHROME_DRIVER_LOCATION)
                
                text = [f"{XeroConfigKey[0]} = <<{XERO_EMAIL}>>", 
                        f"\n{XeroConfigKey[1]} = <<{CHROME_DRIVER_LOCATION}>>", 
                        f"\n{XeroConfigKey[2]} = <<{CLIENT_ID}>>", 
                        f"\n{XeroConfigKey[3]} = <<{CLIENT_SECRET}>>", 
                        f"\n{XeroConfigKey[4]} = <<{CLIENT_SECRET2}>>",
                        f"\n{XeroConfigKey[5]} = <<{CONFIGURED}>>"]
                
                print(f"{XeroConfigKey[0]} = <<{XERO_EMAIL}>>", 
                      f"\n{XeroConfigKey[1]} = <<{CHROME_DRIVER_LOCATION}>>", 
                      f"\n{XeroConfigKey[2]} = <<{CLIENT_ID}>>", 
                      f"\n{XeroConfigKey[3]} = <<{CLIENT_SECRET}>>", 
                      f"\n{XeroConfigKey[4]} = <<{CLIENT_SECRET2}>>",
                      f"\n{XeroConfigKey[5]} = <<{CONFIGURED}>>")
                ChromeConfig.writelines(text) 
                ChromeConfig.close()
    
        
    except FileNotFoundError as e:
        print(e)


OverWrite()


# =============================================================================
# XERO_EMAIL = 'wencengcs@gmail.com'
# CHROME_LOCATION = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
# CHROME_DRIVER_LOCATION = 'C:\\XERO-API\\xerogui\\chromedriverv91.exe'
# =============================================================================