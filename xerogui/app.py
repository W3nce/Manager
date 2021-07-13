#GUI|API connector

from tkinter import *
# from ttkwidgets import CheckboxTreeview
from tkinter import ttk
#from xerogui.config import OverWrite,CLIENT_ID,CLIENT_SECRET,XERO_EMAIL,CHROME_DRIVER_LOCATION
from xerogui.config import CLIENT_ID,CLIENT_SECRET,CLIENT_SECRET2,XERO_EMAIL,CHROME_DRIVER_LOCATION,XeroConfigKey
import pprint as pp


import requests
from requests.exceptions import Timeout
import base64
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

import time

LoginURL = 'https://login.xero.com/identity/connect/authorize?response_type=code&client_id=YOURCLIENTID&redirect_uri=YOURREDIRECTURI&scope= offline_access openid profile email accounting.transactions&state=123'

#https://login.xero.com/identity/connect/authorize/callback?client_id=xero_business_xeroweb_hybrid-web&redirect_uri=https%3A%2F%2Fgo.xero.com%2Fsignin-oidc&response_mode=form_post&response_type=code%20id_token&scope=xero_all-apis%20openid%20email%20profile&state=OpenIdConnect.AuthenticationProperties%3D1W88diWb5i9q8jFKqLqTq8JQw_N8bcgjb5kbZYNAPIXXosrquYCy7d7ci6UKm6usSKDMfq05wOhm04CyVsJfirpF5_QLcWMy2oLuiVd_ktA9GFfFYVEdGIA_HHhvw-B6Qk8_apIMBXXp1MwxP4HeJGjMaGP8UKt3cDy7DIdrF5aderu-QLMkTuC2iF4PYOxmSo6IUcPgmjpQ5-JOP8TQjaZY6kA&nonce=637601966675361249.OWUwNjIxYTQtZjRhOS00YzNhLWJlZGMtMmU1MDNiNGY2ZmQ4NTM2Y2M2OGQtZTE5YS00MGU5LWJlNjUtMTA3MWUyM2I2MTAz&x-client-SKU=ID_NET451&x-client-ver=1.3.14.0

TEMP_XERO_EMAIL = None

client_id =  CLIENT_ID
client_secret = CLIENT_SECRET
redirect_url = 'https://xero.com/'
scope = 'offline_access accounting.transactions accounting.attachments'
b64_id_secret = base64.b64encode(bytes(client_id + ':' + client_secret, 'utf-8')).decode('utf-8')

LOGIN = 0
tkn = None
ConnectedTENANT = None

#def CheckChrome():
    #try:
      # ChromeConfig = open('refresh_token.txt', 'w')
        #

def XeroFirstAuth():
    global LOGIN
     
    # 1. Send a user to authorize your app   
    
    def GetLoginDetails(Error = None):
        global xeroemail,xeropw,img
        
        xeroemail = None
        xeropw = None
        ReqPW = Tk()
        ReqPW.iconbitmap("MWA_Icon.ico")
        ReqPW.title('Xero Login')
        ReqPW.state("zoomed")
        ReqPW.rowconfigure(0, weight=1)
        ReqPW.columnconfigure(0, weight=1)
        ReqPW.resizable(height = None, width = None) 
        
        BorderFrame = Frame(ReqPW,width = 800, height = 400)
        BorderFrame.place(relx=0.5, rely=0.5, anchor=CENTER)
        
        canvas = Canvas(BorderFrame, width = 330, height = 150)      
        canvas.pack( anchor = NW) 
        
        img = PhotoImage(file="Xero-LogoRE.png")
        canvas.create_image(15,15, anchor=NW, image=img) 
        
        
        
        OuterFrame = LabelFrame(BorderFrame)
        OuterFrame.pack(pady = 5, anchor = CENTER)
        
        LoginFrame = LabelFrame(OuterFrame)
        LoginFrame.pack(anchor = CENTER,pady = 10,padx = 10,ipady = 10,ipadx = 8)
        def Close():
            global xeroemail,xeropw
            if messagebox.askokcancel("Close Xero Login", "Do you want to continue without XERO?",parent = ReqPW):
               
                driver.quit()
                ReqPW.destroy()
            
        ReqPW.protocol("WM_DELETE_WINDOW", Close)
    
        EnterEmailLabel = Label(LoginFrame, text = 'Xero Email :').grid(row = 0, column = 0,padx=5, pady=5,sticky = E)
        EnterEmailEntry = Entry(LoginFrame, width=24)
        EnterEmailEntry.grid(row=0, column=1, padx=5, pady=10,sticky = W)
        EnterEmailEntry.insert(END, '' if XERO_EMAIL == 'None' else XERO_EMAIL)
        
        EnterPWLabel = Label(LoginFrame, text = 'Xero Password :').grid(row = 1, column = 0,padx=5, pady=5,sticky = E)
        EnterPWEntry = Entry(LoginFrame, show="*", width = 24)
        EnterPWEntry.grid(row=1, column=1, padx=5, pady=10,sticky =W)
        
        Label(LoginFrame, text = f'Error : {Error}',font = 'Helvetica 9 bold',fg = 'red').grid(row = 2, column = 0,columnspan =2 ,padx=5, pady=5,sticky = EW) if Error else None
        
        
    
        def Enter():
            global xeroemail,xeropw,TEMP_XERO_EMAIL
            xeroemail = EnterEmailEntry.get()
            TEMP_XERO_EMAIL = EnterEmailEntry.get()
            xeropw= EnterPWEntry.get()
            ReqPW.destroy()
            
        ButtonFrame = Frame(LoginFrame)
        ButtonFrame.grid(row = 3, column = 0,columnspan = 2) 
        EnterButton = Button(ButtonFrame,text = 'Sign In',command = Enter,width = 14).pack(side = LEFT, padx = 10)#grid(row = 0, column = 0,padx=5, pady=5) 
        CloseButton = Button(ButtonFrame,text = 'Close',command = Close,width = 14).pack(side = RIGHT, padx = 10)#grid(row = 0, column = 1,padx=5, pady=5)
        ReqPW.mainloop()
        return xeroemail,xeropw
    
    auth_url = ('''https://login.xero.com/identity/connect/authorize?''' +
                '''response_type=code''' +
                '''&client_id=''' + client_id +
                '''&redirect_uri=''' + redirect_url +
                '''&scope=''' + scope +
                '''&state=''' + '''OAUTH2''')
    chrome_options = Options()
    chrome_options.headless = True
# =============================================================================
#     driver = webdriver.Chrome(CHROMEDRIVER_PATH, options=options)
#     chrome_options = Options()
#     chrome_options.add_argument("--headless")
#     chrome_options.binary_location = CHROME_LOCATION
# 
# =============================================================================

    driver = webdriver.Chrome(CHROME_DRIVER_LOCATION, options = chrome_options)

    driver.get(auth_url)
    def TryLogin(Error = None):
    
        try : 
            driver.get(auth_url)
            #Title = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/head/meta[4]')))
            if driver.title != 'Login | Xero Accounting Software':
                raise NoSuchElementException
            email = driver.find_element(by = By.ID, value = 'xl-form-email')
            password = driver.find_element(by = By.ID, value = 'xl-form-password')
            loginButton = driver.find_element(by = By.ID, value = "xl-form-submit")
            
        
        except NoSuchElementException:
            print("Xero Webpage Error") #Either wrong page or element not found
            return TryLogin(Error = "Xero Webpage Error")
        except TimeoutException:
            print("Too Much Time Taken")
            return TryLogin(Error = "Too Much Time Taken")
    
        
        xeroemail,xeropw = GetLoginDetails(Error = Error)
        if xeropw == None and xeroemail == None:
            
            return False
        elif xeropw and xeroemail:
            
            email.send_keys(xeroemail)
            password.send_keys(xeropw)  # Tkinter Request password
            loginButton.click()
            print('Login Done')
            return True
        else:
            return True
        
    
        
    # 2. Obtain Auth Code from Url
    def GetAuth():
        auth_code = None
        LoginCount = 3
        while LoginCount >0:   
            try : 
                LoginCount-=1
                state = 'Logging into Tenant Connections'
                print(state)
                
                if driver.title != 'Xero | User Consent':
                    raise NoSuchElementException    
                    
                state = 'Finding Button'
                print(state)
                ButtonCont = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div/form/div/div/div/div/div[3]/div[2]/button[1]")))
                time.sleep(1)
                state  = 'Clicking Button'
                print(state)
                ActionChains(driver).move_to_element(ButtonCont).click(ButtonCont).perform()
                state = 'Retreiving URL'
                print(state)
                time.sleep(1)
                auth_res_url = driver.current_url 
                state = 'URL Retrieved, Closing Browser'
                print(state)
                driver.quit()
                state = 'Retrieving Auth Code'
                print(state)
                start_number = auth_res_url.find('code=') + len('code=')
                end_number = auth_res_url.find('&scope')
        
                auth_code = auth_res_url[start_number:end_number]
                
                break
            
            except NoSuchElementException:
                print(f"Xero Webpage Error While {state}") #Either wrong page or element not found
                print('Please Try Again')
                if TryLogin(f"Xero Webpage Error While {state} (Tries Remaining: {LoginCount})"):
                    continue
                else:
                    break
            except TimeoutException:
                print(f"Too Much Time Taken While {state}")
                print('Please Try Again')
                if TryLogin(f"Too Much Time Taken While {state} (Tries Remaining: {LoginCount})"):
                    continue
                else:
                    break
            except EXCEPTION as e:
                print('Error: {}'.format(str(e)))
        return auth_code   
    auth_code = GetAuth() if TryLogin() else None
    
        
    # 3. Exchange the code
    exchange_code_url = 'https://identity.xero.com/connect/token'
    response = requests.post(exchange_code_url, 
                             headers = {
                                 'Authorization': 'Basic ' + b64_id_secret
                                 },
                             data = {
                                 'grant_type': 'authorization_code',
                                 'code': auth_code,
                                 'redirect_uri': redirect_url
                                 })
    json_response = response.json()
    #pp.pprint([auth_code, json_response])
    print(json_response.keys())
    
    # 4. Receive your tokens
    if 'error' not in json_response.keys():
        global tkn
        with open('ChromeConfig.txt', 'w+') as ChromeConfig:
               
                text = [f"{XeroConfigKey[0]} = <<{TEMP_XERO_EMAIL}>>", 
                        f"\n{XeroConfigKey[1]} = <<{CHROME_DRIVER_LOCATION}>>", 
                        f"\n{XeroConfigKey[2]} = <<{CLIENT_ID}>>", 
                        f"\n{XeroConfigKey[3]} = <<{CLIENT_SECRET}>>", 
                        f"\n{XeroConfigKey[4]} = <<{CLIENT_SECRET2}>>",
                        f"\n{XeroConfigKey[5]} = <<1>>"]
                
                ChromeConfig.writelines(text) 
                ChromeConfig.close()
        LOGIN = 1
        tkn =  [json_response['access_token'], json_response['refresh_token']]
        return tkn
    
def XeroTenants(access_token):
    global ConnectedTENANT
    connections_url = 'https://api.xero.com/connections'
    response = requests.get(connections_url,
                           headers = {
                               'Authorization': 'Bearer ' + access_token,
                               'Content-Type': 'application/json'
                           })
    json_response = response.json()
    print('\nTENANTS CONNECTED :')
    tenantdict = {tenant['tenantName']:tenant['tenantId'] for tenant in json_response}
       
    def GetTenant():
        global ConnectedTENANT
        ConnectedTENANT = None
        TenantWin = Tk()
        TenantWin.title("Tenant Connections")
        TenantWin.rowconfigure(0, weight=1)
        TenantWin.columnconfigure(0, weight=1)
        Label(TenantWin,text = 'Choose Tenant to Connect').grid(row= 0 , column = 0,padx = 10,pady = 5)
        Label(TenantWin,text = 'Tenant :').grid(row= 1 , column = 0,padx = 10,pady = 5)
        tenantComb = ttk.Combobox(TenantWin, width=30, value=list(tenantdict.keys()), state="readonly")
        tenantComb.current(0)
        tenantComb.grid(row= 1 , column = 0,padx = 10,pady = 5)

        def ChooseTenant():
            global ConnectedTENANT
            
            ConnectedTENANT = tenantComb.get()
            TenantWin.destroy()

        Button(TenantWin,text = 'Connect', command = ChooseTenant).grid(row =2, column = 0,columnspan = 2,sticky = EW)
            
        TenantWin.mainloop()
        
        return ConnectedTENANT
    

    return tenantdict[ConnectedTENANT if ConnectedTENANT else GetTenant()]


def XeroRefreshToken(refresh_token):
    global LOGIN
    token_refresh_url = 'https://identity.xero.com/connect/token'
    response = requests.post(token_refresh_url,
                            headers = {
                                'Authorization' : 'Basic ' + b64_id_secret,
                                'Content-Type': 'application/x-www-form-urlencoded'
                            },
                            data = {
                                'grant_type' : 'refresh_token',
                                'refresh_token' : refresh_token
                            })
    json_response = response.json()
    print('\nREFRESH')
    
    new_refresh_token = json_response['refresh_token']
    rt_file = open('refresh_token.txt', 'w')
    rt_file.write(new_refresh_token)
    rt_file.close()
    LOGIN = 1
    
    return [json_response['access_token'], json_response['refresh_token']]

def XeroRequestsBase(URL,method,json = None):
    global LOGIN,tkn
    if LOGIN and tkn:
        old_refresh_token = open('refresh_token.txt', 'r').read()
        new_tokens = XeroRefreshToken(old_refresh_token)
        xero_tenant_id = XeroTenants(new_tokens[0])
    else:
        
        tkn = XeroFirstAuth()
        new_tokens = XeroRefreshToken(tkn[1])
        xero_tenant_id = XeroTenants(tkn[0])
        
    get_url = URL
    response = requests.request(method,get_url,
                           headers = {
                               'Authorization': 'Bearer ' + new_tokens[0],
                               'Xero-tenant-id': xero_tenant_id,
                               'Accept': 'application/json',
                               'Content-Type': 'application/json'
                           },
                           json = json if json else None
                           )
    json_response = response.json()
    print('\nREQUEST')
    return(json_response)
  
def XeroSuppliersGetAll():

    ContactList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/Contacts','GET')
    print('\nGET ALL CONTACTS')
    #print([Contact["IsSupplier"] for Contact in ContactList['Contacts']])
    return {Contact['Name'] if Contact["IsSupplier"] == True else None: Contact['ContactID'] for Contact in ContactList['Contacts']}

    
def XeroContactsGetAll():

    ContactList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/Contacts','GET')
    print('\nGET ALL CONTACTS')
    #print([Contact["IsSupplier"] for Contact in ContactList['Contacts']])
    return {Contact['Name'] : Contact['ContactID'] for Contact in ContactList['Contacts']}

def XeroContactsGetAllInfo():

    ContactList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/Contacts','GET')
    print('\nGET ALL CONTACTS')
    return ContactList
    #{Contact['Name'] : Contact['ContactID']for Contact in ContactList['Contacts']}




def XeroContactsGet(ContactID):
    ContactList = XeroRequestsBase(f'https://api.xero.com/api.xro/2.0/Contacts/{ContactID}',
                                   'GET')
    print('\nGET ALL CONTACTS')
    #pp.pprint({Contact['Name'] : Contact['ContactID']for Contact in ContactList['Contacts']})
    return {Contact['Name'] : Contact['ContactID']for Contact in ContactList['Contacts']}


def XeroThemesGetAll():

    ThemeList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/BrandingThemes','GET')
    print('\nGET ALL THEMES')
    pp.pprint([Theme['Name'] for Theme in ThemeList['BrandingThemes']])
    return {Theme['Name'] : Theme['BrandingThemeID'] for Theme in ThemeList['BrandingThemes']}

def XeroTaxRatesGetAll():
    TaxRatesList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/TaxRates','GET')
    print('\nGET ALL TAX RATES')
    pp.pprint([TaxRate['TaxType'] for TaxRate in TaxRatesList['TaxRates']])
    return {TaxRate['Name'] : TaxRate['TaxType'] for TaxRate in TaxRatesList['TaxRates']}
    
def XeroAccountCodeGetAll():
    AccountCodeList = XeroRequestsBase('https://api.xero.com/api.xro/2.0/Accounts','GET')
    
    print('\nGET ALL ACCOUNT CODE')
    pp.pprint([f"{AccountCode['Code']} - {AccountCode['Class']}" if 'Code' in AccountCode.keys() else None for AccountCode in AccountCodeList['Accounts']])
    dic = {}
    for AccountCode in AccountCodeList['Accounts']:
        if 'Code' in AccountCode.keys():
            dic.update({f"{AccountCode['Code']} - {AccountCode['Name']}" : (AccountCode['Code'], AccountCode['TaxType'])})
    return dic
    
def XeroPOPost(jsoninp = None):   
    #XeroPOPost(jsoninp = POS.to_dict()) ==> POS([PO([LineItem()])])
    jsoninput = None
    if jsoninp:
        jsoninput =jsoninp
    else:
        return 'Input Error'
    
    POlist = XeroRequestsBase('https://api.xero.com/api.xro/2.0/PurchaseOrders?summarizeErrors=false','POST',
                                   json = jsoninput)
    print('\nPOST PO')
    return POlist

    
def XeroPOSPost(jsoninp = None):   #Mutliple POS
    jsoninput = None
    if jsoninp:
        jsoninput =jsoninp
    else:
        return 'Input Error'
    
    POlist = XeroRequestsBase('https://api.xero.com/api.xro/2.0/PurchaseOrders?summarizeErrors=false','POST',
                              json = jsoninput)
    print('\nPOST POS')
    return POlist


def XeroPOGetAll():
    POlist = XeroRequestsBase('https://api.xero.com/api.xro/2.0/PurchaseOrders','GET')
    print('\nGET ALL PO')
    pp.pprint(POlist)
    return [jsonPO["PurchaseOrderNumber"] for jsonPO in POlist["PurchaseOrders"]]

def XeroPOGet(PurchaseOrderID):
    POlist = XeroRequestsBase(f'https://api.xero.com/api.xro/2.0/PurchaseOrders/{PurchaseOrderID}','GET')
    print('\nGET PO')
    pp.pprint(POlist)
    return [jsonPO["PurchaseOrderNumber"] for jsonPO in POlist["PurchaseOrders"]]

def XeroPODelete():
    pass

def RetrieveToken():
    global tkn,LOGIN,ConnectedTENANT
    try:
        if LOGIN and tkn and ConnectedTENANT:
            tkn = XeroRefreshToken(tkn[1])
            tenant_id = XeroTenants(tkn[0])
        else:
            tkn = XeroFirstAuth()

            
            if not tkn:
                print('Login Unsuccessful')

                return None
            
            tkn = XeroRefreshToken(tkn[1])
            tenant_id = XeroTenants(tkn[0])
        return tkn
    except Exception as e:
        print(e)