from datetime import datetime
class VendorClass():


        
    
    attribute_map = {
        "_VENDOR_NAME": "VENDOR_NAME",
        "_CLASS": "CLASS",
        "_COMPANY_NAME": "COMPANY_NAME",
        "_COMPANY_COUNTRY": "COMPANY_COUNTRY",
        "_COMPANY_STATE": "COMPANY_STATE",
        "_COMPANY_CITY": "COMPANY_CITY",
        "_POSTAL_CODE": "POSTAL_CODE",
        "_COMPANY_ADDRESS_A": "COMPANY_ADDRESS_A",
        "_COMPANY_ADDRESS_B": "COMPANY_ADDRESS_B",
        "_VENDOR_DESCRIPTION": "VENDOR_DESCRIPTION",
        "_MAIN_CONTACT_PERSON": "MAIN_CONTACT_PERSON",
        "_MAIN_CONTACT_NUMBER": "MAIN_CONTACT_NUMBER",
        "_MAIN_CONTACT_EMAIL": "MAIN_CONTACT_EMAIL",
        "_SECONDARY_CONTACT_PERSON": "SECONDARY_CONTACT_PERSON",
        "_SECONDARY_CONTACT_NUMBER": "SECONDARY_CONTACT_NUMBER",
        "_SECONDARY_CONTACT_EMAIL": "SECONDARY_CONTACT_EMAIL",
        "_DATE_OF_ENTRY": "DATE_OF_ENTRY",
        "_STATUS": "STATUS"
    }
    
    attribute_details = {
        "_VENDOR_NAME": (str,20),
        "_CLASS": (str,10),
        "_COMPANY_NAME": (str,100),
        "_COMPANY_COUNTRY": (str,50),
        "_COMPANY_STATE": (str,50),
        "_COMPANY_CITY": (str,50),
        "_POSTAL_CODE": (str,30),
        "_COMPANY_ADDRESS_A": (str,100),
        "_COMPANY_ADDRESS_B": (str,100),
        "_VENDOR_DESCRIPTION": (str,100),
        "_MAIN_CONTACT_PERSON": (str,100),
        "_MAIN_CONTACT_NUMBER": (str,15),
        "_MAIN_CONTACT_EMAIL": (str,100),
        "_SECONDARY_CONTACT_PERSON": (str,100),
        "_SECONDARY_CONTACT_NUMBER": (str,15),
        "_SECONDARY_CONTACT_EMAIL": (str,100),
        "_DATE_OF_ENTRY": (datetime,0),
        "_STATUS": (int,0)
    }
        
    
    def __getitem__(self, attr):
        return getattr(self, attr)
    #Vendor[attr] returns Vendor.attr
    
    def __setitem__(self,attr,value):
        assert (type(value) == self.attribute_details[attr][0] or value == None)  
        if value is not None:
            assert (len(value) <= self.attribute_details[attr][1]) if self.attribute_details[attr][0] ==str else True
            
            setattr(self,attr,value)
        else:
            if self.attribute_details[attr][0] == str:
                setattr(self,attr,'')
            elif self.attribute_details[attr][0] == int:
                setattr(self,attr,1)
            elif self.attribute_details[attr][0] == datetime:
                setattr(self,attr,datetime.now())
        
    #Vendor[attr] = info (assigning attributes)
    
    def __init__(self,
                 VENDOR_NAME = None, 
                 CLASS = None, 
                 COMPANY_NAME = None, 
                 COMPANY_COUNTRY = None, 
                 COMPANY_STATE = None, 
                 COMPANY_CITY = None, 
                 POSTAL_CODE = None, 
                 COMPANY_ADDRESS_A = None, 
                 COMPANY_ADDRESS_B = None, 
                 VENDOR_DESCRIPTION = None,
                 MAIN_CONTACT_PERSON = None, 
                 MAIN_CONTACT_NUMBER = None, 
                 MAIN_CONTACT_EMAIL = None, 
                 SECONDARY_CONTACT_PERSON = None, 
                 SECONDARY_CONTACT_NUMBER = None, 
                 SECONDARY_CONTACT_EMAIL = None, 
                 DATE_OF_ENTRY = None, 
                 STATUS = None):
        
        self._VENDOR_NAME = ""
        self._CLASS =  ""
        self._COMPANY_NAME = ""
        self._COMPANY_COUNTRY  = ""
        self._COMPANY_STATE = ""
        self._COMPANY_CITY = ""
        self._POSTAL_CODE = ""
        self._COMPANY_ADDRESS_A = ""
        self._COMPANY_ADDRESS_B = ""
        self._VENDOR_DESCRIPTION = ""
        self._MAIN_CONTACT_PERSON = ""
        self._MAIN_CONTACT_NUMBER = ""
        self._MAIN_CONTACT_EMAIL = ""
        self._SECONDARY_CONTACT_PERSON = ""
        self._SECONDARY_CONTACT_NUMBER = ""
        self._SECONDARY_CONTACT_EMAIL = ""
        self._DATE_OF_ENTRY = None
        self._STATUS = 1
        
        if VENDOR_NAME is not None:
            assert len(VENDOR_NAME) <= 20
            self._VENDOR_NAME = VENDOR_NAME
        if CLASS is not None:
            assert len(CLASS) <= 10
            self._CLASS = CLASS
        if COMPANY_NAME is not None:
            assert len(COMPANY_NAME) <= 100
            self._COMPANY_NAME = COMPANY_NAME
        if COMPANY_COUNTRY is not None:
            assert len(COMPANY_COUNTRY) <= 50
            self._COMPANY_COUNTRY = COMPANY_COUNTRY
        if COMPANY_STATE is not None:
            assert len(COMPANY_STATE) <= 50
            self._COMPANY_STATE = COMPANY_STATE
        if COMPANY_CITY is not None:
            assert len(COMPANY_CITY) <= 50
            self._COMPANY_CITY = COMPANY_CITY
        if POSTAL_CODE is not None:
            assert len(POSTAL_CODE) <= 30
            self._POSTAL_CODE = POSTAL_CODE
        if COMPANY_ADDRESS_A is not None:
            assert len(COMPANY_ADDRESS_A) <= 100
            self._COMPANY_ADDRESS_A = COMPANY_ADDRESS_A
        if COMPANY_ADDRESS_B is not None:
            assert len(COMPANY_ADDRESS_B) <= 100
            self._COMPANY_ADDRESS_B = COMPANY_ADDRESS_B
        if VENDOR_DESCRIPTION is not None:
            assert len(VENDOR_DESCRIPTION) <= 100
            self._VENDOR_DESCRIPTION = VENDOR_DESCRIPTION
        if MAIN_CONTACT_PERSON is not None:
            assert len(MAIN_CONTACT_PERSON) <= 100
            self._MAIN_CONTACT_PERSON = MAIN_CONTACT_PERSON
        if MAIN_CONTACT_NUMBER is not None:
            assert len(MAIN_CONTACT_NUMBER) <= 15
            self._MAIN_CONTACT_NUMBER = MAIN_CONTACT_NUMBER
        if MAIN_CONTACT_EMAIL is not None:
            assert len(MAIN_CONTACT_EMAIL) <= 100
            self._MAIN_CONTACT_EMAIL = MAIN_CONTACT_EMAIL
        if SECONDARY_CONTACT_PERSON is not None:
            assert len(SECONDARY_CONTACT_PERSON) <= 100
            self._SECONDARY_CONTACT_PERSON = SECONDARY_CONTACT_PERSON
        if SECONDARY_CONTACT_NUMBER is not None:
            assert len(SECONDARY_CONTACT_NUMBER) <= 15
            self._SECONDARY_CONTACT_NUMBER = SECONDARY_CONTACT_NUMBER
        if SECONDARY_CONTACT_EMAIL is not None:
            assert len(SECONDARY_CONTACT_EMAIL) <= 100
            self._SECONDARY_CONTACT_EMAIL = SECONDARY_CONTACT_EMAIL
        if DATE_OF_ENTRY is not None:
            assert type(DATE_OF_ENTRY) == datetime
            self._DATE_OF_ENTRY = DATE_OF_ENTRY
        else:
            self._DATE_OF_ENTRY = datetime.now()
        if STATUS is not None:
            stat = int(STATUS)
            assert (stat == 1 or stat == 0)
            self._STATUS = STATUS
            
