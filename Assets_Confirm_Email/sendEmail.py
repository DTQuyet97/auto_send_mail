import openpyxl
from openpyxl import Workbook
from findElements import findlistName
from findElements import findName_BarCode
import win32com.client as win32
import time

class Account_Email:
    def __init__(self):
        # Get data from file input.txt
        self.cc_email = "phucnd4@fsoft.com.vn"
        self.path_Device = ""
        lines = []
        self.Project = ""
        with open("input.txt", "r") as obj:
            lines = obj.readlines()
        if(len(lines) > 0):
            for line in lines:
                if "CC" == line[:2]:
                    self.cc_email = line.split('"')[1]
                if "Path_Device" == line[:11]:
                    self.path_Device = line.split('"')[1]
                if "Project" == line[:7]:
                    self.Project = line.split('"')[1]

        # get data from workbook
        self.Asset_Dic_All = {}         # dic contain: all inform of all PIC
        workbook = openpyxl.load_workbook(self.path_Device)
        sheet = workbook.get_sheet_by_name('DeviceInfo')
        self.list_PIC = findlistName(sheet) # all pic 
        for pic in self.list_PIC:
            list_BarCode = findName_BarCode(sheet,pic)
            self.Asset_Dic_All[pic.lower()] = list_BarCode

    def send_email(self, receiver):
        to = receiver + '@fsoft.com.vn'
        # all asset
        list_asset={}
        try:
            list_asset =  self.Asset_Dic_All[receiver.lower()]
        except Exception as e: 
            print(e)
        # get data from sheet
        message = ""       
        for i in range(len(list_asset)):
            message += "{STT}.".format(STT = str(i + 1)) + " "*10
            message += " {Barcode}".format(Barcode = list_asset[i]["Barcode"]) + ": "
            message += " {AssetType}".format(AssetType = list_asset[i]["AssetType"]) + " - "
            message += " {Description}".format(Description = list_asset[i]["Description"])
            message += "<br>"

        # send email
        message1 = "Hi " + receiver.upper() + ",<br>"
        message2 = "You are incharge of the below Asset Code. Please check and reply mail to {cc} to confirm: <br>{msg}".format(cc = self.cc_email, msg=message)
        message3 = "Thanks and Best regards, <br> {Sign} <br><br>".format(Sign = self.Project[:6])
        msg = message1 + "<br>" +  message2 + "<br>" + message3
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'QuyetDT@fsoft.com.vn'
        #mail.CC(self.cc_email)
        mail.Subject = "[{Project}] Check and Confirm project assets".format(Project = self.Project)
        mail.HTMLBody  = msg
        mail.Send()
        print("Send email to {rv}: Done".format(rv=to))

    def send_all(self):
        for pic in self.list_PIC:
            self.send_email(pic)
            time.sleep(3)