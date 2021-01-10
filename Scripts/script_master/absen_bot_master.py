import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ExpCond
from openpyxl import load_workbook

dir_name = os.path.dirname(__file__)

class testAbsent:

    def __init__ (self, rN) :
        self.row_num = rN

    def driverInit(self):
        global dir_name
        self.chrDrv = webdriver.Chrome(os.path.join(dir_name, '..\\..\\Browser_Driver\\chromedriver.exe'))
        self.web_url = str(self.active_sheet['F2'].value)
        self.chrDrv.get(self.web_url)

    def dataInitFromExcel(self):
        global dir_name
        self.datas = load_workbook(filename=os.path.join(dir_name, '..\\..\\Datasources\\Datasources.xlsx'))
        self.active_sheet = self.datas['LIST MAPEL']
        return self.active_sheet

    # TYPE USERNAME
    def doLogin(self, data):
        self.uname_elem = self.chrDrv.find_element_by_xpath('//input[@name=\'username\']')
        self.uname_data = str(data['F3'].value)

        self.uname_elem.send_keys(self.uname_data)
        self.pwd_elem = self.chrDrv.find_element_by_xpath('//input[@name=\'password\']')
        self.pwd_data = str(data['F4'].value)

        self.pwd_elem.send_keys(self.pwd_data)
        self.login_btn = self.chrDrv.find_element_by_xpath('//button[@type=\'submit\']')
        self.login_btn.click()


    def doAbsen(self, data, row_number):
        try :
            # WAIT FOR LOADED
            self.elemCheck = WebDriverWait(self.chrDrv, 15).until(ExpCond.presence_of_element_located((By.XPATH, '//div[@class=\'logo\']')))

            self.url_kls_data = str(data[row_number].value)

            self.chrDrv.get(self.url_kls_data)

            self.confirm_hadir_btn = self.chrDrv.find_element_by_xpath('//a[@id=\'konfirmasikehadiran\']')
            self.confirm_hadir_btn.click()
            
            return True
        
        except:
            return False        

    def Execute(self):
        print('--------  ABSENT BOT by OXXYDDE  --------')
        
        self.xlsx_Data = self.dataInitFromExcel()
        self.mapel_idx = 'A' + str(self.row_num)
        self.mapel_name = str(self.xlsx_Data[self.mapel_idx].value)

        print('Mapel : %s\nLaunching Google Chrome....' % self.mapel_name)
        
        self.driverInit()
        self.doLogin(self.xlsx_Data)

        self.kelas_url_idx = 'B' + str(self.row_num)
        
        if self.doAbsen(self.xlsx_Data, self.kelas_url_idx) :
            print("Mapel : %s\nStatus : Success" % self.mapel_name)
        else :
            print("Mapel : %s\nStatus : Failed" % self.mapel_name)

        input("PRESS ENTER TO EXIT...")
        self.chrDrv.quit()
    

testAbsent(6).Execute()