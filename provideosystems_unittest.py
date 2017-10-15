from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import os

class Provideosystems(unittest.TestCase):
    def setUp(self):
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", os.getcwd()+'\\tmp')
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk", 
            ",application/octet-stream" + 
            ",application/vnd.ms-excel" + 
            ",application/vnd.msexcel" + 
            ",application/x-excel" + 
            ",application/x-msexcel" + 
            ",application/xls" + 
            ",application/vnd.ms-excel" +
            ",application/vnd.ms-excel.addin.macroenabled.12" +
            ",application/vnd.ms-excel.sheet.macroenabled.12" +
            ",application/vnd.ms-excel.template.macroenabled.12" +
            ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
            ",application/vnd.ms-fontobject" +
            ",application/vnd.ms-htmlhelp" +
            ",application/vnd.ms-ims" +
            ",application/vnd.ms-lrm" +
            ",application/vnd.ms-officetheme" +
            ",application/vnd.ms-pki.seccat" +
            ",application/vnd.ms-pki.stl" +
            ",application/vnd.ms-word.document.macroenabled.12" +
            ",application/vnd.ms-word.template.macroenabed.12" +
            ",application/vnd.ms-works" +
            ",application/vnd.ms-wpl" +
            ",application/vnd.ms-xpsdocument" +
            ",application/vnd.openofficeorg.extension" +
            ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
            ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
            ",application/vnd.openxmlformats-officedocument.presentationml.template" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
            ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
            ",application/x-ms-application" +
            ",application/x-ms-wmd" +
            ",application/x-ms-wmz" +
            ",application/x-ms-xbap" +
            ",application/x-msaccess" +
            ",application/x-msbinder" +
            ",application/x-mscardfile" +
            ",application/x-msclip" +
            ",application/x-msdownload" +
            ",application/x-msmediaview" +
            ",application/x-msmetafile" +
            ",application/x-mspublisher" +
            ",application/x-msschedule" +
            ",application/x-msterminal" +
            ",application/x-mswrite" +
            ",application/xml" +
            ",application/xml-dtd" +
            ",application/xop+xml" +
            ",application/xslt+xml" +
            ",application/xspf+xml" +
            ",application/xv+xml" +
            ",application/excel")

        self.driver = webdriver.Firefox(ffprofile)
        self.driver.implicitly_wait(30)
        self.base_url = "http://www.provis.ru/"
        self.verificationErrors = []
        self.accept_next_alert = True

    
    def test_provideosystems(self):
        driver = self.driver
        driver.get(self.base_url + "/partners/dealer/")
        driver.find_element_by_id("enterProfile").click()
        driver.find_element_by_css_selector("#modal_enterProfile > div.cntlogin > form[name=\"auth\"] > div > input[name=\"login\"]").clear()
        driver.find_element_by_css_selector("#modal_enterProfile > div.cntlogin > form[name=\"auth\"] > div > input[name=\"login\"]").send_keys("egn@av-prom.ru")
        driver.find_element_by_css_selector("#modal_enterProfile > div.cntlogin > form[name=\"auth\"] > div > input[name=\"pass\"]").clear()
        driver.find_element_by_css_selector("#modal_enterProfile > div.cntlogin > form[name=\"auth\"] > div > input[name=\"pass\"]").send_keys("Dexim0oghu")
        driver.find_element_by_css_selector("#modal_enterProfile > div.cntlogin > form[name=\"auth\"] > div.controls > #partners_auth_btn").click()
        driver.find_element_by_css_selector("div.price-list_item.clearfix > div.f_l > div.clearfix > a.dealer_red_btn").click()
        time.sleep(20)
                 


    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
