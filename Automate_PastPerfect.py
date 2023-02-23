#Connor Johnson
#Selenium bot to automate the input of data given by an excel sheet to teh datebase: PastPerfect in the Photos option
#Good chrome extension to find paths in website: SelectorsHub

from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from datetime import datetime

import time


PATH = "/Users/Johnson_code/chromedriver"
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.get("https://mypastperfect.com/Account/Login?ReturnUrl=%2F")     #opens browser to website

username = driver.find_element(By.ID, 'Email')
username.send_keys("******")                        #finds unsername

password = driver.find_element(By.ID, 'Password')
password.send_keys("******")
password.send_keys(Keys.RETURN)                                         #finds password and clicks enter to login

                                                                        #use the try-catch for it to wait until it finds the page


            ###--------------Page Navigation--------------###
try:
    catalog_click = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "catalogs-home-button"))
    )
    catalog_click.click()#clicks the catalog button

    photograph_click = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="object-type-tabs"]/div/ul/li[2]/a'))
    )
    photograph_click.click()#clicks the photograph button

except:
    print('Couldnt not find button')


            ###--------------Excel Import--------------###
book = load_workbook('TestFile.xlsx')                               #excel file from personal computer
sheet = book.active                                                 #reads excel file inside folder
row_num = 2

while row_num <= sheet.max_row:
   for cell in sheet[row_num]:

        excel_Identifier = sheet['A' + str(row_num)].value                                #not using value yet, only tests
        excel_Title = sheet['B' + str(row_num)].value
        excel_Description = sheet['C' + str(row_num)].value
        excel_Object = sheet['D' + str(row_num)].value
        excel_Collection = sheet['E' + str(row_num)].value
        excel_Date = sheet['F' + str(row_num)].value
        excel_Catalog_Date = sheet['G' + str(row_num)].value
        excel_Cataloged_By = sheet['H' + str(row_num)].value
        excel_Attachment = sheet['I' + str(row_num)].value
        excel_Public = sheet['J' + str(row_num)].value

                    ###--------------Insert New Record--------------###

        try:
            add_record_click = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="catalogs-grid"]/div/div[2]/div[1]/div/button[1]'))
            )
            add_record_click.click()#clicks the 'add record' button

        except:
            print("Could not add new record")

        time.sleep(2)
        ObjectID_New_Record = driver.find_element(By.XPATH, '//*[@id="newCatalogRecordDialog-modal"]/div/div/div[2]/div[3]/div[4]/div/input')
        ObjectID_New_Record.send_keys(excel_Identifier)  #identifier                    #finds Object ID Box and inserts from excel sheet

        Search_ObjectName_New_Record = driver.find_element(By.XPATH, '//*[@id="newRecordNameLex"]/div/div/span[2]')
        Search_ObjectName_New_Record.click()                           #finds the search icon to go find Object Name

        time.sleep(2)
        ObjectName_New_Record = driver.find_element(By.XPATH, '/html/body/div[8]/div/div/div[2]/div/div[1]/div/div[1]/input')
        ObjectName_New_Record.send_keys(excel_Object)                        #Inserts object from excel into search

        time.sleep(2)
        Select_New_Record = Select(driver.find_element(By.XPATH, '/html/body/div[8]/div/div/div[2]/div/div[4]/div[2]/div[4]/div[3]/select'))
        Select_New_Record.select_by_visible_text('20')                 #Selects the drop down to make list bigger

        time.sleep(2)
        Action = ActionChains(driver)
        Action_ObjectName_New_Record = driver.find_element(By.XPATH, '/html/body/div[8]/div/div/div[2]/div/div[4]/div[2]/div[3]/table/tbody/tr[13]/td[1]')
        Action.double_click(Action_ObjectName_New_Record).perform()    #finds the 'Photograph' object name

        time.sleep(2)
        Title_New_Record = driver.find_element(By.XPATH, '//*[@id="newCatalogRecordDialog-modal"]/div/div/div[2]/div[3]/div[6]/div/input')
        Title_New_Record.send_keys(excel_Title)                              #Inserts title from excel

        Description_New_Record = driver.find_element(By.XPATH, '//*[@id="newCatalogRecordDialog-modal"]/div/div/div[2]/div[3]/div[7]/div/textarea')
        Description_New_Record.send_keys(excel_Description)                  #Inserts description from excel

        time.sleep(2)
        Add_New_Record = driver.find_element(By.XPATH, '//*[@id="newCatalogRecordDialog-modal"]/div/div/div[2]/div[3]/div[8]/button[1]')
        Add_New_Record.click()                              #clicks the 'add new record' button
        time.sleep(2)

                        ###--------------Going to second webpage after this point--------------###


                        ###--------------Image Management--------------###
        try:
            Enter_Image_Managemnet = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.menu-btn-primary.col-md-8'))
            )
            Enter_Image_Managemnet.click()

            Add_Image_Management = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )
            Add_Image_Management.send_keys("/Users/connorjohnson/Desktop/"+excel_Attachment)

            time.sleep(4)
            Checkbox_Image_Management = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-bind='checked: IsPublic, enable: $parent.editMode()']"))
            )
            Checkbox_Image_Management.click()

            Save_Image_Management = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-bind='click: save, enable: isSaveButtonEnabled()']"))
            )
            Save_Image_Management.click()
        except:
            print("Could not finish Image Management")


        try:

            currentDate = time.strftime("%m/%d/%Y")
            Catalog_Date_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="catalogDate"]/div/input'))
            )
            Catalog_Date_Edit.send_keys(currentDate)

            Date_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="Date"]'))
            )
            Date_Edit.send_keys(excel_Date)

            if excel_Public.lower() == 'yes':
                Public_Access_Edit = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, 'public-access-checkbox'))
                )
                Public_Access_Edit.click()

            Cataloged_By_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='catalogedByDictionary'] span[class='show-dictionary-popup dictionary-icon-enabled']"))
            )
            Cataloged_By_Edit.click()

            Staff_Select_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(3) > td:nth-child(1)"))
            )
            Staff_Select_Edit.click()

            Collection_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='collectionDictionary'] span[class='show-dictionary-popup dictionary-icon-enabled']"))
            )
            Collection_Edit.click()

            if excel_Collection.lower() == 'coffin family papers':
                Coffin_Family_Papers = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(1) > td:nth-child(1)"))
                )
                Coffin_Family_Papers.click()

            if excel_Collection.lower() == 'college archives':
                College_Archives = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(2) > td:nth-child(1)"))
                )
                College_Archives.click()

            if excel_Collection.lower() == 'college photo archives':
                College_Photo_Archives = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(3) > td:nth-child(1)"))
                )
                College_Photo_Archives.click()

            if excel_Collection.lower() == 'honors thesis collection':
                Honors_Thesis_Collection = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(4) > td:nth-child(1)"))
                )
                Honors_Thesis_Collection.click()

            if excel_Collection.lower() == 'lowell thomas papers':
                Lowell_Thomas_Papers = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(5) > td:nth-child(1)"))
                )
                Lowell_Thomas_Papers.click()

            if excel_Collection.lower() == 'lowell thomas papers - radio news show scripts':
                Lowell_Thomas_Papers_Radio = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(6) > td:nth-child(1)"))
                )
                Lowell_Thomas_Papers_Radio.click()

            if excel_Collection.lower() == 'poughkeepsie regatta collection':
                Poughkeepsie_Regatta_Collection = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(7) > td:nth-child(1)"))
                )
                Poughkeepsie_Regatta_Collection.click()

            if excel_Collection.lower() == 'rare book collection':
                Rare_Book_Collection = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(8) > td:nth-child(1)"))
                )
                Rare_Book_Collection.click()

            if excel_Collection.lower() == 'robert hoe music collection':
                Robert_Hoe_Music_Collection = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(9) > td:nth-child(1)"))
                )
                Robert_Hoe_Music_Collection.click()

            if excel_Collection.lower() == 'student newspapers: the record and the circle':
                Student_Newspapers = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "body > div:nth-child(123) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(3) > table:nth-child(2) > tbody:nth-child(2) > tr:nth-child(10) > td:nth-child(1)"))
                )
                Student_Newspapers.click()

            time.sleep(4)
            Final_Save_Edit = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-bind='click: save, enable: isValid']"))
            )
            Final_Save_Edit.click()

            Back_Home = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a[href='/']"))
            )
            Back_Home.click()

            Back_Catalogs = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#catalogs-home-button"))
            )
            Back_Catalogs.click()

        except:
            print("Could not finish Record Edit")

        row_num+=1

print("Task reached end")
time.sleep(60)




































