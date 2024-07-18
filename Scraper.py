from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_recaptcha_solver import RecaptchaSolver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import load_workbook
from datetime import datetime
import os
import re

def checkChanges():

    test_ua = 'Mozilla/5.0 (Windows NT 4.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36'

    options = Options()
    options.add_argument("start-maximized")
    options.add_argument("disable-infobars")

    options.add_argument("--window-size=1920,1080")

    options.add_argument(f'--user-agent={test_ua}')

    options.add_argument('--no-sandbox')
    options.add_argument("--disable-extensions")

    workbook = load_workbook('ES Registry Database - Updated.xlsx')
    worksheet = workbook['Reuse']

    print(worksheet.max_row)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    #solver = RecaptchaSolver(driver=driver)

    driver.get('https://excesssoilnotices.rpra.ca/s/?language=en_US')   
    wait = WebDriverWait(driver, 5)
    name = input("Click enter when done the captcha")
    counter = -1
    while counter < 4:
        counter += 1
        time.sleep(3)
        elements = driver.find_elements(By.XPATH, '//*[@data-aura-class="cRegistryPublicPortalFilingAction"]')
        municipalities = driver.find_elements(By.XPATH, '//span[contains(@class, "uiOutputText") and substring(text(), string-length(text()) - 2) = " of"]')
        notice = driver.find_elements(By.XPATH, '//span[contains(@class, "uiOutputText") and (text()="RS" or text()="PA")]')
        if notice[counter].text == "RS":
            worksheet = workbook['Reuse']
        elif notice[counter].text == "PA":
            worksheet = workbook['Project Area']
        else:
            continue
        municipality = municipalities[counter]
 
        if worksheet == workbook['Reuse']:
             row = worksheet.max_row + 1

             worksheet['F' + str(row)] = municipality.text.split(",")[0]

             elements[counter].click()

             time.sleep(3)

             projectName = driver.find_element(By.XPATH, '//*[starts-with(@data-aura-rendered-by, "225")]')
             worksheet['A' + str(row)] = projectName.text

             companyName = driver.find_element(By.XPATH, '//*[starts-with(@data-aura-rendered-by, "40:") and @class = "slds-cell-wrap"]')
             worksheet['B' + str(row)] = companyName.text

             location = driver.find_element(By.XPATH, '//span[contains(@data-aura-rendered-by, "24:") and @class="uiOutputText" and contains(., ",")]')
             worksheet['C' + str(row)] = location.text

             table = driver.find_element(By.XPATH, '//span[contains(@data-aura-rendered-by, "47") and contains (., "-")]')
             worksheet['D' + str(row)] = table.text

             soil = driver.find_element(By.XPATH, '//*[contains(@data-aura-rendered-by, "729")]')
             worksheet['E' + str(row)] = soil.text

             dateAdded = driver.find_element(By.XPATH, '//lightning-formatted-date-time[starts-with(@data-aura-rendered-by, "53")]')
             cleanDate = dateAdded.text
             formattedDate = datetime.strptime(cleanDate, '%d-%b-%Y') # Convert datetime object to date
             worksheet['G' + str(row)] = formattedDate.strftime('%Y-%m-%d') # Format as string without time

             contactName = driver.find_element(By.XPATH, '//span[contains(@data-aura-rendered-by, "24:") and @class="uiOutputText" and not(contains(., ","))]')
             contactMail = driver.find_element(By.XPATH, '//span[contains(@data-aura-rendered-by, "87")]')
             worksheet['H' + str(row)] = contactName.text + " " + contactMail.text

             #url = driver.current_url
             worksheet['I' + str(row)] = driver.current_url

             coordinates = driver.find_element(By.XPATH, '//span[contains(@data-aura-rendered-by, "46") and @class="uiOutputText" and contains(., ",") and contains(translate(., "0123456789,-", ""), "")]')
             worksheet['J' + str(row)] = coordinates.text

        elif (worksheet == workbook['Project Area']):
             row = worksheet.max_row + 1

             #worksheet['F' + str(row)] = municipality.text.split(",")[0]

             elements[counter].click()

             time.sleep(3)
             
             siteNames = driver.find_elements(By.XPATH, "//div[@class='slds-cell-wrap' and text()='Site Name']/following::td[1]")
             for item in siteNames:
                projectName = driver.find_element(By.XPATH, '//*[contains(@data-aura-rendered-by, "228")]')
                worksheet['A' + str(row)] = projectName.text

                companyName = driver.find_element(By.XPATH, '//*[starts-with(@data-aura-rendered-by, "40:") and @class = "slds-cell-wrap"]')
                worksheet['B' + str(row)] = companyName.text
                try:
                    driver.find_element(By.XPATH, '//*[contains(@data-aura-rendered-by, "316") and @class = "slds-truncate"]')
                    location = driver.find_element(By.XPATH, '//lightning-formatted-rich-text[contains(@data-aura-rendered-by, "323")]/span[@part="formatted-rich-text"]')
                except:
                    location = driver.find_elements(By.XPATH, '//*[self::lightning-formatted-rich-text/span[@part="formatted-rich-text" and normalize-space(.)] or self::span[@part="formatted-rich-text" and normalize-space(.)]]')
                    location = location[5]
                worksheet['C' + str(row)] = location.text

                city = driver.find_element(By.XPATH, "//div[text()='Municipality']/following::td[1]")
                worksheet['D' + str(row)] = city.text.split(",")[0]

                site = item.text
                worksheet['E' + str(row)] = site

                soil = driver.find_element(By.XPATH, "//div[text()='Total Estimated Amount of Excess Soil (m3)']/following::td[1]")
                worksheet['F' + str(row)] = soil.text

                dateAdded = driver.find_element(By.XPATH, '//lightning-formatted-date-time[starts-with(@data-aura-rendered-by, "53")]')
                cleanDate = dateAdded.text #Might have to change depending on whetehr it's date submitted or modified, right now it's submitted
                formattedDate = datetime.strptime(cleanDate, '%d-%b-%Y') # Convert datetime object to date
                worksheet['H' + str(row)] = formattedDate.strftime('%Y-%m-%d') # Format as string without time     

                operator = driver.find_element(By.XPATH, "//span[@class='uiOutputEmail' and contains(@data-aura-rendered-by, '197')]/a")
                worksheet['F' + str(row)] = operator.text

        driver.execute_script("window.history.go(-1)")

    workbook.save(filename="test.xlsx")

checkChanges()