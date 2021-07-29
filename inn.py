import time, os, xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from xlutils.copy import copy


def Tinkoff_search_for_inn(fam, nam, otch, bdate, pnum, passportdate):
    DriverPath = r'C:\INNer\chromedriver.exe'
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_experimental_option( 'excludeSwitches', ['enable-logging'] )
    driver = webdriver.Chrome(f'{DriverPath}', options=options)
    try:
        driver.get('https://www.tinkoff.ru/inn/')
        driver.find_element_by_name( "fullName" ).click()
        time.sleep(0.5)
        for i in range(0, (len(fam))):
            driver.find_element_by_name("last").send_keys(fam[i])
        for i in range(0, (len(nam))):
            driver.find_element_by_name("first").send_keys(nam[i])
        for i in range(0, (len(otch))):
            driver.find_element_by_name("middle").send_keys(otch[i])
        for i in range(0, (len(bdate))):
            driver.find_element_by_name("birthday").send_keys(bdate[i])
        for i in range(0, (len(pnum))):
            driver.find_element_by_name("passport_number").send_keys(pnum[i])
        for i in range(0, (len(passportdate))):
            driver.find_element_by_name("passport_date").send_keys(passportdate[i])
        for i in range( 0, 10 ):
            try:
                driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[1]/div[2]/div[1]/div[2]/div/div[2]/section/div/div/div[1]/div/form/div[4]/button/h2").click()
                time.sleep(1)
                result = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/div[2]/div[1]/div[2]/div/div[2]/section/div/div/div[1]/div[2]/div/div' ).text
                if len(result) == 12:
                    return result
                else:
                    continue
                time.sleep(1)
            except:
                errortext = driver.find_element_by_xpath('/html/body/div[1]/div/div/span/div/div/div/div/div/div/div/div[2]/div/span[1]' ).text
                print(f'Ошибка: {errortext}')
                return '-'
                break
    except:
        return '-'
    print('Работа завершена')
    driver.quit()

print(Tinkoff_search_for_inn('Агаева', 'Наиля', 'Джавидовна', '10-11-1996', '4518071547', '28-03-2017'))
