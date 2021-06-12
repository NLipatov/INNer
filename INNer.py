print("""
Source code can be found on my GitHub: github.com/NLipatov
ver: 1.0
""")



import time, os, xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from xlutils.copy import copy
direction = input('Вставьте путь к файлам:')
workbookname = input('Вставьте название файла:')
os.chdir(direction)
workbook = xlrd.open_workbook(workbookname + '.xls')
sheet = workbook.sheet_by_index(0)
workbookWT = xlrd.open_workbook( 'Sample.xls', formatting_info=True )
sheetWT = workbookWT.sheet_by_index(0)
wb = copy(workbookWT)
wbsheet = wb.get_sheet(0)


RFW = 1
try:
    for i in range (1, sheet.nrows):
        nam = sheet.cell_value(i, 2)
        otch = sheet.cell_value(i,3)
        fam = sheet.cell_value(i,4)
        pnum = sheet.cell_value(i, 7) + sheet.cell_value(i, 8)
        bdate = sheet.cell_value(i, 19).replace('-', '')

        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(r"C:\Users\Admin\PycharmProjects\INN\chromedriver.exe", options=options)

        driver.get('https://service.nalog.ru/static/personal-data.html?svc=inn&from=%2Finn.do')
        driver.find_element_by_id("unichk_0").click()
        driver.find_element_by_id("btnContinue").click()
        time.sleep(0.5)
        for i in range(0, (len(fam))):
            driver.find_element_by_name("fam").send_keys(fam[i])
        for i in range(0, (len(nam))):
            driver.find_element_by_name("nam").send_keys(nam[i])
        for i in range(0, (len(otch))):
            driver.find_element_by_name("otch").send_keys(otch[i])
        for i in range(0, (len(bdate))):
            driver.find_element_by_name("bdate").send_keys(bdate[i])
        for i in range(0, (len(pnum))):
            driver.find_element_by_name("docno").send_keys(pnum[i])

        driver.find_element_by_id("btn_send").click()

        print(f'****Работаю над: {fam} {nam[0]}. {otch[0]}.****')
        count = 0
        while True:
            result = driver.find_element_by_xpath( '//*[@id="result_1"]/div' ).text
            count += 1
            if result == '':
                time.sleep(1)
                if count > 20:
                    print('ИНН не найден. Либо не работает сайт, либо ошибка в данных.\n')
                    result = ''
                    break
            else:
                # print(f'Для клиента {fam} {nam} {otch} успешно найден ИНН.')
                print(f'ИНН: {result[32:69]}')
                wbsheet.write(RFW, 17, (result[32:69]))
                wb.save( 'Sample.xls' )
                RFW += 1
                break
except PermissionError:
    print('Критическая ошибка — Файл нужно закрыть!')
driver.quit()

print('\n\nРабота завершена, файл обработан.')
while True:
    pass
