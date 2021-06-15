print("""
Source code can be found on my GitHub: github.com/NLipatov
This software is released under the MIT license

ver: 1.0
""")


import time, os, xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from xlutils.copy import copy

critical_counter = 0
INNerSettingsFilePath = os.path.dirname(os.path.abspath(__file__)).replace('\\', '\\\\') + '\\\\chromedriver.exe'
print(INNerSettingsFilePath)
direction = input('Вставьте путь к файлам:')
workbookname = input('Вставьте название файла:')
os.chdir(direction)
workbook = xlrd.open_workbook(workbookname + '.xls')
sheet = workbook.sheet_by_index(0)
workbookWT = xlrd.open_workbook( (workbookname + '.xls'), formatting_info=True )
sheetWT = workbookWT.sheet_by_index(0)
wb = copy(workbookWT)
wbsheet = wb.get_sheet(0)



RFW = 1
try:
    for i in range (1, sheet.nrows):
        RFW = i
        nam = sheet.cell_value(i, 2)
        otch = sheet.cell_value(i,3)
        fam = sheet.cell_value(i,4)
        pnum = sheet.cell_value(i, 7) + sheet.cell_value(i, 8)
        bdate = sheet.cell_value(i, 19).replace('-', '')
        print(i, RFW)
        if fam == sheet.cell_value(i,4):
            print(f'PASSED — {fam} / {sheet.cell_value(i,4)}')
        if fam != sheet.cell_value(i,4):
            print(f'NOT PASSED — {fam} / {sheet.cell_value(i,4)}')
            critical_counter += 1

        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_experimental_option( 'excludeSwitches', ['enable-logging'] )
        driver = webdriver.Chrome(f'{INNerSettingsFilePath}', options=options)

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
                    wbsheet.write(RFW, 17, '')
                    break
            else:
                wbsheet.write(RFW, 17, (result[32:69]))
                print(f'ИНН:{result[32:69]}')
                wb.save('INNERED — ' + workbookname + '.xls')
                break

except PermissionError:
    print('Критическая ошибка — Файл нужно закрыть!')
driver.quit()

print('\n\nРабота завершена, файл обработан.')
print(f'Ошибок присваивания: {critical_counter}. \n/'
      f'При наличии хотя бы одной ошибки присваивания, результатом выдачи программы пользоваться нельзя!')
while True:
    pass
