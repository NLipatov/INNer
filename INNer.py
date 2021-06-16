print("""
Source code can be found on my GitHub: github.com/NLipatov
This software is released under the MIT license

ver: 1.0
""")


import time, os, xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from xlutils.copy import copy

DriverPath = r'C:\INNer\chromedriver.exe'

direction = input('Вставьте путь к файлам:')
workbookname = input('Вставьте название файла:')
os.chdir(direction)

TotalWorkingTimeStart = time.time()
Chk1 = 0


DriverPathExist = os.path.isfile(DriverPath)

if DriverPathExist == True:
    pass
else:
    while True:
        print('Не обнаружен chromedriver.exe в папке с INNer\'ом')
        pass





workbook = xlrd.open_workbook(workbookname + '.xls')
sheet = workbook.sheet_by_index(0)
workbookWT = xlrd.open_workbook( (workbookname + '.xls'), formatting_info=True )
sheetWT = workbookWT.sheet_by_index(0)
wb = copy(workbookWT)
wbsheet = wb.get_sheet(0)
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_experimental_option( 'excludeSwitches', ['enable-logging'] )
driver = webdriver.Chrome(f'{DriverPath}', options=options)

TotalClientsNumber = sheet.nrows - 1
TotalClientsNumberNowWorking = 0
RFW = 1
print(f'Ожидаемое время обработки {(3.7*sheet.nrows)/60} минут(-ы)\n\n')
try:
    for i in range (1, sheet.nrows):
        print( f'Клиентов обработано: {TotalClientsNumberNowWorking}/{TotalClientsNumber}\n' )
        ST = time.time()
        RFW = i
        nam = sheet.cell_value(i, 2)
        otch = sheet.cell_value(i,3)
        fam = sheet.cell_value(i,4)
        pnum = sheet.cell_value(i, 7) + sheet.cell_value(i, 8)
        bdate = sheet.cell_value(i, 19).replace('-', '')
        print( f'****Работаю над: {fam} {nam[0]}. {otch[0]}.****' )
        if RFW != i:
            Chk1 +=1 # Проставляю Чекван + 1, так как по какой-то причине RFW не равен i. \
            # Программа не будет работать дальше
            while True:
                print('Критическая ошибка: RFW != i!')
        if fam == sheet.cell_value(i,4):
            pass
        else:
            print('\nЗавершение работы - Критическая ошибка - Фамилия из сайта и из файла не совпали')
            break
        if Chk1 < 1:
            for i in range(0,10):
                try:
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
                    break
                except:
                    print('Ошибка при загрузке данных. Проверьте работу сайта ФНС.')

            count = 0
            while True:
                try:
                    result = driver.find_element_by_xpath( '//*[@id="result_1"]/div' ).text
                    count += 1
                    if result == '':
                        time.sleep(1)
                        if count > 5:
                            print('ИНН не найден. Вероятно, ошибка в данных.\n')
                            TotalClientsNumberNowWorking += 1
                            wbsheet.write(RFW, 17, '')
                            break
                    else:
                        wbsheet.write(RFW, 17, (result[32:69]))
                        ET = time.time()
                        print(f'ИНН:{result[32:69]}')
                        print( 'Результат получен и записан за %s секунды' % (ET - ST) )
                        wb.save('INNERED — ' + workbookname + '.xls')
                        TotalClientsNumberNowWorking += 1
                        break
                except:
                    if count == 20:
                        print('Не нашёл элемент, содержащий ИНН.')
        else:
            print('Не пройдена проверка Chk1.')
except PermissionError:
    print('Критическая ошибка — Файл нужно закрыть!')
driver.quit()

if Chk1 == 0:
    print('\n\nРабота завершена, файл обработан.')
else:
    print('Работа завершена с ошибками!')
TotalWorkingTimeEnd = time.time()
print('Программа отработала за %s минут(-ы)' % ((TotalWorkingTimeEnd - TotalWorkingTimeStart) / 60))
print('Темп равен %s секунд на клиента' % (((TotalWorkingTimeEnd-TotalWorkingTimeStart)/sheet.nrows)))
while True:
    pass
