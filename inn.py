import time, os, xlrd, xlwt, re, webbrowser, time_convert, tkinter as tk
from tkinter import messagebox as mb
from xlwt import easyxf
from threading import Thread
from selenium import webdriver
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW
from xlutils.copy import copy

direction = None
workbookname = None
status = ''


def DirWor(Dir, Wor):
    global direction
    global workbookname
    direction = Dir
    workbookname = Wor

innermessage = 'Работа над файлом'
message = ''
refreshedINN = 0

def Add_to_Message(arg):
    global message
    message += str(arg)

def Get_message():
    return innermessage+message


def TT():
    os.chdir( direction )
    def Draw():
        global text, status
        status = 'Файл в работе'
        frame=tk.Frame(root,relief='solid',bd=1, bg='grey')
        root.geometry('700x90+500+600')
        root.iconbitmap( r"C:\INNer\INNERICC.ico" )
        root.iconbitmap( r"C:\INNer\INNERICC.ico" )
        root.title( 'Файл в работе' )
        try:
            window.iconbitmap(r"C:\INNer\INNERICC.ico")
        except:
            pass
        frame.pack()
        text=tk.Label(frame,text='HELLO')
        text.pack(side='top')

    def Refresher():
        global text
        text.configure(text=Get_message())
        root.after(1000, Refresher)
        if status == 'Готово':
            root.title(status)
            root.lift()
            root.attributes( '-topmost', True )
            root.after_idle( root.attributes, '-topmost', False )
        else:
            root.title(status)


    root=tk.Tk()
    Draw()
    Refresher()
    root.mainloop()


def Main():
    global innermessage, message, status
    DriverPath = r'C:\INNer\chromedriver.exe'

    os.chdir(direction)

    TotalWorkingTimeStart = time.time()
    Chk1 = 0

    DriverPathExist = os.path.isfile(DriverPath)

    if DriverPathExist == True:
        pass
    else:
        while True:
            status = 'Ошибка'
            innermessage = 'Ошибка'
            message = ('Не обнаружен chromedriver.exe в папке с INNer\'ом')
            pass
    workbook = xlrd.open_workbook(workbookname)
    sheet = workbook.sheet_by_index(0)
    workbookWT = xlrd.open_workbook( (workbookname), formatting_info=True )
    sheetWT = workbookWT.sheet_by_index(0)
    wb = copy(workbookWT)
    wbsheet = wb.get_sheet(0)
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_experimental_option( 'excludeSwitches', ['enable-logging'] )

    try:
        driver = webdriver.Chrome(f'{DriverPath}', options=options)
    except SessionNotCreatedException as session_not_created_ex:
        error_text = str(session_not_created_ex)[29:]
        status = 'Ошибка'
        innermessage = 'Описание возникшей ошибки:'
        message = ( f'\n{error_text}' )
        if mb.askyesno(title='Требуется обновление chromedriver.exe',\
                       message='Перейти на страницу загрузи chromedriver?'):
            webbrowser.open('https://chromedriver.chromium.org/downloads')
            mb.showinfo(title='Подсказка', \
                        message='После загрузки новой версии \'chromedriver.exe\',\
                    скопируйте её в C:\INNer с заменой файлов')
        while True:
            pass




    TotalClientsNumber = sheet.nrows - 1
    TotalClientsNumberNowWorking = 0
    RFW = 1
    infoApproxTime = (f'Ожидаемое время работы {(3.7*sheet.nrows-TotalClientsNumber)/60} минут(-ы)\n\n\n')
    message = str(infoApproxTime)
    try:
        for i in range (1, sheet.nrows):
            message = str( f'\n\nКлиентов обработано: {TotalClientsNumberNowWorking}/{TotalClientsNumber}\n' )
            ST = time.time()
            RFW = i
            translit_name = sheet.cell_value( i, 5 )
            re_res = re.findall( r'^\w+', f'{translit_name}' )
            if re_res[0].lower() in ['mister', 'mistress', 'mizz', 'miss','mr', 'mrs', 'miss', 'ms',  'dr']:
                corrected_translit_name = translit_name.replace( f'{re_res[0]} ', '' )
                wbsheet.write( i, 5, corrected_translit_name, easyxf( 'font: name Calibri, height 220;' ) )
                wb.save( 'INNERED — ' + workbookname )
            nam = sheet.cell_value(i, 2)
            if nam == '':
                message += str('\n\n\nBreak. File ends here — blank cell at clients name\n\n\n')
                break
            otch = sheet.cell_value(i,3)
            lotch = str(sheet.cell_value(i,3)).lower()
            fam = sheet.cell_value(i,4)
            pnum = sheet.cell_value(i, 7) + sheet.cell_value(i, 8)
            bdate = sheet.cell_value(i, 19).replace('-', '')
            if lotch != '' or lotch != 'отсутствует' or lotch != 'нет' or lotch != '-':
                middle_name = 0
            else:
                middle_name = 1
            if middle_name == 1:
                infoNowWorkingOnClient = str( f'Приступил к: {fam} {nam[0]}. {otch[0]}.' )
            else:
                infoNowWorkingOnClient = str( f'Приступил к: {fam} {nam[0]}.' )
            message = message + str(infoNowWorkingOnClient)
            refreshedINN = 0
            if RFW != i:
                Chk1 +=1
                while True:
                    status = 'Ошибка'
                    innermessage = 'Ошибка'
                    message = ('Критическая ошибка: RFW != i!')
                    pass
            if fam == sheet.cell_value(i,4):
                pass
            else:
                while True:
                    status = 'Ошибка'
                    innermessage = 'Ошибка'
                    message = ('\nЗавершение работы - Критическая ошибка - Фамилия из сайта и из файла не совпали')
                    pass
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
                            if lotch == '' or lotch == 'отсутствует' or lotch == 'нет' or lotch == '-':
                                driver.find_element_by_xpath( '//*[@id="unichk_0"]' ).click()
                            else:
                                for i in range(0, (len(otch))):
                                    driver.find_element_by_name("otch").send_keys(otch[i])
                            for i in range(0, (len(bdate))):
                                driver.find_element_by_name("bdate").send_keys(bdate[i])
                            for i in range(0, (len(pnum))):
                                driver.find_element_by_name("docno").send_keys(pnum[i])
                            driver.find_element_by_id("btn_send").click()
                            break
                        except:
                            status = 'Ошибка'
                            innermessage = 'Ошибка возникла до получения результата запроса, после отправки запроса.'
                    for i in range(0, 10):
                        try:
                            result = driver.find_element_by_xpath( '//*[@id="result_1"]/div' ).text
                            if result == '':
                                time.sleep(0.5)
                                if i == 9:
                                    res_msg = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div/form/div[1]/div[1]/div/div[2]/p[1]').text
                                    status = 'Ошибка'
                                    innermessage = f'{res_msg}'
                                    TotalClientsNumberNowWorking += 1
                                    wbsheet.write(RFW, 17, '-')
                                    break
                            else:
                                wbsheet.write(RFW, 17, (result[32:69]))
                                ET = time.time()
                                print(f'\nИНН:{result[32:69]}')
                                wb.save('INNERED — ' + workbookname)
                                status = 'Файл в работе'
                                innermessage = f'Записал — ИНН клиента {fam}: {result[32:69]}' + \
                                               '\nРезультат получен и записан за %.1f секунды' % ((ET - ST))
                                TotalClientsNumberNowWorking += 1
                                break
                        except:
                            if i == 9:
                                status = 'Файл в работе'
                                innermessage = 'Не нашёл элемент, содержащий ИНН'
                                message = (f'i - {i}, fam - {fam}')
                                wbsheet.write( RFW, 17, '-' )
                                wb.save( 'INNERED — ' + workbookname )
                                TotalClientsNumberNowWorking += 1
            else:
                status = 'Критическая ошибка'
                innermessage = 'Работа программы остановлена'
                message = ('Не пройдена проверка Chk1.')
    except PermissionError:
        status = 'Критическая ошибка'
        innermessage = 'Работа программы остановлена'
        message = ('Критическая ошибка — Файл нужно закрыть!')
    driver.quit()

    if Chk1 == 0:
        status = 'Готово!'
        innermessage = '\nРабота завершена остановлена'
        message = ('\nФайл обработан')
    else:
        status = 'Результат работы не подлежит использованию!'
        innermessage = 'Внимание:'
        message = ('Работа завершена с ошибками!')
        while True:
            pass
    TotalWorkingTimeEnd = time.time()
    innermessage = 'Готово'
    EndingA = ('\nПрограмма отработала за %.1f минут(-ы)' % (((TotalWorkingTimeEnd - TotalWorkingTimeStart) / 60)))
    EndingB = ('\nТемп равен %.1f секунд(-ы) на одного клиента' % (((TotalWorkingTimeEnd-TotalWorkingTimeStart)/sheet.nrows)))
    message = str(EndingA)
    message += str(EndingB)
    workbook.release_resources()
    status = 'Работа завершена'






def Start():
    MAin_thread = Thread(target = Main)
    MAin_thread.start()
    TT_thread = Thread(target = TT)
    TT_thread.start()
    # Destroyer_thread = Thread(target=Destroyer())
    # Destroyer_thread.start()
    # MAin_thread.join()
    # TT_thread.join()










