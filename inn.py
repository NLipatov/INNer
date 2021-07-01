# print("""
# Source code can be found on my GitHub: github.com/NLipatov
# This software is released under the MIT license
#
# ver: 1.1.1
# """)


import time, os, xlrd, xlwt, tkinter as tk
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW
from xlutils.copy import copy

direction = None
workbookname = None




def DirWor(Dir, Wor):
    global direction
    global workbookname
    direction = Dir
    workbookname = Wor
    print(direction, workbookname)


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
        global text
        frame=tk.Frame(root,relief='solid',bd=1)
        root.geometry('250x90+500+600')
        root.title( 'Файл в работе' )
        try:
            window.iconbitmap( "INNer.ico" )
        except:
            pass
        frame.pack()
        text=tk.Label(frame,text='HELLO')
        text.pack(side='top')

    def Refresher():
        global text
        text.configure(text=Get_message())
        root.after(1000, Refresher)

    root=tk.Tk()
    Draw()
    Refresher()
    root.mainloop()

#
# def UI_ON():
#     global TextFrame
#     window = tk.Tk()
#     window.title('Файл в работе')
#     window.geometry('500x200+700+300')
#     masterFrame = tk.Frame(window)
#     TextFrame = tk.Label(master=masterFrame, text=f'{message}')
#     masterFrame.pack()
#     TextFrame.pack()
#     TextFrame.configure( text=message )
#     window.after( 60)
#     window.mainloop()





def Main():
    global message
    DriverPath = r'C:\INNer\chromedriver.exe'


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


    driver = webdriver.Chrome(f'{DriverPath}', options=options)


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
            nam = sheet.cell_value(i, 2)
            if nam == '':
                message += str('\n\n\nBreak. File ends here — blank cell at clients name\n\n\n')
                break
            otch = sheet.cell_value(i,3)
            fam = sheet.cell_value(i,4)
            pnum = sheet.cell_value(i, 7) + sheet.cell_value(i, 8)
            bdate = sheet.cell_value(i, 19).replace('-', '')
            if otch != '':
                infoNowWorkingOnClient = str( f'****Работаю над: {fam} {nam[0]}. {otch[0]}.****' )
            else:
                infoNowWorkingOnClient = str( f'****Работаю над: {fam} {nam[0]}.****' )
            message = message + str(infoNowWorkingOnClient)
            refreshedINN = 0
            if RFW != i:
                Chk1 +=1 # Проставляю Чекван + 1, так как по какой-то причине RFW не равен i. \
                # Программа не будет работать дальше
                while True:
                    print('Критическая ошибка: RFW != i!')
                    pass
            if fam == sheet.cell_value(i,4):
                pass
            else:
                while True:
                    print('\nЗавершение работы - Критическая ошибка - Фамилия из сайта и из файла не совпали')
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
                            if otch == '' or otch == 'Отсутствует' or otch == 'отсутствует'\
                                    or otch == 'Нет' or otch == 'нет' or otch == '-':
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
                            print('\nОшибка при загрузке данных. Проверьте работу сайта ФНС.')
                    for i in range(0, 10):
                        try:
                            result = driver.find_element_by_xpath( '//*[@id="result_1"]/div' ).text
                            if result == '':
                                time.sleep(0.5)
                                if i == 9:
                                    print('\nИНН не найден. Вероятно, ошибка в данных.\n')
                                    TotalClientsNumberNowWorking += 1
                                    wbsheet.write(RFW, 17, '-')
                                    break
                            else:
                                wbsheet.write(RFW, 17, (result[32:69]))
                                ET = time.time()
                                print(f'\nИНН:{result[32:69]}')
                                wb.save('INNERED — ' + workbookname)
                                print( '\nРезультат получен и записан за %s секунды' % (ET - ST) )
                                TotalClientsNumberNowWorking += 1
                                break
                        except:
                            if i == 9:
                                print(f'Не нашёл элемент, содержащий ИНН.i - {i}, fam - {fam}')
                                wbsheet.write( RFW, 17, '-' )
                                TotalClientsNumberNowWorking += 1
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
    innermessage = 'Done'
    EndingA = ('\nПрограмма отработала за %s минут(-ы)' % ((TotalWorkingTimeEnd - TotalWorkingTimeStart) / 60))
    EndingB = ('\nТемп равен %s секунд(-ы) на одного клиента' % (((TotalWorkingTimeEnd-TotalWorkingTimeStart)/sheet.nrows)))
    message = str(EndingA)
    message += str(EndingB)


def Start():
    Thread(target = Main).start()
    Thread(target = TT).start()
