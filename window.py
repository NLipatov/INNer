import tkinter as tk
from tkinter import PhotoImage
import os, time, inn, no_payment_sms_gui, AWBFGUI
import urllib.request
from tkinter.filedialog import askopenfilename, asksaveasfilename


def Response_Code_Meaning():
    response = ''
    # response = urllib.request.urlopen( "https://service.nalog.ru/static/personal-data.html?svc=inn&from=%2Finn.do" ).getcode()
    if response == 200:
        FNSStatus = ' Доступен'
    elif response == 500:
        FNSStatus = ' Недоступен'
    else:
        FNSStatus = ' Необходимо проверить статус сайта ФНС вручную'
    return FNSStatus



helloraised = 0
def write_hello():
    global helloraised
    global w
    if helloraised == 0:
        helloraised = 1
        w = tk.Label(window, text = "MIT License\nCopyright (c) 2021 N.Lipatov", width=21)
        w.grid()
    else:
        helloraised == 1
        helloraised = 0
        w.grid_forget()


def addtotext(arg):
    text.insert(tk.END, arg)

def FileSelect():
    filepath = askopenfilename( filetypes=[("Книга Excel", "*.xls")], title='Выберите файл' )
    if filepath == '':
        print('No Filepath')
        pass
    else:
        print(f' Filepath = {filepath}')
        workbook = os.path.basename( filepath )
        direction = filepath.replace( f'/{workbook}', '' )
        inn.DirWor(direction, workbook)
        inn.Start()


def innlog():
    inn.ToLog()


window = tk.Tk()


window.resizable(width=False, height=False)
window.resizable(0, 0)
window.title('INNer v1.2.0.1')
try:
    window.iconbitmap(r"C:\INNer\INNERICC.ico")
except:
    pass


window.rowconfigure([0],minsize=300, weight=1)
window.columnconfigure([0, 1],minsize=10, weight=1)

clear_icon = PhotoImage(file=r'C:\INNer\Clear.png')
clear_icon_su = clear_icon.subsample(40,67)

buttonFrame = tk.Frame(window, bd= 7, bg='grey')
AuthorsFrame = tk.Frame(window, bd= 7, bg='grey')
buttonOpen = tk.Button(buttonFrame, text='Прогнать файл', command=FileSelect, width=20)
ButtonSMS_nodata_payed = tk.Button(buttonFrame, text='SMS: нет данных, оплачено', command=no_payment_sms_gui.GUI, width=20)
ButtonAWBF = tk.Button(buttonFrame, text='Собрать а/н', command=AWBFGUI.Main, width=24)
ButtonAWBF_clear = tk.Button(buttonFrame, text='Очистить буфер', image=clear_icon_su, command=AWBFGUI.Clear, width=13)
buttonAuthors = tk.Button(AuthorsFrame, text='Автор', command=write_hello, width=20)


buttonOpen.grid(row=0, column=0, ipadx=23, sticky='w')
ButtonSMS_nodata_payed.grid(row=1, column=0, ipadx=23, sticky='w')
ButtonAWBF.grid(row=2, column=0, sticky='w')
ButtonAWBF_clear.grid(row=2, column=0, sticky='e')
buttonAuthors.grid(row=3, column=0, ipadx=23, sticky='s')

buttonFrame.grid(row=0, column=0, sticky='nsew')
AuthorsFrame.grid(row=1, column=0, sticky='s')

logFrame = tk.Frame(window)
logLabel = tk.Label(master=logFrame, text='Статус-бар:', bg='#C0C0C0')
logContainer = tk.Label(master=logFrame,text=f'Сервис ФНС —{Response_Code_Meaning()}', width=100)
logLabel.grid(row=0, column= 1, sticky="new")
logFrame.grid(row=0, column=1, sticky="new")
logContainer.grid(row=1, column=1, sticky="w")



window.mainloop()
