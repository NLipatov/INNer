import tkinter as tk
import os, time, inn
import urllib.request
from tkinter.filedialog import askopenfilename, asksaveasfilename


def Response_Code_Meaning():
    response = urllib.request.urlopen( "https://service.nalog.ru/inn.do" ).getcode()
    if response == 200:
        FNSStatus = 'Доступен'
    elif response == 500:
        FNSStatus = 'Недоступен'
    else:
        FNSStatus = 'Необходимо проверить статус сайта ФНС вручную'
    return FNSStatus


def LogRead(label):
    def read():
        with open(r'C:\Users\Admin\Desktop\INNer\Log.txt') as Log:
            lines = Log.read()
            logContainer.configure(text=(str(lines)))
        label.after(10, read)
    read()


helloraised = 0
def write_hello():
    global helloraised
    global w
    if helloraised == 0:
        helloraised = 1
        w = tk.Label(window, text = "This software is released under the MIT Licence\nAuthor: Lipatov Nikita\nIcon: upper left by Simphiwe Mangole from the Noun Project")
        w.grid()
    else:
        helloraised == 1
        helloraised = 0
        w.grid_forget()


def addtotext(arg):
    text.insert(tk.END, arg)

def FileSelect():
    filepath = askopenfilename( filetypes=[("Книга Excel", "*.xls")] )
    if filepath == '':
        pass
    else:
        print(f' Filepath = {filepath}')
        workbook = os.path.basename( filepath )
        direction = filepath.replace( f'/{workbook}', '' )
        inn.DirWor(direction, workbook)
        inn.Start()

def innlog():
    inn.ToLog()

# CheckIfLogExists()



window = tk.Tk()


window.resizable(width=False, height=False)
window.resizable(0, 0)
window.title('INNer')
try:
    window.iconbitmap( "INNer.ico" )
except:
    pass

disclaimerFrame = tk.Frame()
disclaimer = tk.Label(master=disclaimerFrame, text="""Source code can be found on my GitHub: github.com/NLipatov
This software is released under the MIT license.

Lead developer: Lipatov Nikita
Icon author: 'Eucalyp' from www.flaticon.com
ver: 1.0.1""")
window.rowconfigure([0],minsize=300, weight=1)
window.columnconfigure([0, 1],minsize=10, weight=1)

buttonFrame = tk.Frame(window, bd= 7, bg='grey')
buttonOpen = tk.Button(buttonFrame, text='Открыть файл', command=FileSelect)
buttonHello = tk.Button(buttonFrame, text='Credits', command=write_hello)

buttonOpen.grid(row=0, column=0, ipadx=23, sticky='w')
buttonHello.grid(row=1, column=0, ipadx=23, sticky='w')

buttonFrame.grid(row=0, column=0, sticky='nsew')

logFrame = tk.Frame(window)
logLabel = tk.Label(master=logFrame, text='Статус-бар:', bg='#C0C0C0')
logContainer = tk.Label(master=logFrame,text=f'Сервис ФНС —{Response_Code_Meaning()}', width=100)
logLabel.grid(row=0, column= 1, sticky="new")
logFrame.grid(row=0, column=1, sticky="new")
logContainer.grid(row=1, column=1, sticky="w")



window.mainloop()
