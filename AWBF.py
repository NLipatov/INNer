import xlrd, pyperclip, re, os
from tkinter import *
from tkinter import messagebox as mb

ValueString = ''
result_info = None


def clear():
    global ValueString
    ValueString = ''
    # pyperclip.copy(ValueString)


def FetchAWB(path):
    global result_info
    path = path
    if (os.path.exists(path) == True) and (len(os.listdir(path)) > 0):
        pass
    else:
        if len(os.listdir(path)) < 1:
            mb.showerror('Ошибка', 'Директория пуста!')
            return 'ERR'
        else:
            mb.showerror( 'Неизвестная ошибка', 'Окно будет закрыто' )
            return 'ERR'
    FileList = os.listdir(f'{path}')
    while True:
        for i in FileList:
            if i[-1] == 's' and i[-2] == 'l' and i[-3] == 'x' and i[-4] == '.':
                pass
            else:
                FileList.remove(i)
                # print(f'Исключен из фетч-листа файл \'{i}\'')
        break
    for i in FileList:
        workbook = xlrd.open_workbook(r'%s\%s' % (path,i))
        def Sheet_quantity():
            counter = 0
            for i in workbook.sheets():
                counter += 1
            return counter
        def Sheet_iterator():
            global ValueString
            for i in range(0, Sheet_quantity()):
                sheet = workbook.sheet_by_index(i-1)
                for c in range(1, sheet.nrows):
                    a = str(sheet.cell_value(c, 0)) + '\n'
                    ValueString += a
        Sheet_iterator()
        result = re.sub(r'\n+', '\n', ValueString)
        resulta = re.sub(r' ', '', result)
        resultb = resulta.count('\n')
        pyperclip.copy(resulta)
        result_info = resultb
