from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as mb
import AWBF

def Clear():
    AWBF.clear()

def Main():
    window = Tk()
    try:
        window.iconbitmap( r"C:\INNer\INNERICC.ico" )
    except:
        pass
    window.withdraw()
    folder_selected = filedialog.askdirectory(title='Выберите папку')
    return_from_fetcher = AWBF.FetchAWB(folder_selected)
    if return_from_fetcher != 'ERR':
        show_info = mb.showinfo ('Готово!', f'Список из {AWBF.result_info} а/н скопирован в буфер обмена.')
    else:
        pass
    window.destroy()
