import tkinter as tk
import no_payment_sms_generator
list_of_values = []



def GUI():
    def data_from_entries_to_list_of_values():
        temp_list = []
        temp_list.append( EntryAWB.get() )
        temp_list.append( EntryTEL.get() )
        temp_list.append( EntryNAME.get() )
        EntryAWB.delete( '0', tk.END )
        EntryTEL.delete( '0', tk.END )
        EntryNAME.delete( '0', tk.END )
        no_payment_sms_generator.list.append( temp_list )
    def quit():
        NPSMSGUI.destroy()

    def create_file():
        no_payment_sms_generator.create_file()
        no_payment_sms_generator.file_name_generator()

    def _onKeyRelease(event):
        ctrl = (event.state & 0x0004) != 0
        if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
            event.widget.event_generate( "<<Cut>>" )

        if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
            event.widget.event_generate( "<<Paste>>" )

        if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
            event.widget.event_generate( "<<Copy>>" )


    NPSMSGUI = tk.Tk()
    NPSMSGUI.title('Нет данных, есть оплата')
    NPSMSGUI.geometry('320x120+700+300')
    NPSMSGUI.bind_all( "<Key>", _onKeyRelease)
    try:
        NPSMSGUI.iconbitmap( r"C:\INNer\INNERICC.ico" )
    except:
        pass
    Frame = tk.Frame(master=NPSMSGUI)
    LabelAWB = tk.Label(master=Frame, text='AWB:', width=14)
    EntryAWB = tk.Entry(master=Frame, width=50)
    LabelTEL= tk.Label(master=Frame, text='Телефон:', width=14)
    EntryTEL = tk.Entry(master=Frame, width=50)
    LabelNAME = tk.Label(master=Frame, text='Имя клиента:', width=14)
    EntryNAME = tk.Entry(master=Frame, width=50)
    buttonNext = tk.Button(master=Frame, text='Сформировать строку', bg='white', width=20, command=data_from_entries_to_list_of_values)
    buttonSave = tk.Button(master=Frame, text='Сформировать файл', bg='white', width=20, command=create_file)

    Frame.grid()

    LabelAWB.grid(row=0, column=0, sticky='w')
    EntryAWB.grid(row=0, column=1)
    LabelTEL.grid(row=1, column=0, sticky='w')
    EntryTEL.grid(row=1, column=1)
    LabelNAME.grid(row=2, column=0, sticky='w')
    EntryNAME.grid(row=2, column=1)
    buttonNext.grid(row=3, column=0)
    buttonSave.grid(row=4, column=0)



    NPSMSGUI.mainloop()
