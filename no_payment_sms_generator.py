import xlsxwriter, os
from datetime import date, datetime

workbook = None
worksheet = None


def file_name_generator():
    now = datetime.now()
    filename = now.strftime('SMSGEN_'+'%H%M (%S)'+'.xlsx')
    return filename

def gen_workbook_and_sheet():
    global workbook, worksheet
    workbook = xlsxwriter.Workbook(f'{file_name_generator()}')
    worksheet = workbook.add_worksheet()
    worksheet.ignore_errors({'number_stored_as_text': 'A1:R150'})
# cell_format = workbook.add_format()
# cell_format.set_num_format("0.0")

def generate_row(airwaybillnum, telephonenum, consigneename):
    today = date.today()
    dt = today.strftime( "%d.%m.%Y" )
    row = ['SVO', f'{int(airwaybillnum)}', f'{int(telephonenum)}', f'{consigneename}', \
            f'DHL информирует: в Ваш адрес ожидается прибытие груза по накладной № {airwaybillnum}. Просим предоставить данные для\
таможенного декларирования, \
пройдя по ссылке https://eshopping.dhl.ru/. Вопросы вы можете задать по e-mail: RUSVOB2C@dhl.ru.\
Тему сообщения просьба указать {airwaybillnum}.', 'none', 'no', 'no', 1, 1, 'EUR', 0.2, 0, 0, str( dt ),\
2.0, 0.0, 0.0, ]

    return row



def create_sys_row():
    gen_workbook_and_sheet()
    row = 0
    for column, value in enumerate(tech_row):
        worksheet.write(row, column, value)


def create_rowx(row_num, generated_row_text):
    for column, value in enumerate( generated_row_text ):
        worksheet.write( row_num, column, value)
        # if column == 8:
        #     worksheet.set_row(row_num, 15, cell_format )

def close_workbook(): #может использоваться как кнопка сформировать файл
    global workbook
    workbook.close()
    workbook = xlsxwriter.Workbook( f'{file_name_generator()}' )



directory = os.chdir(r'C:\Users\Admin\Desktop\INNer project\SMS')


tech_row = ['gtw', 'waybill', 'phone', 'consignee', 'sms_text', 'brkr_fee',\
'brkr_agrmnt', 'psp_data', 'ddp', 'inv_val', 'inv_curr', 'vat', 'cstm_fee', 'dhl_fee', 'arrival', 'wgt',\
'stor_min', 'stor_day']





list = []
lenlist = int(len(list))
# list.append(['1337','79998143975', 'Gregory M. Bryant']) #добавление в лист листов значения с кнопки
# list.append(['1338','79998143977', 'Pat Galsinger'])
# list.append(['1340','79998143980', 'Lisa Su'])


def row_gen(row_number, list_of_values):
    for i in range(row_number+1, row_number+2):
        create_rowx( i, generate_row( f'{int(list_of_values[0])}', f'{int(list_of_values[1])}', f'{list_of_values[2]}' ) )


def create_file():
    global list
    create_sys_row()
    for count, row in enumerate(list):
        row_gen(count, row)
    close_workbook()
    list = []


