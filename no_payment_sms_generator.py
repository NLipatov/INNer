import xlsxwriter, os
from datetime import date, datetime

workbook = None
worksheet = None

def cheking_saving_directory():
    if os.path.exists(r'C:\Work') == True:
        os.chdir(r'C:\Work')
        pass
    else:
        os.mkdir( r'C:\Work' )
        os.chdir( r'C:\Work' )


def file_name_generator():
    now = datetime.now()
    filename = now.strftime('SMSGEN_'+'%H%M (%S)'+'.xlsx')
    return filename

def gen_workbook_and_sheet():
    global workbook, worksheet, bold_font, date_format
    workbook = xlsxwriter.Workbook(f'{file_name_generator()}')
    worksheet = workbook.add_worksheet()
    worksheet.ignore_errors({'number_stored_as_text': 'A1:R150'})
    bold_font = workbook.add_format({'bold': True})
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})


def generate_row(airwaybillnum, telephonenum, consigneename):
    today = date.today()
    row = ['SVO', f'{(airwaybillnum)}', f'{(telephonenum)}', f'{consigneename}', \
            f'DHL информирует: в Ваш адрес ожидается прибытие груза по накладной № {airwaybillnum}. Просим предоставить данные для\
таможенного декларирования, \
пройдя по ссылке https://eshopping.dhl.ru/. Вопросы вы можете задать по e-mail: RUSVOB2C@dhl.ru.\
Тему сообщения просьба указать {airwaybillnum}.', 'none', 'no', 'no', 1, 1, 'EUR', 0.2, 0, 0, today,\
2.0, 0.0, 0.0, ]

    return row



def create_sys_row():
    gen_workbook_and_sheet()
    row = 0
    for column, value in enumerate(tech_row):
        worksheet.write(row, column, value, bold_font)


def create_rowx(row_num, generated_row_text):
    for column, value in enumerate( generated_row_text ):
        worksheet.write( row_num, column, value)
        if row_num == 0:
            worksheet.write( row_num, column, value, bold_font )
        if column == 14:
            worksheet.write( row_num, column, value, date_format )

def file_info():
    file_was_generated_succesfully = os.path.isfile( f'C:\Work\{file_name_generator()}' )
    file_name = f'C:\Work\{file_name_generator()}'
    if file_was_generated_succesfully == True:
        return f'Путь к файлу: {file_name}'

def close_workbook():
    cheking_saving_directory()
    global workbook
    workbook.close()
    workbook = xlsxwriter.Workbook( f'C:\Work\{file_name_generator()}' )


tech_row = ['gtw', 'waybill', 'phone', 'consignee', 'sms_text', 'brkr_fee',\
'brkr_agrmnt', 'psp_data', 'ddp', 'inv_val', 'inv_curr', 'vat', 'cstm_fee', 'dhl_fee', 'arrival', 'wgt',\
'stor_min', 'stor_day']





list = []
lenlist = int(len(list))



def row_gen(row_number, list_of_values):
    for i in range(row_number+1, row_number+2):
        create_rowx( i, generate_row( f'{(list_of_values[0])}', f'{(list_of_values[1])}', f'{list_of_values[2]}' ) )


def create_file():
    global list
    create_sys_row()
    for count, row in enumerate(list):
        row_gen(count, row)
    close_workbook()
    list = []


