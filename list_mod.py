import openpyxl 
import time
#from datetime import datetime, timedelta
from datetime import date, timedelta, datetime

current_date = datetime.now().strftime('%d.%m.%Y')
file_address='Y:\\Супер общий зал\\Нормальные условия.xlsx'
#file_address='E:\\OneDrive\\Programming\\Python\\project\\exel\\weather.xlsx'
datecol=7 # Номер колонки где сохранена дата в файле
chort=0
start_search = 1000
end_searhc = 2001
wb = openpyxl.load_workbook(file_address)
ws = wb.active    
title_parameter = ['\nТемпература: ', 'Влажность: ', 'Давление: ', 'Напряжение: ', 'Частота: ']
weather_unit = [' °С',' %',' кПа',' В',' Гц']

def search_line_cell(search_date):
    #for i in range(1, ws.max_row + 1, ): # поиск по максимуму, до конца документа
    for i in range(start_search, end_searhc, ):
        if search_date == ws.cell(i,datecol).value:
            data_line = str(ws.cell(i,datecol).row)
    weather = ['B' + data_line,'C' + data_line,'D' + data_line,'E' + data_line,'F' + data_line]
    return weather  

def print_weather (weather_cell):
    a=[]
    for i in range(0, len(title_parameter)):
        item=str(ws[weather_cell[i]].value)
        if item == 'None':
            item = 'н/д '
        a.append(title_parameter[i] + item + weather_unit[i])
    print(a[0],'|', a[1],'|', a[2],'|', a[3],'|', a[4])

print(f'Норманьные условия сегодня: {current_date}')
weather_cell=search_line_cell(current_date) 
print_weather(weather_cell)

ans = input('\nВвести данные  - 1\nДанные по дате - 2\nДанные по дням - 3\n\nВыберете действие: ')
if ans == '1':
    # можно вынести в начало и проверять условие там
    try: 
        myfile = open(file_address, "r+") # or "a+", whatever you need
    except IOError:
        print ('\n!!! Какойто ЧОРТ уже открыл твой файл !!!\nВвод отмен')
        chort=1

    if chort != 1:
        
        for i in range(0, len(title_parameter)):
            ws[weather_cell[i]] = float(input(title_parameter[i]))
            ws[weather_cell[i]].number_format='0.00'           

        print('\nСохранение...')
        wb.save(file_address)
        print('Сохранение выполнено!')

        print(f'\nВведенные данные: {current_date}')
        print_weather(weather_cell)

elif ans == '2':
    search_date=input('Введите дату в формате дд.мм.гггг: ')
    print(f'\nНормальные условия: {search_date}')
    weather_cell = search_line_cell(search_date) 
    print_weather(weather_cell)

elif ans == '3':
    day=int(input(f'\nЗа колько дней показать погоду? '))
    for i in range(0,day+1):
        back_date = (datetime.now() - timedelta(days=i)).strftime('%d.%m.%Y')
        weather_cell = search_line_cell(back_date) 
        print('-'*102)
        print(back_date)
        print_weather(weather_cell)
        
else:
    print('Ввод отмен')

input('\nНажмите чтонить для выхода ')

