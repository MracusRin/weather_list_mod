import openpyxl
import time
from datetime import datetime, timedelta
from art import *

current_date = datetime.now().strftime('%d.%m.%Y')
# file_address='Y:\\Супер общий зал\\Нормальные условия.xlsx'
file_address = 'E:\\OneDrive\\Programming\\Python\\project\\exel\\weather.xlsx'
datecol = 7
chort = 0
start_search = 1000
end_search = 2001
wb = openpyxl.load_workbook(file_address)
ws = wb.active
title_parameter = ['\nТемпература: ', 'Влажность: ', 'Давление: ', 'Напряжение: ', 'Частота: ']
weather_unit = [' °С', ' %', ' кПа', ' В', ' Гц']
format_list = ['н/д ', 'н/д ', 'н/д   ', 'н/д', 'н/д']


def search_line_cell(search_date):
    for i in range(start_search, end_search, ):
        if search_date == ws.cell(i, datecol).value:
            data_line = str(ws.cell(i, datecol).row)
    weather = ['B' + data_line, 'C' + data_line, 'D' + data_line, 'E' + data_line, 'F' + data_line]
    return weather


def print_weather(weather_cell, date=''):
    a = []
    for i in range(0, len(title_parameter)):
        item = str(ws[weather_cell[i]].value)
        if item == 'None':
            item = format_list[i]
        a.append(title_parameter[i] + item + weather_unit[i])
    if date != '':
        a[0] = a[0][1:]
    print(date, a[0], '|', a[1], '|', a[2], '|', a[3], '|', a[4])


tprint("Weather 2.0")


def main():
    try:
        myfile = open(file_address, "r+")  # or "a+", whatever you need
    except IOError:
        print('\n!!! Какойто ЧОРТ уже открыл твой файл !!!\nВвод отмен')
        chort = 1

    ans = input(
        '\n***Главное меню***\nДанные сегодня - 0\nВвести данные  - 1\nДанные по дате - 2\nДанные по дням - 3\nВыход  '
        '        - 4\nВыберете действие: ')

    if ans == '0':
        print(f'\nНорманьные условия сегодня: {current_date}')
        weather_cell = search_line_cell(current_date)
        print_weather(weather_cell)

    elif ans == '1':
        print('\nВвод данных:')
        weather_cell = search_line_cell(current_date)
        for i in range(0, len(title_parameter)):
            ws[weather_cell[i]] = float(input(title_parameter[i]))
            ws[weather_cell[i]].number_format = '0.00'
        print('\nСохранение...')
        wb.save(file_address)
        print(f'Сохранение выполнено!\n\nВведенные данные: {current_date}')
        print_weather(weather_cell)

    elif ans == '2':
        search_date = input('Введите дату в формате дд.мм.гггг: ')
        print(f'\nНормальные условия: {search_date}')
        weather_cell = search_line_cell(search_date)
        print_weather(weather_cell)

    elif ans == '3':
        day = int(input(f'\nЗа колько дней показать погоду? '))
        for i in range(0, day):
            back_date = (datetime.now() - timedelta(days=i)).strftime('%d.%m.%Y')
            weather_cell = search_line_cell(back_date)
            print('_' * 113)
            print_weather(weather_cell, back_date)

    elif ans == '4':
        print(art("random"))
        exit()

    else:
        print('Не верный ввод')

    main()


main()
