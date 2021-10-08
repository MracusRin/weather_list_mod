from datetime import datetime, timedelta
from colorama import init, Fore, Back, Style
from art import *
import openpyxl
import json
import codecs

init()

with codecs.open('list_cfg.json', 'r', 'utf-8-sig') as f:
    cfg = json.load(f)

current_date = datetime.now().strftime('%d.%m.%Y')
title_parameter = ['\nТемпература: ', 'Влажность: ', 'Давление: ', 'Напряжение: ', 'Частота: ']
weather_unit = [' °С', ' %', ' кПа', ' В', ' Гц']
format_list = ['н/д ', 'н/д ', 'н/д   ', 'н/д', 'н/д']

wb = openpyxl.load_workbook(cfg['file_address'])
ws = wb.active


def search_line_cell(search_date):
    data_line = ''
    for i in range(cfg["start_search"], cfg["end_search"], ):
        if search_date == ws.cell(i, cfg["date_column"]).value:
            data_line = str(ws.cell(i, cfg["date_column"]).row)
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


print(Fore.YELLOW)
tprint("Weather 2.0")
print(Style.RESET_ALL)


def main():
    # try:
    #     my_file = open(file_address, "r+")  # or "a+", whatever you need
    # except IOError:
    #     print('\n!!! Какойто ЧОРТ уже открыл твой файл !!!\nВвод отмен')
    print(Back.YELLOW + Fore.BLACK + '\n***Главное меню***' + Style.RESET_ALL)
    ans = input(
        '\nДанные сегодня - 0\nВвести данные  - 1\nДанные по дате - 2\nДанные по дням - 3\nВыход  '
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
        wb.save(cfg['file_address'])
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
        exit()
    elif ans == '6':
        print('\n', Fore.YELLOW + art("random") + Style.RESET_ALL)

    elif ans == '8':
        print('\nНЕВЕРЫЙ ВВОД! Тут ничего нет, перестать тыкать не те кнопки!')

    elif ans == '9':
        print(
            Fore.YELLOW + '\nРомчег, автор приложухи сей.\nБлагодарность ему и низкий поклон.\nЛучшему из людей.\nИ '
                          'вообще себя не похвалишь, никто не похвалит.\nА теперь харош глазеть. Иди работай!' +
            Style.RESET_ALL)
        tprint("Waaagh!", "graffiti")
    else:
        print(Fore.RED + '\nНе верный ввод ' + Style.RESET_ALL)

    main()


main()
