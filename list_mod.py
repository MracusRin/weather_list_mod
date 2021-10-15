from datetime import datetime, timedelta
from colorama import init, Fore, Back, Style
from art import *
import openpyxl
import json
import codecs

init()

# TODO: использовать вместо import codecs CLICK
#   https://habr.com/ru/company/oleg-bunin/blog/551424/
# TODO: попробовать предусмотреть режим работы без цветов
# TODO: зупускать главное меню функцией и все вызовы перпеместить в функции
# TODO: довавить эксепшен если фаил уже открыт
#   предусмотреть в таком случае отключение меню редактирования
# TODO: убрать art и земенить его на простой print
# TODO: попробовать перещитывать формант даты Экселя в нормальную дату


with codecs.open('list_cfg_home.json', 'r', 'utf-8-sig') as f:
    config_paramm = json.load(f)

current_date = datetime.now().strftime('%d.%m.%Y')
cell_name = ['\nТемпература: ', 'Влажность: ', 'Давление: ', 'Напряжение: ', 'Частота: ']
weather_measure = [' °С', ' %', ' кПа', ' В', ' Гц']
format_list = ['н/д ', 'н/д ', 'н/д   ', 'н/д', 'н/д']

wb = openpyxl.load_workbook(config_paramm['file_address'])
ws = wb.active


def search_cell_address(search_date):
    date_line = ''
    for i in range(config_paramm["start_search"], config_paramm["end_search"], ):
        if search_date == str(ws['A'+str(i)].value.strftime('%d.%m.%Y')):
            date_line = str(i)
            break
    weather_cell_address = ['B' + date_line, 'C' + date_line, 'D' + date_line, 'E' + date_line, 'F' + date_line]
    return weather_cell_address


def print_weather(weather_cell_addres, date=''):
    a = []
    for i in range(0, len(cell_name)):
        cell_value = str(ws[weather_cell_addres[i]].value)
        if cell_value == 'None':
            cell_value = format_list[i]
        a.append(cell_name[i] + cell_value + weather_measure[i])
    if date != '':
        a[0] = a[0][1:]
    print(date, a[0], '|', a[1], '|', a[2], '|', a[3], '|', a[4])


print(Fore.YELLOW)
tprint("Weather 2.0")
print(Style.RESET_ALL)


def main():
    print('\n',Back.YELLOW + Fore.BLACK + '**Главное меню**' + Style.RESET_ALL)
    ans = input(
        '\nДанные сегодня - 0\nВвести данные  - 1\nДанные по дате - 2\nДанные по дням - 3\nВыход  '
        '        - 4\nВыберете действие: ')

    if ans == '0':
        print(f'\nНорманьные условия сегодня: {current_date}')
        weather_cell_addres = search_cell_address(current_date)
        print_weather(weather_cell_addres)

    elif ans == '1':
        print('\nВвод данных:')
        weather_cell_addres = search_cell_address(current_date)
        for i in range(0, len(cell_name)):
            ws[weather_cell_addres[i]] = float(input(cell_name[i]))
            ws[weather_cell_addres[i]].number_format = '0.00'
        print('\nСохранение...')
        wb.save(config_paramm['file_address'])
        print(f'Сохранение выполнено!\n\nВведенные данные: {current_date}')
        print_weather(weather_cell_addres)

    elif ans == '2':
        search_date = input('Введите дату в формате дд.мм.гггг: ')
        print(f'\nНормальные условия: {search_date}')
        weather_cell_addres = search_cell_address(search_date)
        print_weather(weather_cell_addres)

    elif ans == '3':
        day = int(input(f'\nЗа колько дней показать погоду? '))
        for i in range(0, day):
            back_date = (datetime.now() - timedelta(days=i)).strftime('%d.%m.%Y')
            weather_cell_addres = search_cell_address(back_date)
            print('_' * 113)
            print_weather(weather_cell_addres, back_date)

    elif ans == '4':
        exit()
    elif ans == '6':
        print('\n', Fore.YELLOW + art("random") + Style.RESET_ALL)

    elif ans == '8':
        print('\nНЕВЕРЫЙ ВВОД! Тут ничего нет, перестать тыкать не те кнопки!')

    elif ans == '9':
        print(
            Fore.YELLOW + '\nШирокую на широкую!' +  Style.RESET_ALL)
        tprint("Waaagh!", "graffiti")
    else:
        print(Fore.RED + '\nНе верный ввод ' + Style.RESET_ALL)

    main()


main()
