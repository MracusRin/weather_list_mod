"""
Тут скоро будет описание скрипта.
Возможно даже на английском.

"""
import json
import codecs
from datetime import datetime, timedelta
from click import clear, secho
from openpyxl import load_workbook

# TODO: Добавь документацию. Missing function or method docstring (missing-function-docstring)
# TODO: цвет добавлен, отчиска добавленна, может что-то еще есть полезного в click?
#   https://click.palletsprojects.com/en/8.0.x/utils/
#   https://habr.com/ru/company/oleg-bunin/blog/551424/
# TODO: попробовать предусмотреть режим работы без цветов, а надо ли оно?
# TODO: довавить эксепшен если фаил уже открыт
#   предусмотреть в таком случае отключение меню редактирования


default_config = {"file_address": "Z:\\Change\\The Address\\To Your File\\",
                  "file_name": "Weather.xlsx",
                  "start_search": 1010,
                  "end_search": 2001}


def check_config(def_config):
    try:
        with codecs.open('list_cfg.json', 'r', 'utf-8-sig') as f:
            config_file = json.load(f)
        if def_config != config_file:
            start_on = True
        else:
            print('Измени конфиг list_cfg.json\nПотом пробуй запускать!')
            input('\nНажми Enter чтобы выйти ')
            start_on = False
        return start_on, config_file
    except FileNotFoundError:
        print('Файл конфигурации list_cfg.json не найден, создан новый файл в папке с программой\n\n'
              'Значение file_address: "Измени на путь к файлу excel согласно шалону"\n'
              'Значение file_name: "Измени на нимя файла с погодой. По умоланию Weather.xlsx"\n'
              'Значение start_search: "Начальная строка поиска. Можно не менять. По умолчаниу 1000"\n'
              'Значение end_search: "Конечная строка поиска. Можно не менять. По умолчаниу 2001"\n\n'
              'Измени конфиг и попробуй запустить еще раз')
        input('\nНажми Enter чтобы выйти ')
        with codecs.open('list_cfg.json', 'w+', 'utf-8-sig') as f:
            json.dump(default_config, f)
        start_on = False
    return start_on, ''


current_date = datetime.now().strftime('%d.%m.%Y')
cell_name = ['\nТемпература: ', 'Влажность: ', 'Давление: ', 'Напряжение: ', 'Частота: ']
weather_measure = [' °С', ' %', ' кПа', ' В', ' Гц']
format_list = ['н/д ', 'н/д ', 'н/д   ', 'н/д', 'н/д']


def search_cell_address(search_date):
    value = ''
    for i in range(config_params["start_search"], config_params["end_search"], ):
        if search_date == str(ws['A' + str(i)].value.strftime('%d.%m.%Y')):
            value = str(i)
            break
    weather_cell_address = ['B' + value, 'C' + value, 'D' + value, 'E' + value, 'F' + value]
    return weather_cell_address


def print_weather(weather_cell_address, date=''):
    value = []
    for i in range(0, len(cell_name)):
        cell_value = str(ws[weather_cell_address[i]].value)
        if cell_value == 'None':
            cell_value = format_list[i]
        value.append(cell_name[i] + cell_value + weather_measure[i])
    if date != '':
        value[0] = value[0][1:]
    print(date, value[0], '|', value[1], '|', value[2], '|', value[3], '|', value[4])


def weather_today():
    clear()
    print(f'Норманьные условия сегодня: {current_date}')
    weather_cell_address = search_cell_address(current_date)
    print_weather(weather_cell_address)


def weather_input():
    clear()
    print('Ввод данных:')
    weather_cell_address = search_cell_address(current_date)
    for i in range(0, len(cell_name)):
        ws[weather_cell_address[i]] = float(input(cell_name[i]))
        ws[weather_cell_address[i]].number_format = '0.00'
    clear()
    print('Сохранение...')
    wb.save(config_params['file_address'] + config_params['file_name'])
    print(f'Сохранение выполнено!\n\nВведенные данные: {current_date}')
    print_weather(weather_cell_address)


def weather_by_date():
    clear()
    search_date = input('Введите дату в формате дд.мм.гггг: ')
    clear()
    print(f'Нормальные условия: {search_date}')
    weather_cell_address = search_cell_address(search_date)
    print_weather(weather_cell_address)


def weather_few_days():
    while True:
        try:
            day = int(input('За колько дней показать погоду? '))
            for i in range(0, day):
                back_date = (datetime.now() - timedelta(days=i)).strftime('%d.%m.%Y')
                weather_cell_address = search_cell_address(back_date)
                print('_' * 113)
                print_weather(weather_cell_address, back_date)
        except ValueError:
            clear()
            secho('Ошибка: всэ погано, ты ввел не число!', bg='red', bold=True)
            continue
        else:
            break


def waaagh():
    secho(r"""
                                              .__     ._.
    __  _  _______   _____   _____      ____  |  |__  | |
    \ \/ \/ /\__  \  \__  \  \__  \    / ___\ |  |  \ | |
     \     /  / __ \_ / __ \_ / __ \_ / /_/  >|   Y  \ \|
      \/\_/  (____  /(____  /(____  / \___  / |___|  / __
                  \/      \/      \/ /_____/       \/  \/
    """, fg='red')


clear()
secho(r"""
__        __              _    _                   ____       ___
\ \      / /  ___   __ _ | |_ | |__    ___  _ __  |___ \     / _ \
 \ \ /\ / /  / _ \ / _` || __|| '_ \  / _ \| '__|   __) |   | | | |
  \ V  V /  |  __/| (_| || |_ | | | ||  __/| |     / __/  _ | |_| |
   \_/\_/    \___| \__,_| \__||_| |_| \___||_|    |_____|(_) \___/
""", fg='yellow')

start_program, config_params = check_config(default_config)

while start_program:
    wb = load_workbook(config_params['file_address'] + config_params['file_name'])
    ws = wb.active
    secho('\n***Главное меню***', bg='yellow', fg='black')
    print('\nДанные сегодня - 0'
          '\nВвести данные  - 1'
          '\nДанные по дате - 2'
          '\nДанные по дням - 3'
          '\nВыход          - 4')
    ans = input('\nВыберете действие: ')
    if ans == '4':
        break
    elif ans == '0':
        weather_today()
    elif ans == '1':
        weather_input()
    elif ans == '2':
        weather_by_date()
    elif ans == '3':
        clear()
        weather_few_days()
    elif ans == '7':
        clear()
        secho('НЕВЕРЫЙ ВВОД! Тут ничего нет, перестать тыкать не те кнопки!', bg='red', bold=True)
    elif ans == '9':
        clear()
        secho('Широкую на широкую!', bg='red', bold=True)
        waaagh()
    else:
        clear()
        secho('Ошибка: не правильный ввод', bg='red', bold=True)
