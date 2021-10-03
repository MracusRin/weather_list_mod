import openpyxl 
import time
from datetime import datetime

current_date = datetime.now().strftime('%d.%m.%Y')
fileadd='Y:\\Супер общий зал\\Нормальные условия.xlsx'
datacol=7 # Номер колонки где сохранена дата в файле
chort=0
start_search = 1000
end_searhc = 2001
wb = openpyxl.load_workbook(fileadd)
ws = wb.active    

#Можно принимать значинея от пользователя в массивы!!!!!!!!

def search_line(search_date):
    #for i in range(1, ws.max_row + 1, ): # поиск по максимуму, до конца документа
    for i in range(start_search, end_searhc, ): #диапазаон поиска даты
        if search_date == ws.cell(i,datacol).value:
            data_line = str(ws.cell(i,datacol).row)
        
    return data_line     

dline=search_line(current_date)                 
tempcell = 'B' + dline
humcell  = 'C' + dline
prescell = 'D' + dline
voltcell = 'E' + dline
freqcell = 'F' + dline

print(f'Норманьные условия сегодня: {current_date}')
print(f'\nТемпература: {ws[tempcell].value} °С | Влажность: {ws[humcell].value} % | Давление: {ws[prescell].value} кПа | Напряжение: {ws[voltcell].value} В | Частота: {ws[freqcell].value} Гц\n')

ans = input('Ввести данные - нажмите 1. Посмотреть данные по дате - нажмите 2: ')
if ans == '1':
    # можно вынести в начало и проверять условие там
    try: 
        myfile = open(fileadd, "r+") # or "a+", whatever you need
    except IOError:
        print ('\n!!! Какойто ЧОРТ уже открыл твой файл !!!\nВвод отмен')
        chort=1

    if chort != 1:

        ws[tempcell] = float(input('\nТемпература: '))
        ws[tempcell].number_format='0.0'
      
        ws[humcell] = float(input('Влажность: '))
        ws[humcell].number_format='0.0'

        ws[prescell] = float(input('Давление: '))
        ws[prescell].number_format='0.00'

        ws[voltcell] = float(input('Напряжение: '))
        ws[voltcell].number_format='0.0'

        ws[freqcell] = float(input('Частота: '))
        ws[freqcell].number_format='0.0'

        print('\nСохранение...')
        wb.save(fileadd)
        print('Сохранение выполнено!')

        print(f'\nВведенные данные: {current_date}\nТемпература: {ws[tempcell].value} °С | Влажность: {ws[humcell].value} % | Давление: {ws[prescell].value} кПа | Напряжение: {ws[voltcell].value} В | Частота: {ws[freqcell].value} Гц\n')
elif ans == '2':
    search_date=input('Введите дату в формате дд.мм.гггг: ')
    sline = search_line(search_date) 
    stempbox = 'B' + sline
    shumbox  = 'C' + sline
    spresbox = 'D' + sline
    svoltbox = 'E' + sline
    sfreqbox = 'F' + sline
    print(f'\nНормальные условия: {search_date}\nТемпература: {ws[stempbox].value} °С | Влажность: {ws[shumbox].value} % | Давление: {ws[spresbox].value} кПа | Напряжение: {ws[svoltbox].value} В | Частота: {ws[sfreqbox].value} Гц\n')      
    time.sleep(7)
else:
    print('Ввод отмен')

time.sleep(3)

