import easyocr
import string, os, time

from openpyxl import load_workbook
from windowcapture import WindowCapture

#Путь к Excel
xl = r'C:\Users\Andrey\Documents\Geo\result.xlsx'
#Название листа в Excel
list_name = 'data'
#Отступ в Excel 
indent = 7
#Навзвание окна для захвата смотерть  файле windowcapture.py
wincap = WindowCapture(r'IMG-20230623-WA0000.jpg ‎- Фотографии')
#Пауза в секундах
time_sleep = 5

def text_recognition(file):
    reader = easyocr.Reader(["ru"])
    result = reader.readtext(file, detail=0)
    return result

count = 0
while True:
    for i in range(0,len(string.ascii_uppercase)-3,3):

        #Убиваем excel
        os.system("TASKKILL /F /IM EXCEL.EXE")
        
        screenshot = wincap.window_capture()
        result = (text_recognition(screenshot))
        
        workbook = load_workbook(xl)
        work_list = workbook[list_name]

        #Вносим данные 
        #Глубина
        work_list[f'{string.ascii_uppercase[i]}{1 + count}'] = result[2]
        work_list[f'{string.ascii_uppercase[i+1]}{1 + count}'] = result[3]
        #Скорость, м/ч
        work_list[f'{string.ascii_uppercase[i]}{2 + count}'] = result[4]
        work_list[f'{string.ascii_uppercase[i+1]}{2 + count}'] = result[6]
        #Время
        work_list[f'{string.ascii_uppercase[i]}{3 + count}'] = result[7]
        work_list[f'{string.ascii_uppercase[i+1]}{3 + count}'] = str(result[8]).replace('.',':')
        #ТМ,гр
        work_list[f'{string.ascii_uppercase[i]}{4 + count}'] = result[9]
        work_list[f'{string.ascii_uppercase[i+1]}{4 + count}'] = result[10]
        #МН, атм
        work_list[f'{string.ascii_uppercase[i]}{5 + count}'] = result[11]
        work_list[f'{string.ascii_uppercase[i+1]}{5 + count}'] = result[12]
        #Натяжение
        work_list[f'{string.ascii_uppercase[i]}{6 + count}'] = result[13]
        work_list[f'{string.ascii_uppercase[i+1]}{6 + count}'] = result[14]

        work_list[f'{string.ascii_uppercase[i+1]}{7 + count}'] = i
        #Сохранняем данный в Excel и закрываем
        workbook.save(xl)
        workbook.close()
        #Запуск excel
        os.system(f'start {xl}')
        time.sleep(time_sleep)
    count += indent