import easyocr
import time
from openpyxl import load_workbook
from windowcapture import WindowCapture
import string
import os

def text_recognition(file):
    reader = easyocr.Reader(["ru"])
    result = reader.readtext(file, detail=0)
    return result


xl = r'C:\Users\Andrey\Documents\Geo\result.xlsx'

list_name = 'data'

wincap = WindowCapture(r'IMG-20230623-WA0000.jpg ‎- Фотографии')


for i in range(0,len(string.ascii_uppercase)-1,2):

    #Убиваем excel
    os.system("TASKKILL /F /IM EXCEL.EXE")
    
    screenshot = wincap.window_capture()
    result = (text_recognition(screenshot))
    
    workbook = load_workbook(xl)
    work_list = workbook[list_name]

    #Вносим данные 
    #Глубина
    work_list[f'{string.ascii_uppercase[i]}1'] = result[2]
    work_list[f'{string.ascii_uppercase[i+1]}1'] = result[3]
    #Скорость, м/ч
    work_list[f'{string.ascii_uppercase[i]}2'] = result[4]
    work_list[f'{string.ascii_uppercase[i+1]}2'] = result[6]
    #Время
    work_list[f'{string.ascii_uppercase[i]}3'] = result[7]
    work_list[f'{string.ascii_uppercase[i+1]}3'] = str(result[8]).replace('.',':')
    #ТМ,гр
    work_list[f'{string.ascii_uppercase[i]}4'] = result[9]
    work_list[f'{string.ascii_uppercase[i+1]}4'] = result[10]
    #МН, атм
    work_list[f'{string.ascii_uppercase[i]}5'] = result[11]
    work_list[f'{string.ascii_uppercase[i+1]}5'] = result[12]
    #Натяжение
    work_list[f'{string.ascii_uppercase[i]}6'] = result[13]
    work_list[f'{string.ascii_uppercase[i+1]}6'] = result[14]

    workbook.save(xl)
    workbook.close()
    #Запуск excel
    os.system(f'start {xl}')
    time.sleep(5)