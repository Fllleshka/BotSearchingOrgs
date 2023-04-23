# Импорт данных из фаила dates.py
from dates import *

# Импорт библиотек
import requests
import win32com.client
import os.path

# Функция формирования общего списка данных
def colletdates():
    # Количество результатов
    results = 500
    # Количество пропущенных результатов
    skip = 0
    # Массив результатов
    massresults = []
    

    # Массив для наполения результурующего массива
    while results == 500:
        # Кладём результат запроса в переменную
        resultrequest = importdates(results, skip)
        # Пробегаемся по массиву и добавляем в результурующий массив
        for element in range(0, len(resultrequest)):
            massresults.append(resultrequest[element])
        # Выбираем сколько записей пропускаем
        skip += results
        # Записываем количество записей 
        results = len(resultrequest)
        # Ограничение из документации Яндекса
        if skip == 1500:
            break
    
    return massresults

# Функция импорта данных
def importdates(results, skip):
    # Формирование строки для запроса
    requestfull = url + "?text=" + requeststext + "&lang=ru_RU&apikey=" + apikey + "&results=" + str(results) +  "&skip=" + str(skip)
    # Формирование запроса типа GET и преобразование в формат JSON
    result = requests.get(requestfull).json()
    # Результирующий массив
    massresult = []
    # Пробегаемся по массиву и добавляем в результурующий массив нужные данные
    for element in range(0, len(result["features"])):
        massresult.append(result['features'][element]['properties'])
    # Возвращаем результат
    return massresult

# Функция вычисления последней строки в в excel файле
def laststr(namefile):
    # Экземпляр COM обьекта
    xlApp = win32com.client.Dispatch("Excel.Application")
    path = os.getcwd() + '/' + str(namefile)
    print(path)
    # Создаём файл
    xlwb = xlApp.Workbooks.Add()
    xlwb.SaveAs(path)
    #xlwb = xlApp.Workbooks.SaveAs(path)
    xlwb = xlApp.Workbooks.Open(path)
    # Выбираем лист(таблицу)
    sheet = xlwb.ActiveSheet
    # Даём инфорацию о таблице
    sheet.Cells(1, 1).value = "Текст запроса"
    sheet.Cells(1, 2).value = requeststext
    # Называем столбцы
    sheet.Cells(2, 1).value = "ID ораганизации"
    sheet.Cells(2, 2).value = "Называние"
    sheet.Cells(2, 3).value = "Сайт"
    sheet.Cells(2, 4).value = "Адрес"
    sheet.Cells(2, 5).value = "Телефон1"
    sheet.Cells(2, 6).value = "Телефон2"

    sheet.column_dimensions['A'].width = 110
    sheet.column_dimensions['B'].width = 380

    # Выбираем данные из range
    alldates = sheet.Range("A1:A10000").Value
    laststr = 1
    for elem in alldates:
        if elem[0] != None:
            laststr += 1
        else:
            continue
    #сохраняем рабочую книгу
    xlwb.Save()

    #закрываем ее
    xlwb.Close()

    #закрываем COM объект
    xlApp.Quit()
    
    return laststr

def insertdates():
    resultrequest = []
    for element in range(0, len(resultrequest["features"])):
        dates = resultrequest["features"][element]["properties"]["CompanyMetaData"]
        #pprint.pprint(dates)
        id = dates["id"]
        name = dates["name"]
        try:
            urlorg = dates['url']
        except:
            urlorg = "нету"
        address = dates["address"]
        try:
            phones = dates["Phones"]
            phonesall = []
            try:
                for elem in phones:
                    phonesall.append(elem["formatted"])
            except:
                phonesall = "нету"
        except:
            phones = "нету"
            phonesall = ["нету"]

        if phonesall[0] == "нету":
            print(f"Не записываем | {id} | {name} | {urlorg} | {address} | {phonesall}")
        else:
            sheet = []
            print(f"{id} | {name} | {urlorg} | {address} | {phonesall}")
            sheet.Cells(laststr, 1).value = id
            sheet.Cells(laststr, 2).value = name
            sheet.Cells(laststr, 3).value = urlorg
            sheet.Cells(laststr, 4).value = address
            try:
                sheet.Cells(laststr, 5).value = phonesall[0]
            except:
                sheet.Cells(laststr, 5).value = "отсуствует"
            try:
                sheet.Cells(laststr, 6).value = phonesall[1]
            except:
                sheet.Cells(laststr, 6).value = "отсуствует"
            laststr += 1