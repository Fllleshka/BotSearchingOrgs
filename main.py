# Импорт данных из фаила
from dates import *

# Импорт библиотек
import requests
import json
import pprint
import win32com.client



# Экземпляр COM обьекта
xlApp = win32com.client.Dispatch("Excel.Application")
# Открываем фаил
xlwb = xlApp.Workbooks.Open("Z:\PythonProjects\ComixBot\export.xlsx")
# Выбираем лист(таблицу)
sheet = xlwb.ActiveSheet
# Выбираем данные из range
alldates = sheet.Range("A1:A100").Value
laststr = 1
for elem in alldates:
    if elem[0] != None:
        laststr += 1
    else:
        continue
print(laststr)

requestfull = url + "?text=" + requeststext + "&lang=ru_RU&apikey=" + apikey + "&results=" + str(results) +  "&skip=" + str(skip)

# + "&ll=" + coordinates + "&spn=" + spn
print(requestfull)
r = requests.get(requestfull)

result =  r.json()
#print(result)
#print(len(result["features"]))

for element in range(0, len(result["features"])):
    dates = result["features"][element]["properties"]["CompanyMetaData"]
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

#сохраняем рабочую книгу
xlwb.Save()

#закрываем ее
xlwb.Close()

#закрываем COM объект
xlApp.Quit()