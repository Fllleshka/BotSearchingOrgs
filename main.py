# Импорт данных из фаила
from dates import *
from functions import *

# Импорт библиотек
import pprint

#laststr = laststr("export.xlsx")
#print(laststr)

res = colletdates()
print(f"Количество импортированных записей: {len(res)}")

pathfile = createexcelfile()
print(f"Путь к фаилу: {pathfile}")

