# Импорт данных из фаила
from dates import *
from functions import *

# Формирование массива данных
res = colletdates()
# Создание файла
pathfile = createexcelfile()
# Запись данных в файл
insertdates(pathfile, res)