# Данный скрипт делает скриншоты страниц
# Читает адрес сайта с таблицы Excel и
# запускает браузер Firefox, тот переходит
# по ссылке и делает скрин.
# Результат сохраняется в указанную ранее
# папку и его путь добавляется в список
# По окончанию работы списки сохраняются
# в Excel документ. При повторной работе
# скрипт может обновить скриншоты или
# не делать новые. Выбор за вами

# ------------------------------------------------------------------------------

#   python.exe -m pip install --upgrade pip    |  Обновление установщика pip
#   pip install -U selenium                    |  Установка модуля selenium
#   pip install pandas                         |  Установка модуля pandas
#   pip install xlrd                           |  Установка модуля xlrd
#   pip install openpyxl                       |  Установка модуля openpyxl
#   pip install Pillow                         |  Установка модуля Pillow

#   py F:\Python\WebScreen\main.py             |  Запуск скрипта с консоли CMD

# ------------------------------------------------------------------------------

#   Импортируем зависимости необходимые для работы скрипта
from PIL import Image                   # Для преобразования PNG в JPG
import time                             # Работа со временем
import random                           # Для генирации уникальных имен файлов
from selenium import webdriver          # Две зависимости для работы с Firefox
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import pandas                           # Для работы с форматом Excel

# ------------------------------------------------------------------------------


#(ВНИМАНИЕ Перезапишет файл) # Если нужно обновить старые скрины то ставим 100
UPD_SCREEN = 100

#   Каталог со скриптом main.py
MAIN_DIR    = 'F:/Python/WebScreen/'

#   Файл с Excel таблицей
EXEL_FILE   = MAIN_DIR+'assets/data/input.xlsx'

#   Имя листа в Excel файле
SHEET       = 'input'

#   Директория сохранения скриншотов
IMG_PATH    = MAIN_DIR+'assets/images/'

#   Добавляет формат в конец файла
IMG_END     = '.jpg'

#   Полный путь к портативной версии браузера Firefox
FFOX        = MAIN_DIR+'bin/Firefox64/firefox.exe'

#   Полный путь к драйверу для работы браузера Firefox
FFOX_DR     = MAIN_DIR+'bin/geckodriver.exe'

# ------------------------------------------------------------------------------

wr_timer0 = time.time()    # Старт программы


# Экземпляр браузера Firefox
binary = FirefoxBinary(FFOX)
# Драйвер для работы с Firefox
driver = webdriver.Firefox(firefox_binary=binary, executable_path=FFOX_DR)

# ------------------------------------------------------------------------------


#   Функция создает уникальный идентификатор
def getUID():
    seed = random.getrandbits(32)
    while True:
       yield seed
       seed += 1

# Есть ли данные о скриншоте
def isData(s):
    return bool(len(s)>5)

#   Функция делает скриншот через браузер
#   принимает адресс страницы и путь сохранения для скриншота
#   хранит данные в формате PNG
def getPageFullScreen(url, path):
    driver.get(url) # передаем адресс в драйвер
    el = driver.find_element_by_tag_name('body') # драйвер ищет элемент body
    el.screenshot(path)    # сохраняет скриншот по указанному пути
     # сохраняем в JPG - 3 строчки костыль
    im = Image.open(path)  # Открываем PNG
    rgb_im = im.convert('RGB')  # Извлекаем RGB спектр
    rgb_im.save(path)

#   Функция учавствует в развязке
def denouement(url, pts, index, index_max):
    # Чтобы понимать что происходит - пишем в консоль
    print('Ссылка ' + str(index) + ' из ' + str(index_max))
    getPageFullScreen(url, pts)    # делаем скриншот сайта
# ------------------------------------------------------------------------------


# Читаем файл Excel
cols = [0, 1, 2, 3];
data = pandas.read_excel(EXEL_FILE, sheet_name=SHEET, header=None, usecols=cols)

#   Запишем содержимое в списки
A1  = data[0].tolist() #   Индекс
B1  = data[1].tolist() #   Название
C1  = data[2].tolist() #   Ссылка
D1  = data[3].tolist() #   Имена скринов


#   Создадим пустой список для записи пути к скринам
DD1        = []        #   Имя и путь скрина
ITEMS_SIZE = len(C1)   #   Получаем количество элементов
UID        = getUID()  #   Уникальные значения

# ------------------------------------------------------------------------------


wr_timer1 = time.time()    # Начинаем читать
print('Начинаем читать. С момента запуска прошло секунд: ', wr_timer1-wr_timer0)


#   Бежим в цикле по нашим Excel ссылкам
for i in range(ITEMS_SIZE):
    if C1[i]==C1[i]:
    # Если ссылка равна самой себе вернет правду и лож если там NaN
        if not isData(str(D1[i])):
        # Если скриншот не делали
            id = next(UID) # Уникальный индекс
            #   формируем путь сохранения файла
            IMG_PTS = IMG_PATH + str(id) + IMG_END
            # Делаем скриншот
            denouement(C1[i], IMG_PTS, i, ITEMS_SIZE)
            #   Добовляем путь скрина в список
            DD1.append(IMG_PTS)
            #   Для понимания происходящего пишем в консоль
            print('Новый скриншот сохранен: ' + IMG_PTS)
        else:
        #   Если уже делали скриншот этой страницы
            if UPD_SCREEN == 100:  # Делать скриншот повторно?
                denouement(C1[i], str(D1[i]), i, ITEMS_SIZE)
                #   Добовляем путь скрина в список
                DD1.append(str(D1[i]))
                #   Для понимания происходящего пишем в консоль
                print('Скриншот ' + str(D1[i]) + ' был обновлен')
            else: # Оставить без изменений если UPD_SCREEN != 100
                DD1.append(str(D1[i]))   #   Перезапишем его в список
                #   Для понимания происходящего пишем в консоль
                print('Скриншот этой страницы уже есть')
    else:
    #   Если ссылка отсутствует в Excel то будет NaN и выполнится это условие
        # Запишем пустоту в DD1, его длина должна быть равна ITEMS_SIZE
        DD1.append('')
        #   Для понимания происходящего пишем в консоль
        print('В этой строке нет адреса')


# ------------------------------------------------------------------------------


wr_timer2 = time.time()    # Прочитали файл
print('Прочитали. С момента чтения прошло секунд: ', wr_timer2-wr_timer1)

driver.quit()   #   После окончания работы закрываем наш браузер




# Убедимся что индексы списков равны - выводим в консоль
print('D1 = ' + str(len(D1)),  'DD1 = ' + str(len(DD1)))


# Сохраним наши списки в DataFrame
data_final = {'A': A1, 'B': B1, 'С': C1, 'D': DD1}
data_frame = pandas.DataFrame(data=data_final)

#   Сохраняем DataFrame в наш файл - это перезапишет исходный Excel файл
data_frame.to_excel(EXEL_FILE, sheet_name=SHEET, index=False, header=None)


wr_timer3 = time.time()    # Общее время работы

#   Сообщаем о том что программа завершена и выводим время работы программы
print('Успешно завершен. Время затрачено в секундах: ', wr_timer3-wr_timer0)

# ------------------------------------------------------------------------------
