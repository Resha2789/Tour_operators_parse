import re

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time
import win32com.client as win32
import os
import keyboard


# Входные данные
class IntPut:
    inp_months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    inp_city = None
    inp_tur_operator = []
    inp_date = dict(start='', end='')
    inp_night = dict(start=7, end=14)
    inp_rewriting = False
    inp_show_browser = False
    inp_services = []


# Возвращаемые данные
class OutPut:
    # Услуги в отеле
    out_services = []
    # Цена
    out_price = []
    # Питание
    out_food = []
    # Туроператор
    out_tour_operator = []
    # Даты
    out_date = []
    # Пляжная линия
    out_beach_line = []
    # Расстояние от отеля до аэропорта
    out_distance_to_airport = []
    # Количество звезд отеля
    out_count_stars = []
    # Рейтинг отеля
    out_rating = []
    # Название отеля
    out_hotel_name = []
    # Город
    out_city = []
    # Количество ночей
    out_night = []
    # Рейтинг отеля
    out_hotel_rating = []


# Запуск Excel
class LoadExcel:
    def __init__(self):
        self.file_name = "Data.xlsx"
        self.Excel = None
        self.work_book = None
        self.sheet_inp = None
        self.sheet_out = None
        self.open_excel()

    def open_excel(self):

        # noinspection PyBroadException
        try:
            # COM объект
            self.Excel = win32.Dispatch("Excel.Application")

            if self.Excel.Application.Workbooks.Count > 0:
                self.set_properties_excel()

            if self.Excel.Application.Workbooks.Count > 0:
                # Проверяем наличие открытой "Data.xlsx"
                for i in range(1, self.Excel.Application.Workbooks.Count + 1):
                    # print(f"{self.file_name} / {self.Excel.Application.Workbooks(i).Name}")
                    if self.file_name == self.Excel.Application.Workbooks(i).Name:
                        # print(f"Найден открытый файл {self.file_name}")
                        self.work_book = self.Excel.Application.Workbooks(i)
                        break

            if self.work_book is None:
                # print(f"Нету открытого файла {self.file_name} / Открываем {self.file_name}")
                self.work_book = self.Excel.Workbooks.Open(os.path.abspath(self.file_name))
                self.set_properties_excel()

            # Выбираем лист "Входные параметры"
            self.sheet_inp = self.work_book.Sheets("Входные параметры")
            # Выбираем лист "Результат парсинга"
            self.sheet_out = self.work_book.Sheets("Результат парсинга")

            return True

        except:
            print(f"Не удалось открыть Excel: {self.file_name};\n"
                  f"Причины:\n"
                  f"1. Нужно выйти из режима редактрирования ячеек.\n"
                  f"2. Файл {self.file_name} не найден.")
            time.sleep(2)
            return self.open_excel()

    def set_properties_excel(self):
        # noinspection PyBroadException
        try:
            self.Excel.Application.DisplayAlerts = False
            self.Excel.Application.ScreenUpdating = True
            self.Excel.Application.visible = True
            return

        except:
            print(f"Не удалось открыть Excel: {self.file_name};\n"
                  f"Причины:\n"
                  f"1. Нужно выйти из режима редактрирования ячеек.\n"
                  f"2. Файл {self.file_name} не найден.")
            time.sleep(2)
            return self.set_properties_excel()

    def excel_close(self):
        self.work_book.Close(True)
        self.Excel.Quit()


# Чтения из Excel (входные параметры)
class ReadExcel(LoadExcel, OutPut, IntPut):
    def __init__(self):
        super().__init__()
        self.row = 2  # Номер строки с которой начнется запись

    def read_excel(self):
        # noinspection PyBroadException
        # try:
        # Массив данных с Data.xlsx лист "Входные параметры"
        array = self.sheet_inp.Range(self.sheet_inp.Cells(2, 1), self.sheet_inp.Cells(12, 11)).Value

        for i in array:
            # Отправная точка
            if i[0] is not None:
                self.inp_city = i[0]
            # Тур операторы
            if i[2] is not None:
                self.inp_tur_operator.append(i[2])
        # Интервал дат вылета
        if array[0][4] is not None and array[0][6] is not None:
            self.inp_date['start'] = array[0][4].strftime('%d/%m/%Y')
            self.inp_date['end'] = array[0][6].strftime('%d/%m/%Y')
        # Ночей от - до
        if array[1][8] is not None and array[1][10] is not None:
            self.inp_night['start'] = int(array[1][8])
            self.inp_night['end'] = int(array[1][10])
        # Опция режима записи в Data.xlsx
        if array[3][6] is not None:
            self.inp_rewriting = True if array[3][6] == 'Перезапись' else False
        # Опция отображения браузера
        if array[4][6] is not None:
            self.inp_show_browser = True if array[4][6] == 'Показать' else False

        # Если метод записи "Продолжать" то считываем названии услуг с Data.xlsx
        if not self.inp_rewriting:
            # Номер строки с которой начнется запись
            self.row = self.sheet_out.Cells(1, 1).CurrentRegion.Rows.Count + 1
            # Массив данных с Data.xlsx лист "Результат парсинга"
            array_2 = self.sheet_out.Range(self.sheet_out.Cells(1, 1), self.sheet_out.Cells(1, 60)).Value
            for i in range(12, len(array_2[0])):
                if array_2[0][i] is not None:
                    self.inp_services.append(array_2[0][i])
                else:
                    # print(f"\nУслуги отеля: {self.inp_services}")
                    break
        else:
            self.sheet_out.Select()
            self.sheet_out.Range(self.Excel.Application.Selection, self.Excel.Application.Selection.End(-4121)).Select()
            self.sheet_out.Range(self.Excel.Application.Selection, self.Excel.Application.Selection.End(-4161)).Select()
            self.Excel.Application.Selection.EntireRow.Delete()
            self.sheet_out.Range('A2').Select()

        print(f"\nОтправная точка: {self.inp_city}"
              f"\nТуроператоры: {self.inp_tur_operator}"
              f"\nИнтервал дат вылета: {self.inp_date}"
              f"\nНочей от-до: {self.inp_night}"
              f"\nМетод записи в Data.xlsx: {array[3][6]}"
              f"\nОтображение браузера: {array[4][6]}\n")
        return True
        # except:
        #     print(f"except read_excel")
        #     return False


# Запись в Excel результата парсинга
class WriteExcel(ReadExcel):
    def __init__(self):
        super().__init__()
        self.count_hotel = 0
        self.count_tur = 0
        self.title_services = []
        self.value_services = []
        self.title_changed = False

    def write_excel(self):
        """
        НОМЕРА СТОЛБЦОВ В Data.xlsx
        hotel_name = 1
        city = 2
        tour_operator = 3
        beach_line = 4
        distance_to_airport = 5
        food = 6
        hotel_rating = 7
        count_stars = 8
        date_start = 9
        date_end = 10
        night = 11
        price = 12
        services = 13
        """
        # Фармируем одномерный массив для записи в Data.xlsx
        data = [self.out_hotel_name[-1],
                self.out_city[-1],
                self.out_tour_operator[-1][-1],
                self.out_beach_line[-1],
                self.out_distance_to_airport[-1],
                self.out_food[-1][-1],
                self.out_hotel_rating[-1],
                self.out_count_stars[-1],
                self.out_date[-1][-1][0],
                self.out_date[-1][-1][1],
                self.out_night[-1][-1],
                self.out_price[-1][-1]
                ]

        try:

            """ //////////////// Записывам услуги отеля /////////////////////////////////////////////////"""
            """
            Так как количесто и виды услуг у отелей разные то с 13 столбца название услуг формируются по предлогаемым услугам отелей
            т.е. например первый отель garden ее услуги - (wifi, СПА, Басейн)
            после столбца 12-Цена будут названии услуг 13-wifi, 14-СПА, 15-Басейн
            у следующего отеля услги - (wifi, СПА, Басейн, места для курении) т.е. на одну услугу больше это (места для курении)
            значит появится еще один столбец 16-места для курении
            теперь это выглядит так 12-Цена, 13-wifi, 14-СПА, 15-Басейн, 16-места для курении
            и в дальнейшем если у отеля будет услуга которой нету в названий услуг то она добавится в конец таблицы
            если у отеля будет услуга wifi то она добавится в 13-wifi если басейн то 15-Басейн и т.д.

            title_services - массив названий услуг в Data.xlsx
            value_services - массив с описаниями услуг в Data.xlsx
            out_services[-1][0] - массив названий услуг текущего отеля
            out_services[-1][1] - массив с описаниями услуг текущего отеля
            """
            # Указываем что массив с названии услуг не изменен
            self.title_changed = False

            # Если данные по услугам отсутствуют, то value_services пустой массив и услуги не записываем
            if not self.out_services[-1]:
                self.value_services = []
            # Если данные по услугам есть
            else:
                # Если Метод записи "Продолжить"
                if not self.inp_rewriting and len(self.title_services) == 0:
                    self.title_services = self.inp_services

                # Если массив названий услуг пустой
                if len(self.title_services) == 0:
                    self.title_services = self.out_services[-1][0]  # Заполняем массив название услуг
                    self.title_changed = True  # Указываем массив названии услуг изменен
                    for i in self.out_services[-1][1]:
                        data.append(i)  # Добовляем в одномерный массив data для записи в Data.xlsx
                # Если массив названий услуг не пустой
                else:
                    # Создаем пустой массив для описаний услуг текущего отеля. Размерность такая же как (title_services -  массив названий услуг в Data.xlsx)
                    self.value_services = ['' for i in range(len(self.title_services))]

                    # Проходим по массиву названий услуг текущего отеля
                    for i in self.out_services[-1][0]:  # i - Название услуги текущего отеля
                        out_index = self.out_services[-1][0].index(i)  # Индекс элемента названий услуг текущего отеля
                        # Если в массиве названий услуг в Data.xlsx ЕСТЬ названий услуг текущего отеля
                        if i in self.title_services:
                            index_1 = self.title_services.index(i)  # Индекс элемента названий услуг Data.xlsx
                            self.value_services[index_1] = self.out_services[-1][1][out_index]  # Описаниями услуг текущего отеля
                        # Если в массиве названий услуг в Data.xlsx НЕТУ названий услуг текущего отеля
                        else:
                            # (Добавляем в массив названий услуг Data.xlsx) название услуги текущего отеля
                            self.title_services.append(i)
                            # (Добавляем в массив с описаниями услуг Data.xlsx) описаниями услуг текущего отеля
                            self.value_services.append(self.out_services[-1][0][out_index])
                            # Указываем массив названии услуг изменен
                            self.title_changed = True

                    for i in self.value_services:
                        data.append(i)  # Добовляем в одномерный массив data, услуги текущего отеля для записи в Data.xlsx

                # Если массив названии услуг изменен
                if self.title_changed:
                    # Обновляем названии услуг Data.xlsx
                    self.sheet_out.Range(self.sheet_out.Cells(1, 13), self.sheet_out.Cells(1, 12 + len(self.title_services))).value = self.title_services
            """//////////////////Конец формирование услуг/////////////////////////////////////////////////////////////////////////////////////"""

            # Записываем в Data.xlsx данные о туре
            self.sheet_out.Range(self.sheet_out.Cells(self.row, 1), self.sheet_out.Cells(self.row, len(data))).value = data
            self.row += 1  # Номер строки в Data.xlsx где производится запись увеличиваем на + 1

        except:
            # Если в excel ячейка будет в режими редактирования то запись прервется и произайдет рекурсия, он будет пытаться записать сного и сного, пока
            # не выйдети из режима редактирования ячейки
            print(f"Не удалось записать данные в Excel: {self.file_name};\n"
                  f"Причины:\n"
                  f"1. Нужно выйти из режима редактрирования ячеек.\n"
                  f"2. Файл {self.file_name} не найден.")
            time.sleep(2)
            return self.write_excel()  # Рекурсия

    def printer(self):
        for i in range(0, len(self.out_hotel_name)):
            print(f"Отель {self.out_hotel_name[i]}")
            print(f"Количество звезд {self.out_count_stars[i]}")
            print(f"До аэропорта {self.out_distance_to_airport[i]}")
            print(f"Пляжная линия {self.out_beach_line[i]}")
            print(f"Сервисы {self.out_services[i]}")
            for j in range(0, len(self.out_price[i])):
                print(f"Цены {self.out_price[i][j]}")
                print(f"Даты {self.out_date[i][j]}")
                print(f"Питание {self.out_food[i][j]}")
            print('\n')


# Запуск WebDriverChrome
class Execute(WriteExcel):
    def __init__(self):
        super().__init__()
        self.driver = None
        self.star_date = None
        self.end_date = None
        self.count_recurs = 0

    def star_driver(self):
        # Устанавливаем опции для webdriverChrome
        options = webdriver.ChromeOptions()
        # Развернуть на весь экран
        options.add_argument("--start-maximized")
        # Утсанавливаем запрет на загрузку изоброжении
        options.add_argument('--blink-settings=imagesEnabled=false')
        options.add_argument("--log-level=3")

        # Если опция Браузер: Скрыть
        if not self.inp_show_browser:
            # Не показываем веб браузер
            options.add_argument('headless')

        # Запускаем webDriverChrome
        self.driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), chrome_options=options)
        return True

    # Добавляем на страницу свою библиотеку js
    def set_library(self):
        try:
            # Считываем скрипты
            lib = open('My_library.js', 'r', encoding='utf-8').read()

            # Устанавливаем в нутри <body><script> последним элемент </body></script>
            time.sleep(5)  # Останавливаем дальнейшее выполнения кода на 5 секунд
            # Внедряем скрирт inc_library(data) в веб страницу
            self.driver.execute_script("""      
    
                lib = function inc_library(data) {
                        if ($('body').length == 0) {
                            return false;
                        }
                            var scr = document.createElement('script');
                            scr.textContent = data;
                            document.body.appendChild(scr);
    
                            scr = document.createElement('script');
                            scr.type = 'text/javascript';
                            scr.src = 'https://code.jquery.com/jquery-3.6.0.min.js';
                            document.head.appendChild(scr);
    
                            return true;
                        }
                var scr = document.createElement('script');
                scr.textContent = lib;
                document.body.appendChild(scr);
            """)
            # Запускаем ранее внедренный скрипт inc_library(data)
            self.js_execute(tr=1, data=f"inc_library({[lib]})")
            return True

        except FileNotFoundError:
            print('Файл My_library.js не найден.')
            return False

    # Для запуска внедренных javaScripts на странице и получения от них ответа
    def js_execute(self, tr=0, sl=0, rt=False, t=0, data=None):
        """
        :param tr: int Количество рекурсией (количество попыток найти элемент)
        :param sl: int Количество секунд ожидать перед выполнения кода
        :param rt: bool Возвращаем значение True если нет то False
        :param t: Если элемент не найден то вернуть: 0 - Данных нет; 1 - 0; 2 - False;
        :param data: Название скрипта
        :return: Возвращает ответ если rt=True
        """
        if sl > 0:
            time.sleep(sl)  # Засыпаем

        # Запускаем javaScript в браузере и получаем результат
        result = self.driver.execute_script(f"return {data}")

        # Если результа False и count_recurs < tr засыпаем на 2 сек. и запускаем рекурсию (рекурсия на случие если элемент не успел появится)
        if not result and self.count_recurs < tr:
            time.sleep(2)
            self.count_recurs += 1  # Увеличиваем счетчик рекурсии на +1
            return self.js_execute(tr=tr, sl=0, rt=rt, t=t, data=data)  # Рекурсия с темеже параметрами

        # Результат False то возвращам по условию значение (Данных нет, 0, False)
        if not result:
            if t == 0:
                result = 'Данных нет'
            if t == 1:
                result = 0
            if t == 2:
                result = False
        # Обнуляем счетчик количество рекурсии
        self.count_recurs = 0

        # Если параметр rt = True то возвращам полученый от скрипта значение
        if rt:
            return result


# Поиск туров
class FindTours(Execute):
    def __init__(self):
        super().__init__()
        self.link = None
        self.page_loaded = None

    def find_tours(self):
        """
        //sletat.ru/search/from-ufa-to-russia-for-tomorrow-nights-22..26-adults-2-kids-zero?datefrom=15/04/2021&dateto=22/04/2021&
        currency=RUB&operators=19,7,20,213,3,54,4,380,9&ticketsincluded=true&hastickets=true&places=true
        """
        # Сопостовляем название городов
        city_dic = dict(Москва='moscow', Хабаровск='khabarovsk', Санкт_Петербург='saint_petersburg', Нижний_Новгород='nizhny_novgorod',
                        Пятигорск='pyatigorsk', Новосибирск='novosibirsk', Екатеринбург='yekaterinburg', Ростов_на_Дону='rostov_on_don',
                        Казань='kazan', Уфа='ufa', Краснодар='krasnodar')

        # Сопостовляем название тур_операторов
        tour_operator_dic = dict(Anex=19, TUI=380, Pegas_Touristik=3, TEZ_TOUR=4, Coral_Travel=6, Biblio_Globus=7, Интурист=9, ICS_Travel_Group=20,
                                 Sunmar=54, Mouzenidis_Travel=213)

        # Формируем строку тур_операторов (19,380,3,4,6,7,9,20,54,213)
        tour_operator_str = ''
        for i in self.inp_tur_operator:
            tour_operator_str += f"{tour_operator_dic[i]},"
        tour_operator_str = re.subn(r'[,]$', '', tour_operator_str)[0]

        # Формируем url ссылку с выбранными параметрами
        self.link = f"https://sletat.ru/search/from-{city_dic[self.inp_city]}" \
                    f"-to-russia-for-tomorrow-nights-{self.inp_night['start']}..{self.inp_night['end']}" \
                    f"-adults-2-kids-zero?datefrom={self.inp_date['start']}&dateto={self.inp_date['end']}&currency=RUB&operators=" \
                    f"{tour_operator_str},&ticketsincluded=true&hastickets=true&places=true"

        # Загружаем сайт по ссылке link
        self.driver.get(self.link)

        # Добавляем на страницу свою библиотеку js
        if not self.set_library():
            return False

        # Указываем что страница не загружена
        self.page_loaded = False

        # Ожидаем загрузки страницы
        # Пока Страница не загружена считываем значение с банера (процент загрузки туров)
        while not self.page_loaded:
            loading = self.js_execute(tr=30, sl=1, rt=True, t=1, data='wait_loading()')

            # Если страница загружена на 100% то прирываем цикл
            if loading == 100:
                print(f"Страница загружена: {loading} %")
                self.page_loaded = True
                break  # Прирываем цикл while
            else:
                print(f"Ожидайте идет загруза страницы: {loading} %")

        return True


# Парсинг страницы
class ParsingPage(FindTours):
    def __init__(self):
        super().__init__()
        self.continuation = True
        self.parsing = True
        self.count_hotel = 0
        self.count_tur = 0
        self.total_tur = 0
        self.start_time = 0
        self.total_time = [0, 0]
        self.time_loaded_page = 0
        self.number_page = 1
        self.page_updated = False
        self.city_and_hotel = []

        keyboard.add_hotkey('Ctrl + P', self.pause_parsing)
        keyboard.add_hotkey('Ctrl + S', self.stop_parsing)

    # События при завершении парсинга
    def close(self):
        print("Для выхода нажмите 'Ctrl + Q'")
        keyboard.wait('Ctrl + Q')
        self.driver.close()
        # self.excel_close()
        print("\nПрограмма завершена")

    # События при приостановки парсинга
    def pause_parsing(self):
        if self.continuation:
            self.continuation = False
        else:
            self.continuation = True
            print(f"Парсинг продалжается.")

    # Для досрочного завершения парсинга
    def stop_parsing(self):
        self.parsing = False

    # Прирывание парсинга по горячим клавишам
    def stop_iterate(self):
        """
        "Ctr + P" - continuation переводит в False / повторное нажатие continuation переводит в True
        "Ctr + S" - parsing переводит в False
        """
        # Приостанавливаем парсинг по горячем клавишам "Ctr + P"
        # Если continuation False
        if not self.continuation:
            print(f"\nПарсинг остановлен.\n")
            # Пока continuation False
            while not self.continuation:
                # Цикл while прервется если сработает "Ctr + P" - continuation переводит в True
                # Таже Цикл while прервется если сработает "Ctr + S" - parsing переводит в False
                time.sleep(3)
                # Если parsing False
                if not self.parsing:
                    # сontinuation переводит в True и цикл завершается
                    self.continuation = True

        # Останавливаем парсинг по горячем клавишам "Ctr + S"
        # Если parsing False
        if not self.parsing:
            print(f"\nПарсинг закончен.\n"
                  f"Всего туров {self.total_tur}; Затраченное время {self.total_time[0]}мин {self.total_time[1]}сек.")
            return False  # Завершаем парсинг

        return True  # Продолжаем парсинг

    # Переход на следующию страницу
    def next_page(self):
        self.number_page += 1  # Номер страницы
        # Если скрипт next_page(number_page) вернет False значит страниц больше нету и парсинг завершится
        if not self.js_execute(tr=1, rt=True, t=2, data=f"next_page({self.number_page})"):
            print(f"\nПарсинг закончен.\n"
                  f"Всего туров {self.total_tur}; Затраченное время {self.total_time[0]}мин {self.total_time[1]}сек.")
            return
        # Если скрипт next_page(number_page) вернет True то запустится цикл while
        while True:
            time.sleep(1)
            # Цикл while прервется если скрипт next_page(number_page) вернет True (скрипт проверит перешли мы на стр. number_page или нет)
            if self.js_execute(tr=1, rt=True, t=2, data=f"number_page({self.number_page})"):
                print(f"Переход на следующию страницу {self.number_page}стр.")
                return self.parsing_page()  # С нова запускаем parsing_page()

    # Обновление страницы
    def page_update(self):
        """
        time_loaded_page - после загрузки страницы записываем время
        number_page - номер страницы
        find_tours() - функция поиска туров
        page_updated - указываем что страница загружена
        """
        # Если это паревая загрузка
        if self.time_loaded_page == 0:
            self.time_loaded_page = datetime.now()

        # Если с момента загрузки страницы прошло 10 мин то обновляем страницу
        if divmod((datetime.now() - self.time_loaded_page).seconds, 60)[0] >= 10:
            print("Обновление страницы")
            self.number_page = 1
            self.find_tours()
            self.time_loaded_page = datetime.now()
            self.page_updated = True

    # Парсинг страницы
    def parsing_page(self):
        """
        Название отеля
        Город
        Туроператор
        Пляжная линия
        Расстояние до аэропорта
        Питание
        Рейтинг отеля
        Количество звезд отеля
        Даты
        Количество ноче
        Цена
        Услуги

        start_time - Время начала парсинга
        count_hotel - Количество отелей на странице
        time_start - Время начало парсинга для однго отеля
        time_end - Время окончания парсинга для отеля
        city_and_hotel - Массив город отель "УФа Garden, Уфа hilten"
        count_tur - Количество туров
        total_time - Общее время затраченное на парсинг
        total_tur - Счетчик по количеству туров
        """

        # Время начала парсинга
        if self.start_time == 0:
            self.start_time = datetime.now()

        # Количество отелей на странице
        self.count_hotel = self.js_execute(tr=10, sl=4, rt=True, t=1, data=f"count_hotel()")

        # Проходим по всем отелям на странице
        for i in range(0, self.count_hotel):
            # Время начало парсинга для однго отеля
            time_start = datetime.now()

            # Город отель
            city_and_hotel = self.js_execute(tr=1, sl=0, rt=True, data=f"city_and_hotel({i})")

            # Если в массиве self.city_and_hotel нету city_and_hotel то пасим этот отель
            if city_and_hotel not in self.city_and_hotel:
                # Массив для хранений названии отелей в каком городе они
                self.city_and_hotel.append(self.js_execute(tr=1, sl=0, rt=True, data=f"city_and_hotel({i})"))

                # Описания отеля
                self.js_execute(tr=5, data=f"show_description({i})")  # Открываем описание
                self.out_hotel_name.append(self.js_execute(tr=1, sl=2, rt=True, data=f"hotel_name({i})"))  # Название отеля
                self.out_city.append(self.js_execute(tr=1, sl=1, rt=True, data=f"city_name({i})"))  # Город
                self.out_hotel_rating.append(self.js_execute(tr=1, rt=True, data=f"hotel_rating({i})"))  # Рейтинг отеля
                self.out_count_stars.append(self.js_execute(tr=1, rt=True, data=f"count_stars({i})"))  # Количество звезд
                self.out_distance_to_airport.append(self.js_execute(tr=1, rt=True, data=f"distance_to_airport({i})"))  # До аэропорта
                self.out_beach_line.append(self.js_execute(tr=1, rt=True, data=f"beach_line({i})"))  # Пляжная линия
                self.out_services.append(self.js_execute(tr=1, rt=True, t=2, data=f"services()"))  # Услуги

                # Описание по турам
                self.js_execute(tr=1, data=f"show_turs({i})")  # Открываем описания по турамё
                self.count_tur = self.js_execute(tr=1, sl=2, rt=True, t=1, data=f"count_tur({i})")  # Количество туров

                # Создаем пустые массивы тур_операторы, цены, даты, питание, ночи
                self.out_tour_operator.append([])
                self.out_price.append([])
                self.out_date.append([])
                self.out_food.append([])
                self.out_night.append([])

                # Проходим по всем турам отеля
                for j in range(0, self.count_tur):
                    # Описание по турам
                    self.out_tour_operator[-1].append(self.js_execute(tr=1, rt=True, data=f"tour_operator({j})"))  # Тур_оператор
                    self.out_price[-1].append(self.js_execute(tr=1, rt=True, data=f"price({j})"))  # Цена
                    self.out_date[-1].append(self.js_execute(tr=1, rt=True, data=f"date({j})"))  # Даты
                    self.out_food[-1].append(self.js_execute(tr=1, rt=True, data=f"food({j})"))  # Питание
                    self.out_night[-1].append(self.js_execute(tr=1, rt=True, data=f"nights({j})"))  # Ночи

                    # Записываем полученные данные в excel
                    self.write_excel()

                # Время окончания парсинга для отеля
                time_end = datetime.now()
                delta = (time_end - time_start).seconds

                # Общее время затраченное
                self.total_time = divmod((time_end - self.start_time).seconds, 60)

                # Счетчик по количеству туров
                self.total_tur += self.count_tur

                print(f"Отель: {self.out_hotel_name[-1]}; Количество туров {self.count_tur}; {delta} сек; "
                      f"Всего туров {self.total_tur}; Общее время {self.total_time[0]}мин {self.total_time[1]}сек.")

            # Прирываем парсинг либо завершаем парсинг
            if not self.stop_iterate():
                return

            # Обновляем страницу каждые 10 мин
            self.page_update()

        # Если есть следующая страница то переходим на следующию если нет завершаем
        self.next_page()
