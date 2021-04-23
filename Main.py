from My_library import ParsingPage


class Web(ParsingPage):
    def __init__(self):
        super().__init__()
        print("\nВременная остановка парсинга 'Ctr + P'"
              "\nПродолжить парсинг 'Ctr + P'"
              "\nПолная остановка парсинга нажмите 'Ctr + S'\n")

    def parsing_site(self):
        # Чтение Data.xlsx
        if not self.read_excel():
            return
        # Запус WebDriverChrome
        if not self.star_driver():
            return
        # Формируем url ссылку с входными параметрами
        if not self.find_tours():
            return
        # Запускае цикл парсинга
        self.parsing_page()
        # Завершаем программу
        self.close()


def main():
    # Инициализируем экземпляр класса Web
    driver = Web()
    # Запускам парсер
    driver.parsing_site()


if __name__ == "__main__":
    main()
