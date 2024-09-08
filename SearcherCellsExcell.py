import os
import sys
import openpyxl
from tqdm import tqdm
from prettytable import PrettyTable
from time import sleep
import textwrap

version = "1.1"
author = "AAIgorevich"


class SCEComands:
    # data text commands
    def __init__(self) -> None:
        self.input_text = textwrap.dedent("""
        Для того чтобы получить список команд,
        введите: 'searcher -help'.""").strip()
        self.stop_text = "Досвидания. Запускайте еще!"
        self.help_text = textwrap.dedent("""
        Доступные команды:
        searcher -stop: Выйти из программы
        searcher -help: Показать этот список команд
        searcher -hi: Приветствие
        searcher -info: Информация о программе
        """).strip()
        self.hi_text = textwrap.dedent("""
        Добро пожаловать в SearcherCellsExcell!
        Эта программа поможет вам быстро находить ячейки
        с определёнными значениями в ваших Excel файлах.
        Начните поиск и упростите свою работу с данными.
        Удачи!
        """).strip()
        self.info_text = textwrap.dedent(f"""
        SearcherCellsExcell program information:
        =======================================
        SearcherCellsExcell or SCE version: {version}
        author: {author}
        link GitHub author: https://github.com/{author}
        """).strip()

    # Остановка и выход из программы
    def command_sce_stop(self):
        print(self.stop_text)
        sleep(0.5)
        sys.exit

    # Вывод всех имеющихся команд для на консоль (помощь)
    def command_sce_help(self):
        print(self.help_text)

    # Приветствие на консоль
    def command_sce_hi(self):
        print(self.hi_text)

    # Информациия о программе
    def command_sce_info(self):
        print(self.info_text)

    # Подсказка для пользователей выводится единожды
    def first_init_command_help(self):
        print(self.input_text)

    # Вызов команд
    def call_comands(self, search_value) -> str:
        if search_value == "searcher -stop":
            self.command_sce_stop()
            return "stop"
        elif search_value == "searcher -help":
            self.command_sce_help()
            return "continue"
        elif search_value == "searcher -hi":
            self.command_sce_hi()
            return "continue"
        elif search_value == "searcher -info":
            self.command_sce_info()
            return "continue"


SCE = SCEComands()

SCE.first_init_command_help()

# Укажите путь к папке с Excel-файлами
folder_path = os.path.abspath(os.curdir)


try:
    while True:
        # Укажите значение, которое нужно найти
        search_value = str(input("Ваше значение: "))

        # Получение результата
        result = SCE.call_comands(search_value)

        # Остановка или продолжение использование программы
        if result == "stop":
            break
        elif result == "continue":
            continue

        def find_cell_by_value(folder_path, search_value: str):
            store_results = []  # Список для хранения результатов поиска

            # Итерируемся по всем файлам в указанной папке
            for filename in tqdm(
                    os.listdir(folder_path),
                    desc="Статус работы: ",
                    colour="#00FF00", position=1, leave=False
                    ):
                if filename.endswith(".xlsx"):
                    file_path = os.path.join(folder_path, filename)
                    # Загружаем книгу Excel
                    workbook = openpyxl.load_workbook(file_path, read_only=True)
                    # Итерируемся по всем листам книги
                    for sheet_name in tqdm(
                            workbook.sheetnames,
                            # Вывод файлов которые были просмотренны
                            desc=f"Просмотр файла {filename}.",
                            colour="#FFFF00",
                            ascii=True, leave=False
                            ):
                        sheet = workbook[sheet_name]
                        # Итерируемся по всем ячейкам в листе
                        for row in sheet.iter_rows():
                            for cell in row:
                                # Проверяем
                                # совпадает ли значение ячейки с искомыми
                                cell_value = str(cell.value)
                                if cell_value == search_value:
                                    # Сохраняем путь к
                                    # файлу, имени листа и адреса ячейки
                                    store_results.append(
                                        [filename,
                                        sheet.title,
                                        cell.coordinate])

            # Возвращаем результаты поиска
            return store_results

        locations = find_cell_by_value(folder_path, search_value)
        table = PrettyTable(["Имя файла", "Название Листа", "Координаты Ячейки"])
        for row in locations:
            table.add_row(row)
        if locations:
            print("\nНайдены совпадения в (.xlsx) файлах с вашем значением: ")
            print(table)
        else:
            print("\nДанное значение не обнаруженно в (.xlsx) файлах.")
        print("|===========================================================|")
except KeyboardInterrupt:
    SCE.command_sce_stop()
