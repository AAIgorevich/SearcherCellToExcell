import os
import sys
import openpyxl
from tqdm import tqdm
from prettytable import PrettyTable
from time import sleep
import textwrap
import configparser


version = "1.1"
author = "AAIgorevich"
name_config_file = 'path_settings.ini'
name_file_save_result = "saved_result.txt"


# Класс в котором сосредоточенны команды для программы
class SCEComands:
    # data text commands
    def __init__(self) -> None:
        self.SCE_greeting = textwrap.dedent(r"""
        ╔══════════════════════════════╗
        ║        ┌/\───/\┐             ║
        ║        │ SCtE  │             ║
        ║        └──╗─╔──┘             ║
        ║ ╔═══──────╝─╚──────═══╗      ║
        ║ ║ ░ ╔──╗ ┌───┐ ╔──╗ ░ ║      ║
        ║ ╚═══╝  │ │>hi│ │  ╚═══╝      ║
        ║        │ └───┘ │             ║
        ║        └───────┘             ║
        ╚══════════════════════════════╝""").strip()
        self.hint_help = textwrap.dedent("""
        Для того чтобы получить список команд,
        введите: 'searcher -help'.""").strip()
        self.stop_text = "Досвидания. Запускайте еще!"
        self.help_text = textwrap.dedent(f"""
        ╔═══════════════════════════════════════════════╗
        ║                                               ║
        ║ Доступные команды:                            ║
        ║ ============================================= ║
        ║ searcher -hi   : Приветствие                  ║
        ║ searcher -stop : Выйти из программы           ║
        ║ searcher -info : Информация о программе       ║
        ║ searcher -help : Показать этот список команд  ║
        ║ searcher -d -fs: Удалить {name_config_file}    ║
        ║ searcher -save : Сохранить в файл, последний  ║
        ║                   выведеный результат поиска. ║
        ╚═══════════════════════════════════════════════╝
        """).strip()
        self.hi_text = textwrap.dedent("""
        ╔════════════════════════════════════════════════════╗
        ║                                                    ║
        ║ Добро пожаловать в SearcherCellsToExcell!          ║
        ║ ================================================== ║
        ║ Эта программа поможет вам быстро находить ячейки   ║
        ║ с определёнными значениями в ваших Excel файлах.   ║
        ║ Начните поиск и упростите свою работу с данными.   ║
        ║ Удачи!                                             ║
        ║                                                    ║
        ╚════════════════════════════════════════════════════╝
        """).strip()
        self.info_text = (textwrap.dedent("""
        ╔════════════════════════════════════════════════════╗
        ║                                                    ║
        ║ SearcherCellsToExcell program information:         ║
        ║ ================================================== ║
        ║ SearcherCellsToExcell or SCtE version: {}         ║
        ║ author: {}                                ║
        ║ link GitHub author: https://github.com/{} ║
        ║                                                    ║
        ╚════════════════════════════════════════════════════╝
        """).strip()).format(version, author, author)

    # Вызов команд
    def call_comands(self, search_value) -> str | None:
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
        elif search_value == "searcher -d config":
            self.remove_config_file()
            return "continue"
        elif search_value == "searcher -save":
            return "save"
        # Если команда не распознана, возвращаем None, чтобы продолжить поиск
        return None

    # Остановка и выход из программы
    def command_sce_stop(self):
        print(self.stop_text)
        sleep(0.2)
        return sys.exit()

    # Вывод всех имеющихся команд на консоль (помощь)
    def command_sce_help(self):
        return print(self.help_text)

    # Приветствие на консоль
    def command_sce_hi(self):
        return print(self.hi_text)

    # Информациия о программе
    def command_sce_info(self):
        return print(self.info_text)

    # Подсказка для пользователей выводится единожды
    def first_init_command_help(self):
        print(self.SCE_greeting)
        sleep(0.2)
        return print(self.hint_help)

    # Комагда удаления конфиг файла
    def remove_config_file(self) -> None:
        if os.path.exists(file_config_ini):
            os.remove(file_config_ini)
            print("Файд был успешно удален!")
        else:
            print("Невозможно удалить конфиг файл, по причине его отсутствия!")

    # Команда сохранения результата
    def save_last_result_in_file(self, table_str) -> None:
        print("Процесс сохранения результатов запущен!")
        try:
            with open(name_file_save_result, "w") as file:
                file.write(table_str)
                print("Файл успешно сохранен!")
        except Exception as error:
            print("Ошибка сохрания в файл: " + error)


# Класс в котором присутсвуют инструменты для извлечение данных из конфиг файла
# или создания собственного конфига файла если последний отсутсвует
class ParserConfigToListOrCreateNew:

    def __init__(self) -> None:
        self.store_result = []  # Для
        self.files_and_path = {}
        self.config = configparser.ConfigParser()

    def read_or_create_config(self) -> dict:
        if os.path.exists(file_config_ini):
            # Читаем конфиг файл
            self.config.read(file_config_ini)
            # Пробегаем по содержимому файла
            for section in self.config.sections():
                if self.config[section].values():
                    # Извлекаем путь
                    path = self.config[section]["path"].strip("''")
                    # Извлекаем наименование файлов
                    files = self.config[section]["files"].strip('"').split()
                    # Создаем словарь куда помещаем данные
                    self.files_and_path.update({
                                section: {
                                    "path": path,
                                    "files": files
                                    }
                            })
            return self.files_and_path
        else:  # Иначе создание конфигурационного файла
            print("config файл отсутсвует!")
            string_excell_files: str = self.search_excell_files_in_root()
            new_config_file = open(file_config_ini, "w")
            if string_excell_files != "":
                new_config_file.write(textwrap.dedent("""
                [ListGroups]
                    """).strip()
                    )
                new_config_file.write(textwrap.dedent("""
                # "ListGroups.GroupFile" Создался по причине того,
                # что в корневой папке программы присутсвуют файлы,
                # в которых можно осуществить поиск ячеек в Excell файлах.
                #  Если вы не желаете искать в этих файлах указанных в
                # "files", то просто удалите все начиная:
                # от "ListGroups.GroupFile", заканчивая "files"(включая).
                [ListGroups.GroupFile]
                path = {}
                files = {}
                    """).format(
                        sce_workspace_dir,
                        string_excell_files))
            new_config_file.write(textwrap.dedent(
                """
                # Ниже представлен пример.
                # Раскоментируя его убрав "#",
                # Вы можете дополнить его или удалить
                # по собственному разумению.
                # [ListGroups.GroupFile1]
                # path = 'C:\Сюда_напишите_путь_к_файлу'
                # files = example_1.xlsx example_2.xlsx example_3.xlsx
                """).strip())
            new_config_file.close()
            print("Создан новый конфиг файл пожалуйста заполните его!")
            exit()

    def parse_dict_to_list(self) -> list:
        # Получаем готовый словарь
        _file_path_dict: dict = self.read_or_create_config()
        # Проверяем пустой словарь или нет
        if _file_path_dict:
            self.store_result = [
                f"{group['path']}\\{file}"
                for group in _file_path_dict.values()
                for file in group['files']
            ]
        else:
            print("Конфиг файл не заполнен!")
            exit()
        _result = self.store_result
        return _result

    def search_excell_files_in_root(self) -> str:
        list_excell_files: list = []
        # Итерируемся по всем файлам в указанной папке
        for filename in os.listdir(sce_workspace_dir):
            if filename.endswith(".xlsx"):  # ! <= TODO *for old formats
                list_excell_files.append(filename)
        string_excell_files: str = " ".join(list_excell_files)
        return string_excell_files


# Боевой класс.
# В нем описан поиск значений в ячейках excel.
class SCESearchInExcellFiles:

    def __init__(self) -> None:
        PCtL = ParserConfigToListOrCreateNew()
        self.list_path: list = PCtL.parse_dict_to_list()
        self.SCECommands = SCEComands()
        self.store_results: list = []
        self.input_search_value: str = ...
        self.str_found = \
            "\nНайдены совпадения в (.xlsx) файлах с вашем значением: "
        self.str_stroke = \
            "|===========================================================|"
        self.str_not_found = \
            "\nДанное значение не обнаруженно в (.xlsx) файлах."
        self.table_str: str = ""

    def search_in_all_sheets(self, workbook, file_path):
        # Итерируемся по всем листам книги
        for sheet_name in tqdm(
                workbook.sheetnames,
                # Вывод файлов которые были просмотренны
                desc=f"Просмотр файла {file_path}.",
                colour="#FFFF00",
                ascii=True, leave=False
                ):
            sheet = workbook[sheet_name]
            self.search_in_all_cells(sheet, file_path, sheet_name)

    def search_in_all_cells(self, sheet, file_path, sheet_name):
        # Итерируемся по всем ячейкам в листе
        for row in tqdm(
                sheet.iter_rows(),
                desc=f"Имя листа: {sheet_name}",
                ascii=True, leave=False
                ):
            for cell in row:
                # Проверяем
                # совпадает ли значение ячейки с искомыми
                cell_value = str(cell.value)
                if cell_value == self.input_search_value:
                    # Сохраняем путь к
                    # файлу, имени листа и адреса ячейки
                    self.store_results.append(
                        [os.path.basename(file_path),
                            sheet.title,
                            cell.coordinate])

    def search_file_and_load_workbook(self):
        for file_path in self.list_path:
            # Загружаем книгу Excel
            workbook = openpyxl.load_workbook(
                file_path, read_only=True
                )
            self.search_in_all_sheets(workbook, file_path)

    def SCE_start_search_in_excel(self):
        try:
            self.SCECommands.first_init_command_help()
            while True:
                self.input_search_value = str(input("Ваше значение: "))
                result = self.SCECommands.call_comands(self.input_search_value)
                if result == "stop":
                    break
                elif result == "continue":
                    self.table_str = ""
                    continue
                elif result == "save":
                    self.SCECommands.save_last_result_in_file(self.table_str)
                    continue
                elif result is None:
                    pass
                self.search_file_and_load_workbook()
                locations = self.store_results
                table = PrettyTable(
                    ["Имя файла", "Название Листа", "Координаты Ячейки"])
                for row in locations:
                    table.add_row(row)
                self.table_str = table.get_string()
                if locations:
                    print(self.str_found)
                    print(table)
                else:
                    print(self.str_not_found)
                print(self.str_stroke)
                self.store_results.clear()  # Очищаем список
        except KeyboardInterrupt:
            self.SCECommands.command_sce_stop()


if __name__ == "__main__":
    sce_workspace_dir = os.path.abspath(os.curdir)
    file_config_ini = os.path.join(sce_workspace_dir, name_config_file)
    root = SCESearchInExcellFiles()
    root.SCE_start_search_in_excel()
