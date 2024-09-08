import os
import sys
import openpyxl
from tqdm import tqdm
from prettytable import PrettyTable
from time import sleep

version = "1.0"

# Укажите путь к папке с Excel-файлами
folder_path = os.path.abspath(os.curdir)


# Подсказка для пользователей выводится единожды
def searcher_first_help():
    input_text = """Для того чтобы получить список команд,
    введите: 'searcher -help'."""
    return print(input_text)


# Остановка и выход из программы
def searcher_stop():
    print("Досвидания. Запускайте еще!")
    sleep(0.5)
    sys.exit


# Вывод всех имеющихся команд для на консоль (помощь)
def searcher_help():
    help_text = """Доступные команды:
    searcher -stop: Выйти из программы
    searcher -help: Показать этот список команд
    searcher -hi: Приветствие
    searcher -info: Информация о программе
    """
    print(help_text)


# Приветствие на консоль
def searcher_hi():
    hi_text = """
Добро пожаловать в SearcherCellsExcell!
    Эта программа поможет вам быстро находить ячейки
    с определёнными значениями в ваших Excel файлах.
    Начните поиск и упростите свою работу с данными.
    Удачи!
    """
    print(hi_text)


# Информациия о программе
def searcher_info():
    info_text = f"""SearcherCellsExcell program information:
|=======================================|
    SearcherCellsExcell or SCE version: {version}
    author: AAIgorevich
    link GitHub author: https://github.com/AAIgorevich
    """
    print(info_text)


# Вызов команд
def call_comands(search_value) -> str:
    if search_value == "searcher -stop":
        searcher_stop()
        return "stop"
    elif search_value == "searcher -help":
        searcher_help()
        return "continue"
    elif search_value == "searcher -hi":
        searcher_hi()
        return "continue"
    elif search_value == "searcher -info":
        searcher_info()
        return "continue"


searcher_first_help()

while True:
    # Укажите значение, которое нужно найти
    search_value = str(input("Ваше значение: "))

    result = call_comands(search_value)

    # Остановк или продолжение использование программы 
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
        print("\nНайдены совпадения(в .xlsx файлах) с вашем значением: ")
        print(table)
    else:
        print("\nДанное значение не обнаруженно в .xlsx файлах.")
    print("|================================================================|")
