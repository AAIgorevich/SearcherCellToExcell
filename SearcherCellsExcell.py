import os
import sys
import openpyxl

# Укажите путь к папке с Excel-файлами
folder_path = os.path.abspath(os.curdir)

# Подсказка для пользователей
print("Для выхода из программы введите: searcher -stop")

while True:

    # Укажите значение, которое нужно найти
    search_value = str(input("Ваше значение: "))

    if search_value == "searcher -stop":
        print("Досвидания. Запускайте еще!")
        sys.exit
        break

    def find_cell_by_value(folder_path, search_value: str):
        store_results = []  # Список для хранения результатов поиска

        # Итерируемся по всем файлам в указанной папке
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(folder_path, filename)

                # Загружаем книгу Excel
                workbook = openpyxl.load_workbook(file_path, read_only=True)

                # Итерируемся по всем листам книги
                for sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]

                    # Итерируемся по всем ячейкам в листе
                    for row in sheet.iter_rows():
                        for cell in row:
                            # Проверяем, совпадает ли значение ячейки с искомыми
                            cell_value = str(cell.value)
                            if cell_value == search_value:
                                # Сохраняем путь к файлу, имя листа и адрес ячейки
                                store_results.append(f"ИмяФайла: {filename} || НазваниеЛиста: {sheet.title} || КоординатыЯчейки: {cell.coordinate} ||")
                print(" => " + filename)  # Вывод файлов в которые были просмотренны

        # Возвращаем результаты поиска
        return store_results

    locations = find_cell_by_value(folder_path, search_value)
    if locations:
        print("Данные найдены: ")
        for location in locations:
            print(location)
    else:
        print("Ячейка не найдена.")
    print("|================================================================|")
