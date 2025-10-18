class VersionInfo:
        VERSION_TEXT = "2.0"
        AUTHOR_TEXT = "AAIgorevich"
        NAME_CONF_FILE_TEXT = 'confpath.ini'
        NAME_FILE_SAVE_RESULT_TEXT = "saved_result.txt"


SCE_STOP_TEXT = "searcher -stop"
SCE_HELP_TEXT = "searcher -help"
SCE_HI_TEXT = "searcher -hi"
SCE_INFO_TEXT = "searcher -info"
SCE_D_CONFIG_TEXT = "searcher -d config"
SCE_SAVE_TEXT = "searcher -save"
SCE_CLEAR_TEXT = "searcher -clear"
SCE_CLEAR_TRUE_TEXT = "searcher -clear -true"
SCE_CLEAR_FASLSE_TEXT = "searcher -clear -false"
STOP_TEXT = "stop"
SAVE_TEXT = "save"
CONTINUE_TEXT = "continue"
PATH_TEXT = "path"
FILES_TEXT = "files"


CONF_FILE_IS_NOT_EXIST = "config файл отсутсвует!"
ERROR_FILE_SAVE = "Ошибка сохрания в файл: {}"
FILE_SAVED_TEXT = "Файл успешно сохранен!"
PROCESS_SAVE_STARTING = "Процесс сохранения результатов запущен!"
IMPOSIBLE_DELETE_CONF_FILE_TEXT = "Невозможно удалить конфиг файл, по причине его отсутствия!"
FILE_WILL_DELETE_TEXT = "Файл был успешно удален!"
HELP_TEXT = """
        Для того чтобы получить список команд,
        введите: 'searcher -help'."""
GOODBYE_TEXT = "Досвидания. Запускайте еще!"
CREATED_NEW_CONF_FILL_OUT_TEXT = "Создан новый конфиг файл пожалуйста заполните его!"
CONF_FILE_NOT_FILL_OUT = "Конфиг файл не заполнен!"
XL_FORMAT_TEXT = ".xlsx"
FIND_SELL_LIKE_U_XL = "\nНайдены совпадения в (.xlsx) файлах с вашем значением: "
BORDER_LINE_TEXT = "|===========================================================|"
VALUE_NOT_FIND_TEXT = "\nДанное значение не обнаруженно в (.xlsx) файлах."
READING_FILE_TEXT = "Просмотр файла {}."
COLOUR_TEXT = "#FFFF00"
TYPE_CHART_TEXT = "Chart"
SKIP_LIST_TEXT = "Пропускаем лист: {}"
U_VALUE_TEXT = "Ваше значение: "


LOGO_TEXT = """
        ╔══════════════════════════════╗
        ║        ┌/\───/\┐             ║
        ║        │ SCtE  │             ║
        ║        └──╗─╔──┘             ║
        ║ ╔═══──────╝─╚──────═══╗      ║
        ║ ║ ░ ╔──╗ ┌───┐ ╔──╗ ░ ║      ║
        ║ ╚═══╝  │ │>hi│ │  ╚═══╝      ║
        ║        │ └───┘ │             ║
        ║        └───────┘             ║
        ╚══════════════════════════════╝"""
LIST_COMMANDS_TEXT = """
        ╔═══════════════════════════════════════════════════════╗
        ║                                                       ║
        ║ Доступные команды:                                    ║
        ║ ===================================================== ║
        ║ searcher -hi          : Приветствие                   ║
        ║ searcher -stop        : Выйти из программы            ║
        ║ searcher -info        : Информация о программе        ║
        ║ searcher -help        : Показать список команд        ║
        ║ searcher -d config    : Удалить {}          ║
        ║ searcher -clear       : Очистка выводу консоли        ║
        ║ searcher -clear -true : Очищать всегда консоль - вкл  ║
        ║ searcher -clear -false: Очищать всегда консоль - выкл ║
        ║ searcher -save        : Сохранить в файл, последний   ║
        ║                         выведеный результат поиска.   ║
        ╚═══════════════════════════════════════════════════════╝
        """
DESCRIPTION_WELCOME_TEXT = """
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
        """
SCE_AUTOR_LINK_TEXT = """
        ╔════════════════════════════════════════════════════╗
        ║                                                    ║
        ║ SearcherCellsToExcell program information:         ║
        ║ ================================================== ║
        ║ SearcherCellsToExcell or SCtE version: {}         ║
        ║ author: {}                                ║
        ║ link GitHub author: https://github.com/{} ║
        ║                                                    ║
        ╚════════════════════════════════════════════════════╝
        """
DESCRIPTION_IN_CONF_TEXT = """
                # "ListGroups.GroupFile" Создался по причине того,
                # что в корневой папке программы присутсвуют файлы,
                # в которых можно осуществить поиск ячеек в Excell файлах.
                #  Если вы не желаете искать в этих файлах указанных в
                # "files", то просто удалите все начиная:
                # от "ListGroups.GroupFile", заканчивая "files"(включая).
                    """
FORMAT_CONF_TEXT = """
                [ListGroups]
                [ListGroups.GroupFile]
                path = {}
                files = {}
                    """
HELP_IN_CONF_TEXT = """
                # Ниже представлен пример.
                # Раскоментируя его убрав "#",
                # Вы можете дополнить его или удалить
                # по собственному разумению.
                # [ListGroups.GroupFile1]
                # path = 'C:\Сюда_напишите_путь_к_файлу'
                # files = example_1.xlsx example_2.xlsx example_3.xlsx
                """


SAMPLE_NAME_COLUMS: list = [
                "Имя файла",
                "Название Листа",
                "Координаты Ячейки"
                ]