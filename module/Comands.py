import os
import sys
import textwrap
from time import sleep
from module.Env import EnvName as N


# Класс в котором сосредоточенны команды для программы
class SCEComands:
    # data text commands
    def __init__(self) -> None:
        self.SCE_logo = textwrap.dedent(r"""
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
        ╔═══════════════════════════════════════════════════════╗
        ║                                                       ║
        ║ Доступные команды:                                    ║
        ║ ===================================================== ║
        ║ searcher -hi          : Приветствие                   ║
        ║ searcher -stop        : Выйти из программы            ║
        ║ searcher -info        : Информация о программе        ║
        ║ searcher -help        : Показать список команд        ║
        ║ searcher -d config    : Удалить {N.name_config_file}          ║
        ║ searcher -clear       : Очистка выводу консоли        ║
        ║ searcher -clear -true : Очищать всегда консоль - вкл  ║
        ║ searcher -clear -false: Очищать всегда консоль - выкл ║
        ║ searcher -save        : Сохранить в файл, последний   ║
        ║                         выведеный результат поиска.   ║
        ╚═══════════════════════════════════════════════════════╝
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
        """).strip()).format(N.version, N.author, N.author)

    # Вызов команд
    def _call_comands(self, search_value) -> str | None:
        commands: dict = {
            "searcher -stop": (self._command_sce_stop, "stop"),
            "searcher -help": (self._command_sce_help, "continue"),
            "searcher -hi": (self._command_sce_hi, "continue"),
            "searcher -info": (self._command_sce_info, "continue"),
            "searcher -d config": (
                self._command_remove_config_file, "continue"),
            "searcher -save": (lambda: None, "save"),
            "searcher -clear": (
                self._command_cleanup_console_output, "continue"),
            "searcher -clear -true": (lambda: None, True),
            "searcher -clear -false": (lambda: None, False)
        }
        command = commands.get(search_value)
        if command:
            func, result = command
            func()
            return result
        # Если команда не распознана, возвращаем None, чтобы продолжить поиск
        return None

    # Остановка и выход из программы
    def _command_sce_stop(self):
        print(self.stop_text)
        sleep(0.5)
        return sys.exit()

    # Вывод всех имеющихся команд на консоль (помощь)
    def _command_sce_help(self):
        return print(self.help_text)

    # Приветствие на консоль
    def _command_sce_hi(self):
        print(self.SCE_logo)
        sleep(1)
        return print(self.hi_text)

    # Информациия о программе
    def _command_sce_info(self):
        return print(self.info_text)

    # Подсказка для пользователей выводится единожды
    def _first_init_command_help(self):
        print(self.SCE_logo)
        sleep(0.9)
        print(self.hi_text)
        sleep(0.9)
        return print(self.hint_help)

    # Комагда удаления конфиг файла
    def _command_remove_config_file(self) -> None:
        turple_file_config_ini = find_path_ini_file()
        file_config_ini = turple_file_config_ini[0]
        if os.path.exists(file_config_ini):
            os.remove(file_config_ini)
            print("Файл был успешно удален!")
        else:
            print("Невозможно удалить конфиг файл, по причине его отсутствия!")

    # Команда сохранения результата
    def _save_last_result_in_file(self, table_str) -> None:
        print("Процесс сохранения результатов запущен!")
        try:
            with open(N.name_file_save_result, "w") as file:
                file.write(table_str)
                print("Файл успешно сохранен!")
        except Exception as error:
            print("Ошибка сохрания в файл: " + error)

    def _command_cleanup_console_output(self):
        os.system('cls' if os.name == 'nt' else 'clear')


def find_path_ini_file():
    sce_workspace_dir = os.path.abspath(os.curdir)
    file_config_ini = os.path.join(sce_workspace_dir, N.name_config_file)
    return file_config_ini, sce_workspace_dir
