import os
import sys
import textwrap
from time import sleep
from constants import VersionInfo as VI
from constants import (
    CONTINUE_TEXT,
    DESCRIPTION_WELCOME_TEXT,
    ERROR_FILE_SAVE,
    FILE_SAVED_TEXT,
    FILE_WILL_DELETE_TEXT,
    GOODBYE_TEXT,
    HELP_TEXT,
    IMPOSIBLE_DELETE_CONF_FILE_TEXT,
    LIST_COMMANDS_TEXT,
    LOGO_TEXT,
    PROCESS_SAVE_STARTING,
    SAVE_TEXT,
    SCE_AUTOR_LINK_TEXT,
    SCE_CLEAR_TEXT,
    SCE_CLEAR_FASLSE_TEXT,
    SCE_CLEAR_TRUE_TEXT,
    SCE_D_CONFIG_TEXT,
    SCE_HELP_TEXT,
    SCE_HI_TEXT,
    SCE_INFO_TEXT,
    SCE_SAVE_TEXT,
    SCE_STOP_TEXT,
    STOP_TEXT,
    )


# Класс в котором сосредоточенны команды для программы
class SCEComands:
    # data text commands
    def __init__(self) -> None:
        self.SCE_logo = textwrap.dedent(LOGO_TEXT).strip()
        self.hint_help = textwrap.dedent(HELP_TEXT).strip()
        self.stop_text = GOODBYE_TEXT
        self.help_text = textwrap.dedent(LIST_COMMANDS_TEXT.format(VI.NAME_CONF_FILE_TEXT)).strip()
        self.hi_text = textwrap.dedent(DESCRIPTION_WELCOME_TEXT).strip()
        self.info_text = (textwrap.dedent(SCE_AUTOR_LINK_TEXT).strip()).format(VI.VERSION_TEXT, VI.AUTHOR_TEXT, VI.AUTHOR_TEXT)

    # Вызов команд
    def _call_comands(self, search_value) -> str | None:
        commands: dict = {
            SCE_STOP_TEXT: (self._command_sce_stop, STOP_TEXT),
            SCE_HELP_TEXT: (self._command_sce_help, CONTINUE_TEXT),
            SCE_HI_TEXT: (self._command_sce_hi, CONTINUE_TEXT),
            SCE_INFO_TEXT: (self._command_sce_info, CONTINUE_TEXT),
            SCE_D_CONFIG_TEXT: (
                self._command_remove_config_file, CONTINUE_TEXT),
            SCE_SAVE_TEXT: (lambda: None, SAVE_TEXT),
            SCE_CLEAR_TEXT: (
                self._command_cleanup_console_output, CONTINUE_TEXT),
            SCE_CLEAR_TRUE_TEXT: (lambda: None, True),
            SCE_CLEAR_FASLSE_TEXT: (lambda: None, False)
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
            print(FILE_WILL_DELETE_TEXT)
        else:
            print(IMPOSIBLE_DELETE_CONF_FILE_TEXT)

    # Команда сохранения результата
    def _save_last_result_in_file(self, table_str) -> None:
        print(PROCESS_SAVE_STARTING)
        try:
            with open(VI.NAME_FILE_SAVE_RESULT_TEXT, "w") as file:
                file.write(table_str)
                print(FILE_SAVED_TEXT)
        except Exception as error:
            print(ERROR_FILE_SAVE.format(error))

    def _command_cleanup_console_output(self):
        os.system('cls' if os.name == 'nt' else 'clear')


def find_path_ini_file():
    sce_workspace_dir = os.path.abspath(os.curdir)
    file_config_ini = os.path.join(sce_workspace_dir, VI.NAME_CONF_FILE_TEXT)
    return file_config_ini, sce_workspace_dir
