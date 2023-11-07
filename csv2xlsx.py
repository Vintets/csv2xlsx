#!/usr/bin/env python

import csv
import os
from pathlib import Path
import sys
from time import sleep

import config
import openpyxl as opx
# from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_from_string
# from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


class ArgumentNotPassedError(Exception):
    """Error argument not passed (file csv)."""
    def __init__(self):
        self.msg = f'Не передан файл для обработки!'

    def __str__(self):
        return self.msg


class ArgumentIsFolderError(Exception):
    """Error argument is folder."""
    def __init__(self):
        self.msg = f'Передана директория, а не файл!'

    def __str__(self):
        return self.msg


class FileNotExistError(Exception):
    """Error argument file not exist."""
    def __init__(self, file_in: str):
        self.msg = f"Файл '{file_in}' не существует!"

    def __str__(self):
        return self.msg


class FileNotCSVError(Exception):
    """Error file not CSV."""
    def __init__(self):
        self.msg = f'Передана не CSV файл!'

    def __str__(self):
        return self.msg


class SaveError(Exception):
    """Error save file."""
    def __init__(self):
        self.msg = f'Ошибка сохранения файла!'

    def __str__(self):
        return self.msg


class Excel:
    def __init__(self, file_out: Path) -> None:
        self.file_out = file_out
        self.wb = None
        self.ws = None

    def create(self) -> None:
        self.wb = opx.Workbook()
        self.ws = self.wb.worksheets[0]  # wb.active
        self.ws.title = self.file_out.stem

    def save(self) -> None:
        try:
            self.wb.save(self.file_out)
        except OSError:
            raise SaveError()


def validate_transferred_argument() -> Path:
    try:
        arg = sys.argv[1]
    except IndexError:
        raise ArgumentNotPassedError()  # from None
    file_in = Path(arg)

    if file_in.is_dir():
        raise ArgumentIsFolderError()
    elif not file_in.exists():
        raise FileNotExistError(file_in)
    elif file_in.suffix != '.csv':
        raise FileNotCSVError()
    return file_in


def convert_csv(file_in: Path, excel: Excel) -> None:
    with open(file_in, 'r', encoding='utf-16le') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            # print(row)
            excel.ws.append(row)
            # ws.append([ILLEGAL_CHARACTERS_RE.sub('', row)])
    max_row = excel.ws.max_row
    max_col = excel.ws.max_column
    print(f'Файл .csv содержит\tстолбцов:{max_col} строк:{max_row}')


def main() -> None:
    file_in = validate_transferred_argument()
    print(f"Конвертируем файл '{file_in.name}'")
    file_out = file_in.parent / f'{file_in.stem}.xlsx'
    excel = Excel(file_out)
    excel.create()
    convert_csv(file_in, excel)
    excel.save()
    if config.REMOVE_FILE_IN:
        os.remove(file_in)


def exit_from_program(code: int = 0, close: bool = False) -> None:
    if not close:
        input('\n---------------   END   ---------------')
    else:
        sleep(1.5)
    try:
        sys.exit(code)
    except SystemExit:
        os._exit(code)


if __name__ == '__main__':
    os.system('color 71')
    try:
        main()
    except (ArgumentNotPassedError,
            ArgumentIsFolderError,
            FileNotExistError,
            FileNotCSVError,
            SaveError
            ) as e:
        print(e)
        exit_from_program(code=1, close=config.CLOSECONSOLE)
    except KeyboardInterrupt:
        print('Отмена. Скрипт остановлен.')
        exit_from_program(code=0, close=config.CLOSECONSOLE)
    except Exception as ex:
        print(ex)
        if config.EXCEPTION_TRACE:
            raise ex
        exit_from_program(code=1, close=config.CLOSECONSOLE)
