#!/usr/bin/env python

import csv
import os
from pathlib import Path
import sys
from time import sleep
from zipfile import BadZipFile, is_zipfile, ZipFile

import config
import openpyxl as opx
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, BORDER_THIN, Side
from openpyxl.utils import get_column_letter
# from openpyxl.utils.cell import column_index_from_string, coordinate_from_string


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
        self.msg = f'Передан не CSV файл!'

    def __str__(self):
        return self.msg


class FileNotZIPError(Exception):
    """Error file not ZIP."""
    def __init__(self):
        self.msg = f'Неверный формат ZIP файла!'

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

    def stylization(self) -> None:
        self._stylization_header()

        # ширина колонок
        for col in range(1, self.ws.max_column + 1):
            self.ws.column_dimensions[get_column_letter(col)].width = config.COL_NAME_WIDTH

        self.set_width_col_name()

    def _stylization_header(self) -> None:
        self.ws.auto_filter.ref = self.ws.dimensions
        self.ws.row_dimensions[1].height = config.HEADER_HEIGHT
        hdr_al, hdr_fill, thin_border, hdr_font = self._get_styles_header()

        for row_cells in self.ws.iter_rows(min_row=1, max_row=1):
            for cell in row_cells:
                cell.alignment = hdr_al
                cell.fill = hdr_fill
                cell.border = thin_border
                cell.font = hdr_font

    def _get_styles_header(self) -> tuple[Alignment, PatternFill, Border, Font]:
        hdr_al = Alignment(
                          horizontal='general',
                          vertical='top',
                          text_rotation=0,
                          wrap_text=True,
                          shrink_to_fit=False,
                          indent=0
                          )
        hdr_fill = PatternFill(
                              fill_type='solid',
                              start_color='B7DEE8',
                              end_color='B7DEE8'
                              )
        thin_border = Border(
                             left=Side(border_style=BORDER_THIN, color='7F7F7F'),
                             right=Side(border_style=BORDER_THIN, color='7F7F7F'),
                             top=Side(border_style=BORDER_THIN, color='7F7F7F'),
                             bottom=Side(border_style=BORDER_THIN, color='7F7F7F')
                             )
        hdr_font = Font(
                       bold=True,
                       color='1F497D'
                       )
        return (hdr_al, hdr_fill, thin_border, hdr_font)

    def get_header_text(self) -> list[str]:
        header_text = []
        for row_cells in self.ws.iter_rows(min_row=1, max_row=1):
            for cell in row_cells:
                header_text.append(cell.value)
        return header_text

    def set_width_col_name(self):
        if config.COL_NAME_WIDTH <= 0:
            return
        header_text = self.get_header_text()
        try:
            col_name_idx = header_text.index('Название товара или услуги') + 1
        except ValueError:
            return
        self.ws.column_dimensions[get_column_letter(col_name_idx)].width = config.COL_NAME_WIDTH

    def remove_columns(self) -> None:
        if not (config.ADDITIONAL_ACTIONS and config.REMOVE_COLUMN and config.REMOVE_COLUMNS):
            return
        for row_cells in self.ws.iter_rows(min_row=1, max_row=1):
            for cell in reversed(row_cells):
                if cell.value in config.REMOVE_COLUMNS:
                    self.ws.delete_cols(cell.column)

    def hidden_columns(self) -> None:
        if not (config.ADDITIONAL_ACTIONS and config.HIDDEN_COLUMNS):
            return
        for row_cells in self.ws.iter_rows(min_row=1, max_row=1):
            for cell in reversed(row_cells):
                if cell.value in config.HIDDEN_COLUMNS:
                    col_letter = get_column_letter(cell.column)
                    self.ws.column_dimensions.group(col_letter, col_letter, outline_level=1, hidden=True)

    def freeze_region(self) -> None:
        if not config.FREEZE_REGION:
            return
        header_text = self.get_header_text()
        try:
            freeze_idx = header_text.index(config.FREEZE_REGION) + 2
        except ValueError:
            return
        self.ws.freeze_panes = self.ws[f'{get_column_letter(freeze_idx)}2']

    def save(self) -> None:
        try:
            self.wb.save(self.file_out)
        except OSError:
            raise SaveError()


def get_transferred_argument() -> str:
    try:
        arg = sys.argv[1]
    except IndexError:
        raise ArgumentNotPassedError()  # from None
    return arg


def validate_transferred_argument(arg: str) -> Path:
    file_in = Path(arg)
    if file_in.is_dir():
        raise ArgumentIsFolderError()
    elif not file_in.exists():
        raise FileNotExistError(file_in)

    if file_in.suffix == '.zip':
        file_in = unzip_file(file_in)

    if file_in.suffix != '.csv':
        raise FileNotCSVError()
    return file_in


def unzip_file(file_zip: Path) -> Path:
    if not is_zipfile(file_zip):
        raise FileNotZIPError()

    try:
        with ZipFile(file_zip, 'r') as myzip:
            # myzip.printdir()
            namelist = myzip.namelist()[0]
            unzipped_name = myzip.extract(namelist)
    except BadZipFile:
        raise FileNotZIPError()

    # rename unzipped file
    unzipped_name = Path(unzipped_name)
    correct_unzipped_name = file_zip.parent / file_zip.stem
    os.rename(unzipped_name, correct_unzipped_name)
    # print(unzipped_name.parts)
    # print(correct_unzipped_name.parts)

    remove_file(file_zip)
    return correct_unzipped_name


def remove_file(filename):
    if config.REMOVE_FILE_IN:
        os.remove(filename)


def convert_csv(file_in: Path, excel: Excel) -> None:
    with open(file_in, 'r', encoding='utf-16') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            excel.ws.append(row)
    max_row = excel.ws.max_row
    max_col = excel.ws.max_column
    print(f'Файл .csv содержит\tстолбцов:{max_col} строк:{max_row}')


def main() -> None:
    arg = get_transferred_argument()
    file_in = validate_transferred_argument(arg)
    print(f"Конвертируем файл '{file_in.name}'")

    file_out = file_in.parent / f'{file_in.stem}.xlsx'
    excel = Excel(file_out)
    excel.create()
    convert_csv(file_in, excel)

    excel.remove_columns()
    excel.stylization()
    excel.hidden_columns()
    excel.freeze_region()
    excel.save()

    remove_file(file_in)


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
            SaveError,
            FileNotZIPError
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
