
# Проект  csv2xlsx

---------------------------------------------------------


Convert csv file to xlsx

## Description

Конвертация .csv файла экспорта InSales в нормальный файл .xlsx

Дополнительные действия:
- Удаляет ненужные колонки
- Перемещает колонки
- Форматирует шапку
- Устанавливает ширину столбцов
- Сворачивает столбцы
- Копирует данные между столбцами (незаполненные)
- Закрепляет область


## Requirements

![Python version](https://img.shields.io/badge/python-3.9%2B-blue)
> Требуется Python 3.9.7+

Установка зависимостей:
```sh
pip3 install -r requirements.txt
```


## Configuration

`./config.py`

```python
# CONFIGURATION

REMOVE_FILE_IN = False

CLOSECONSOLE = True
EXCEPTION_TRACE = True
```
``REMOVE_FILE_IN`` : удалять исходный файл после завершения
``CLOSECONSOLE`` : закрывать консоль
``EXCEPTION_TRACE`` : показывать ошибки


## Usage

### Запуск

перетащить файл `.csv` или `.zip` (содержащий csv) на файл скрипта `csv2xlsx.py`

### Запуск из консоли

```bash
python csv2xlsx.py fullpath\file.csv
```
- аргумент `fullpath\file.csv` - полный путь к обрабатываемому файлу


____

## License

![License](https://img.shields.io/badge/license-MIT-green)
:license:  [MIT](https://github.com/toorusr/csv2xlsx/tree/master/LICENSE)


/*******************************************************
 * Written by Vintets <programmer@vintets.ru>, November 2023
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
 * without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
*******************************************************/

____

:copyright: 2023 by Vint
____
