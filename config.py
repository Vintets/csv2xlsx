#!/usr/bin/env python
# CONFIGURATION

# удалять исходный файл
REMOVE_FILE_IN = False

# выполнять удаление, перемещение, сворачивание
ADDITIONAL_ACTIONS = True
# ------------------------------------------------------------------------------
REMOVE_COLUMN = True
REMOVE_COLUMNS = (
    'ID товара',
    # 'Название товара или услуги',
    'Название товара в URL',
    'URL',
    # 'Краткое описание',
    # 'Полное описание',
    # 'Видимость на витрине',
    'Применять скидки',
    'Тег title',
    'Мета-тег keywords',
    'Мета-тег description',
    'Размещение на сайте',
    'Весовой коэффициент',
    'Валюта склада',
    'НДС',
    'Единица измерения',
    # 'Габариты',
    # 'Изображения',
    'Ссылка на видео',
    'ID варианта',
    # 'Артикул',
    # 'Штрих-код',
    'Внешний ID',
    # 'Габариты варианта',
    # 'Цена продажи',
    'Старая цена',
    'Цена закупки',
    'Остаток',
    'Остаток: Москва',
    'Остаток: Санкт-Петербург',
    'Остаток: Новосибирск',
    # 'Вес',
    'Изображения варианта',
    'Тип цен: Интернет',
    'Тип цен: МОЦ',
    'Тип цен: Распродажа',
    'Параметр: Складская (скрытый параметр)',
    'Параметр: Распродажа (скрытый параметр)',
    'Параметр: Пометка удаление (скрытый параметр)',
    'Параметр: Выгружать в интернет-магазин (скрытый параметр)',
    'Параметр: Срок поставки (скрытый параметр)',
    # 'Параметр: Бренд',
    # 'Параметр: Код',
    # 'Параметр: Изготовитель',
    # 'Дополнительное поле: Документация',
    'Дополнительное поле: Видео',
    'Дополнительное поле: ID 1С',
    'Дополнительное поле: Изображение',
    'Дополнительное поле: Специальное',
    # 'Дополнительное поле: Короткое наименование',
    )

# свернуть столбцы
HIDDEN_COLUMNS = (
    'Краткое описание',
    'Габариты варианта',
    'Параметр: Изготовитель',
)
# ------------------------------------------------------------------------------

# заморозить регион после столбца
FREEZE_REGION = 'Артикул'
# высота шапки
HEADER_HEIGHT = 60
# ширина колонки  с наименованием
COL_NAME_WIDTH = 66

CLOSECONSOLE = True
EXCEPTION_TRACE = True
