#!/usr/bin/env python
# CONFIGURATION

# удалять исходный файл
REMOVE_FILE_IN = True

# удалять столбцы
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
    # 'Старая цена',
    'Цена закупки',
    'Себестоимость',
    # 'Остаток',
    'Остаток: Москва',
    'Остаток: Санкт-Петербург',
    'Остаток: Новосибирск',
    # 'Вес',
    'Изображения варианта',
    'Тип цен: Интернет',
    'Тип цен: МОЦ',
    # 'Тип цен: Распродажа',
    'Параметр: Складская (скрытый параметр)',
    # 'Параметр: Распродажа (скрытый параметр)',
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

# вкл/выкл (выполнять перемещение, сворачивание, копирование данных)
ADDITIONAL_ACTIONS = True
# ------------------------------------------------------------------------------
# перености столбец: после столбца
MOVE_COLUMNS = {
    'Артикул': 'Название товара или услуги',
    'Видимость на витрине': 'Артикул',
    'Габариты варианта': 'Габариты',
    'Параметр: Изготовитель': 'Параметр: Бренд',
    'Старая цена': 'Цена продажи',
    'Тип цен: Распродажа': 'Старая цена',
}

# заполнить данные (если пусто) в столбце: из столбца
COPY_DATA_COLUMNS = {
    'Параметр: Бренд': 'Параметр: Изготовитель',
    'Габариты варианта': 'Габариты',
}

# свернуть столбцы
HIDDEN_COLUMNS = (
    'Краткое описание',
    'Габариты варианта',
    'Параметр: Изготовитель',
)
# ------------------------------------------------------------------------------

# разделять картинки по столбцам
IMAGE_SEPARATION = True

# заморозить регион после указанного столбца
FREEZE_REGION = 'Артикул'

# высота шапки (*1.33)
HEADER_HEIGHT = 60

# ширина колонки  с наименованием (*7)
COL_NAME_WIDTH = 66

# ширина остальных колонок (*7)
COL_WIDTH = 15

CLOSECONSOLE = True
EXCEPTION_TRACE = True
