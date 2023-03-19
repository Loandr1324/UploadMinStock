# Author Loik Andrey 7034@balancedv.ru
# TODO:
#  Готово 1. Загружаем файл в DataFrame из определённых колонок
#  2. Убираем из окончания наименования столбцов "пробел" МО
#  3. Дублируем колонку каждого склада
#  4. Сортируем колонки
#  5. Добавляем строку с наименованием "внутренний" "внешний" для каждого склада
#  6. Сохраняем в эксель
#  7. Форматируем для удобства восприятия при необходимости
import pandas as pd
import os

FILE_NAME = 'Анализ МО по компании.xlsx'
NEW_FILE_NAME = 'Загрузка МО по компании.xlsx'


def run():
    n = input('Выберите вариант работы. Введите номер варианта.\n1 - По колонке "...МО"\n2 - По колонке "... МО расчет"'
                  '\nq - Выход\nВариант: ')
    df_result = None
    if n == '1':
        df_result = create_df_mo(FILE_NAME)  # Загружаем и подготавливаем DataFrame
    elif n == '2':
        df_result = create_df_mo_calc(FILE_NAME)
    elif n == 'q':
        exit()
    else:
        print('Вы ввели не корректный вариант.')
        run()
    df_write_xlsx(df_result)  # Записываем в эксель


def create_df_mo(file):
    """
    :param file: Загружаем в DataFrame файлы из file_list
    :return: df_result Дата фрэйм с данными из файлов
    """
    print('Читаем файл:' + file)
    df = pd.read_excel(file, header=0, engine='openpyxl')

    # Добавляем строку со значениями "Внешний" и перемещаем её в начало DataFrame
    df_index = df.index.values.tolist()
    df.loc[''] = 'Внешний'
    df = df.reindex([''] + df_index)
    df.set_index(['Код', 'Номенклатура'], inplace=True)
    df['05 Павловский'] = df['05 Павловский МО']
    df = df.filter(regex='МО$')

    # Копируем колонки и прописываем в них значение "Внутренний" в первую стоку
    for col in df.columns:
        df[col[:-3]] = df[col]
        df[col[:-3]].loc[('Внешний', 'Внешний')] = ('Внутренний')

    # Переименовываем название колонок индексов для корректной записи в эксель
    df.index.names = ['Номенклатура', '']

    # Сортируем колонки для удобства восприятия
    df = sort_df_mo(df)

    # Переименовываем колонки (удаляем из наименования складов " МО") для корректной записи в эксель
    df.columns = df.columns.str.replace(' МО', '')

    return df


def create_df_mo_calc(file):
    """
    :param file: Загружаем в DataFrame файлы из file_list
    :return: df_result Дата фрэйм с данными из файлов
    """
    print('Читаем файл:' + file)
    df = pd.read_excel(file, header=0, engine='openpyxl')

    # Добавляем строку со значениями "Внешний" и перемещаем её в начало DataFrame
    df_index = df.index.values.tolist()
    df.loc[''] = 'Внешний'
    df = df.reindex([''] + df_index)
    df.set_index(['Код', 'Номенклатура'], inplace=True)
    df = df.filter(regex='МО расчёт$')

    # Копируем колонки и прописываем в них значение "Внутренний" в первую стоку
    for col in df.columns:
        df[col[:-10]] = df[col]
        df[col[:-10]].loc[('Внешний', 'Внешний')] = ('Внутренний')

    # Переименовываем название колонок индексов для корректной записи в эксель
    df.index.names = ['Номенклатура', '']

    # Сортируем колонки для удобства восприятия
    df = sort_df_mo_calc(df)

    # Переименовываем колонки (удаляем из наименования складов " МО") для корректной записи в эксель
    df.columns = df.columns.str.replace(' МО расчёт', '')

    return df


def sort_df_mo(df):
    sort_list = ['01 Кирова', '01 Кирова МО',
                 '02 Автолюбитель', '02 Автолюбитель МО',
                 '03 Интер', '03 Интер МО',
                 '04 Победа', '04 Победа МО',
                 '08 Центр', '08 Центр МО',
                 '09 Вокзалка', '09 Вокзалка МО',
                 '05 Павловский', '05 Павловский МО',
                 'Компания MaCar', 'Компания MaCar МО']
    df = df[sort_list]
    return df


def sort_df_mo_calc(df):
    sort_list = ['01 Кирова', '01 Кирова МО расчёт',
                 '02 Автолюбитель', '02 Автолюбитель МО расчёт',
                 '03 Интер', '03 Интер МО расчёт',
                 '04 Победа', '04 Победа МО расчёт',
                 '08 Центр', '08 Центр МО расчёт',
                 '09 Вокзалка', '09 Вокзалка МО расчёт',
                 '05 Павловский', '05 Павловский МО расчёт',
                 'Компания MaCar', 'Компания MaCar МО расчёт']
    df = df[sort_list]
    return df


def df_write_xlsx(df):
    # Сохраняем в переменные значения конечных строк и столбцов
    row_end, col_end = len(df), len(df.columns)
    row_end_str, col_end_str = str(row_end), str(col_end)


    # Изменяем встроенный формат заголовков и индексов pandas
    format_header_css = "font-size: 9px; font-family: Arial; text-align: center; vertical-align: top; " \
                        "white-space: normal; font-weight: bold;" \
                        "background-color: #F4ECC5; border: 1px solid #CCC085;"
    df_style = df.style.apply_index(lambda x: [format_header_css for _ in x], axis="columns")
    name_format_css = "font-size: 11px; font-family: Arial; text-align: left; vertical-align: top; " \
                      "white-space: normal; font-weight: normal; border: 1px solid #CCC085;"
    df_style = df_style.apply_index(lambda x: [name_format_css for _ in x], axis="index")


    # Создаём эксель и сохраняем данные
    name_file = NEW_FILE_NAME
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    with pd.ExcelWriter(name_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        df_style.to_excel(writer, sheet_name=sheet_name)
        wks1 = writer.sheets[sheet_name]  # Сохраняем в переменную вкладку для форматирования

        # Получаем словари форматов для эксель
        header_format, con_format, border_storage_format_left, border_storage_format_right, \
        name_format, MO_format, data_format = format_custom(workbook)

        # Форматируем таблицу
        wks1.set_default_row(12)
        wks1.set_row(0, 40, None)
        wks1.set_row(1, 10, None)
        wks1.set_column('A:A', 12, None)
        wks1.set_column('B:B', 32, None)
        wks1.set_column('C:R', 10, data_format)

        # Делаем жирным рамку между складами
        i = 2
        while i < col_end + 2:
            wks1.set_column(i, i, None, border_storage_format_left)
            wks1.set_column(i + 1, i + 1, None, border_storage_format_right)
            i += 2

        # Объединяем ячейки
        wks1.merge_range(0, 0, 1, 1, 'Номенклатура', header_format)

        # Добавляем фильтр в первую колонку
        wks1.autofilter(1, 0, row_end + 1, col_end + 1)
    return


def format_custom(workbook):
    header_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '7',
        'align': 'center',
        'valign': 'top',
        'text_wrap': True,
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    })

    border_storage_format_left = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'left': 2,
        'left_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'right': True,
        'right_color': '#CCC085',
    })
    border_storage_format_right = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })

    name_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'left',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })

    MO_format = workbook.add_format({
        'num_format': '# ### ##0.00;;',
        'bold': True,
        'font_name': 'Arial',
        'font_size': '8',
        'font_color': '#FF0000',
        # 'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    data_format = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    con_format = workbook.add_format({
        'bg_color': '#FED69C',
    })

    return header_format, con_format, border_storage_format_left, border_storage_format_right, \
           name_format, MO_format, data_format


if __name__ == '__main__':
    run()
