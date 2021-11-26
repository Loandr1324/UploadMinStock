# Author Loik Andrey 7034@balancedv.ru
# TODO:
#  Готово 1. Загружаем файл в DataFrame из определённых колонок
#  2. Убираем из окончания наименования столбцов "пробел" МО
#  3. Дублируем колонку каждого склада
#  4. Сортируем колонки
#  5. Добаляем строку с наименованием "внутренний" "внешний" для каждого склада
#  6. Сохраняем в эксель
#  7. Форматируем для удобства восприятия при необходимости
import pandas as pd
import os

FOLDER = 'Исходные данные'
FILE_NAME = 'Анализ МО по компании.xlsx'

def Run ():
    create_df()

def create_df ():
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """

    df_result = None


    file = FOLDER + "/" + FILE_NAME
    df = read_excel(file)
    # TODO Разобраться как считать формулы из эксель
    df.to_excel('test.xlsx')
    return
    df_index = df.index.values.tolist()
    df.loc['Тип МО'] = 'внешний'
    df = df.reindex(['Тип МО'] + df_index)
    df.set_index(['Код', 'Номенклатура'], inplace=True)
    print (df.loc[(df.index.get_level_values('Код') == 'внешний') & (df.index.get_level_values('Номенклатура') == 'внешний')])
    print (df['01 Кирова МО'].loc[('внешний', 'внешний')])

    for col in df.columns:
        df[col[:-3]] = df[col]
        df[col[:-3]].loc[('внешний', 'внешний')] = ('внутренний')
    print (df.loc[('внешний', 'внешний')])
    # TODO Переименовать колонки (удалить из наименования складов ' МО')
    # TODO Отсортировать колонки для удобства восприятия
    df.to_excel('test.xlsx')
    return

    for filename in file_list: # проходим по каждому элементу списка файлов
        print (filename) # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_excel(filename)
        df_search_header = df.iloc[:15, :2] # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
        # print (df_search_header)
        # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
        mask = df_search_header.replace('.*Номенклатура.*', True, regex=True).eq(True)
        # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
        # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
        f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых колонок, если axis=0, то строк
        # print (df.iloc[:15, :2])
        df = df.iloc[int(f):, :] # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
        df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name # Добавляем в наименование тип данных
        df.iloc[0, 0] = 'Код'
        df.iloc[0, 1] = 'Номенклатура'
        df.columns = df.iloc[0] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
        df = df.iloc[2:, :] # Убираем две строки с верха DF
        df['Номенклатура'] = df['Номенклатура'].str.strip() # Удалить пробелы с обоих концов строки в ячейке
        df.set_index(['Код', 'Номенклатура'], inplace=True) # переносим колонки в индекс, для упрощения дальнейшей работы
        # print(df.iloc[:15, :2]) # Для тестов выводим в консоль 15 строк и два столбца полученного DF
        # Добавляем преобразованный DF в результирующий DF
        df_result = concat_df(df_result, df)
    # Добавляем в результирующий DF по продажам расчётные данные
    if add_name == 'продажи':
        df_result = payment(df_result)
    return df_result

def read_excel (file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    # read_df = pd.read_excel(file_name, header=0, usecols= 'A,B,D,G,J,M,P,S,V,Y', engine='pyxlsb')
    # read_df = pd.read_excel(file_name, header=0, usecols= 'A,B,D,G,J,M,P,S,V,Y', engine='openpyxl')
    print ('Попытка загрузки файла:'+file_name)
    try:
        df = read_df
        return (df)
    except KeyError as Error:
        print (Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix (file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            df = read_df
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')

def bug_fix (file_name):
    """
    Переименовываем не корректное имя файла в архиве excel
    :param file_name: Имя excel файла
    """
    import shutil
    from zipfile import ZipFile

    # Создаем временную папку
    tmp_folder = '/temp/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку и удаляем excel
    with ZipFile(file_name) as excel_container:
        excel_container.extractall(tmp_folder)
    os.remove(file_name)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive(f'{FOLDER}/correct_file', 'zip', tmp_folder)
    os.rename(f'{FOLDER}/correct_file.zip', file_name)

if __name__ == '__main__':
    Run()
    # Run1()


