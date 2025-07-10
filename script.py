import json
import os

import pandas as pd


JSON_FILE_PATH = 'response.json'  # путь к вашему ответу
GUIDE_FILE_PATH = 'guide.xlsx'   # путь к файлу с GUID и Наименованиями
USE_FILE = True  # True — читать из файла, False — делать запрос
DATE = '3.06.2020' # Нужная дата


#Получение данных через API или из файла response.json
def load_json():
    # Получение данных из файла response.json
    if USE_FILE:
        with open(JSON_FILE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # Получение данных через API
        import requests
        API_URL_OR_PATH = 'http://localhost:9006/borders/chart-data' #базовый URL
        objects_list = ["vFHlJ5Es", "Gp7yQFaE", "Ab7WxWkw"]  # ваши GUID
        # Создаем строку объектов, объединяя их через '&' и добавляя в URL
        objects_str = '&'.join(objects_list)
        # Формируем полный URL с параметрами
        full_url = f"{API_URL_OR_PATH}?mode=halfhour&date={DATE}&objects={objects_str}"
        # Выполняем HTTP GET-запрос по сформированному URL
        response = requests.get(full_url)
        # Проверяем статус ответа; при ошибке выбрасываем исключение
        response.raise_for_status()
        # Возвращаем распарсенный JSON-ответ
        return response.json()


    # Загружаем GUID и  Название из файла guide.xlsx
def load_guide_mapping():
    # Читаем файл guide.xlsx
    df_guide = pd.read_excel(GUIDE_FILE_PATH, header=None, names=['GUID', 'Name'])
    # Создаем словарь GUID -> Name
    guide_map = pd.Series(df_guide.Name.values, index=df_guide.GUID).to_dict()
    return guide_map


# Обработка JSON-данных и получения списка словарей с данными
def process_data(json_data, name_map):
    # Инициализация пустого списка для хранения обработанных строк
    result_rows = []

    # Получаем список блоков данных из JSON по ключу 'data'
    data_blocks = json_data.get('data', [])
    # Если блоков данных нет, выводим сообщение и возвращаем пустой список
    if not data_blocks:
        print('Данных нет')
        return result_rows

    # Проходим по каждому блоку данных в списке
    for block in data_blocks:
        # Извлекаем GUID (идентификатор) из первого элемента блока
        guid = block[0]
        # Извлекаем наименование из второго элемента блока
        name = block[1]
        # Извлекаем временные метки и данные о посетителях из четвертого элемента блока
        timeseries = block[3]

        # Обрабатываем каждую запись времени в массиве timeseries
        for time_entry in timeseries:
            # Первая часть - время события
            time_str = time_entry[0]
            # Вторая часть - количество посетителей, вошедших за это время
            visitors_in = time_entry[1]
            # Третья часть - количество посетителей, вышедших за это время
            visitors_out = time_entry[2]

            # Получаем имя из словаря name_map по GUID, если есть, иначе используем исходное имя
            name_value = name if not name_map else name_map.get(guid, guid)

            # Добавляем сформированную строку в список result_rows
            result_rows.append({
                'GUID': guid,
                'Наименование': name_value,
                'Время': time_str,
                'Посетителей вошло': visitors_in,
                'Посетителей вышло': visitors_out
            })

    # Возвращаем список всех собранных строк
    return result_rows

# Получение данных через API или из файла response.json
json_data = load_json()

# Загружаем GUID и  Название из файла guide.xlsx
name_map = None
if os.path.exists(GUIDE_FILE_PATH):
    name_map = load_guide_mapping()

# Обработка JSON-данных и получения списка словарей с данными
data_rows = process_data(json_data, name_map)
# Создаем DataFrame из списка словарей для удобной работы с табличными данными
df = pd.DataFrame(data_rows)
# Сохраняем DataFrame в Excel-файл 'Результат.xlsx' без индексов
df.to_excel('Результат.xlsx', index=False, engine='openpyxl')
print('Готово! Результат сохранен в "Результат.xlsx"')