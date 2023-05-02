import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import ASSORTMENT_FILE_PATH, STOCKS_FILE_PATH
from services import merge_json_files

# При изменении этой переменной удалите файл token.json
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def authorize():
    """
    Проверить, авторизован ли пользователь. Если не авторизован, перейти на страницу авторизации.
    """
    creds = None

    # Файл token.json хранит токены доступа пользователя
    # Он создаётся автоматически после первой удачной авторизации
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # Если отсутствуют валидные токены доступа, происходит процедура входа в Google аккаунт
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Сохранение токенов доступа
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    return service


def create(title: str) -> str | HttpError:
    """
    Создать Google-таблицу. При успешном создании возвращает id таблицы.
    :param title:
    :return:
    """
    try:
        spreadsheet = {
            "properties": {
                "title": title,
                "locale": "en_US"
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get('spreadsheetId')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def create_tabs(spreadsheet_id: str, tab_name: str, row_amount: int):
    """
    Создать листы в Google-таблице, названные по категориям товаров.
    :param row_amount:
    :param spreadsheet_id:
    :param tab_name:
    :return:
    """
    try:
        body = {
            "requests": [{
                "addSheet": {
                    "properties": {
                        "title": tab_name,
                        "gridProperties": {
                            "columnCount": 5,
                            "rowCount": row_amount + 1
                        },
                    }
                }
            }]
        }
        result = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def adjust_columns(spreadsheet_id: str, tab_id: int, row_amount: int):
    """
    Настроить ширину и высоту ячеек, а также настроить перенос слов внутри ячеек
    и добавить числовой формат для ячеек с ценами
    :param row_amount:
    :param spreadsheet_id:
    :param tab_id:
    :return:
    """
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": tab_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 1
                    },
                    "properties": {
                        "pixelSize": 100
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": tab_id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,
                        "endIndex": 2
                    },
                    "properties": {
                        "pixelSize": 200
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": tab_id,
                        "dimension": "ROWS",
                        "startIndex": 1,
                        "endIndex": row_amount + 1
                    },
                    "properties": {
                        "pixelSize": 200
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": tab_id
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP",
                        }
                    },
                    "fields": "userEnteredFormat.wrapStrategy",
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": tab_id,
                        "startColumnIndex": 2
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER",
                                "pattern": "0.0#"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            }
        ]
    }

    try:
        result = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def group_consumables(spreadsheet_id: str, tab_id: int, group_id_list: list[tuple]):
    """
    Сгруппировать товары категории расходных материалов в соответствующей вкладке таблицы
    :param spreadsheet_id:
    :param tab_id:
    :param group_id_list:
    :return:
    """
    body = {
        "requests": []
    }

    body_after = {
        "requests": []
    }

    for pair in group_id_list:
        body["requests"].append(
            {
                "addDimensionGroup": {
                    "range": {
                        "dimension": "ROWS",
                        "sheetId": tab_id,
                        "startIndex": pair[0],
                        "endIndex": pair[1]
                    }
                }
            }
        )
        body_after["requests"].append(
            {
                "updateDimensionGroup": {
                    "dimensionGroup": {
                        "range": {
                            "dimension": "ROWS",
                            "sheetId": tab_id,
                            "startIndex": pair[0],
                            "endIndex": pair[1]
                        },
                        "collapsed": True,
                        "depth": 1
                    },
                    "fields": "collapsed"
                }
            }
        )

    try:
        result = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        result_2 = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body_after
        ).execute()
        return result, result_2
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def count_consumables(json_file: dict) -> int:
    """
    Сосчитать количество строк для таблицы с расходными материалами.
    :param json_file:
    :return:
    """
    count = 0
    for subcategory in json_file["Расходные материалы"]:
        count += len(json_file["Расходные материалы"][subcategory]) + 1
    return count


def delete_first_tab(spreadsheet_id: str):
    """
    Удалить первый лист в таблице, т.к. он по умолчанию создаётся пустым.
    :param spreadsheet_id:
    :return:
    """
    body = {
        "requests": [
            {
                "deleteSheet": {
                    "sheetId": 0
                }
            }
        ]
    }
    try:
        result = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def transfer_to_sheets(spreadsheet_id: str, json_file: dict):
    """
    Перенести данные о товарах из json файла в гугл таблицы, выполнив
    предварительную настройку таблиц.
    :param spreadsheet_id:
    :param json_file:
    :return:
    """
    data = []
    consumables_number = count_consumables(json_file)
    try:
        for category in json_file:
            # Создание листов по категориям
            result = create_tabs(spreadsheet_id, category,
                                 len(json_file[category]) if category != "Расходные материалы" else consumables_number)
            tab_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]

            # Регулировка высоты и ширины столбцов
            adjust_columns(spreadsheet_id, tab_id,
                           len(json_file[category]) if category != "Расходные материалы" else consumables_number)

            # Добавление заголовков столбцов
            data.append(
                {
                    "range": f"{category}!A1:F1",
                    "values": [["Наименование", "Изображение", "Цена: розница", "Цена: от 5 т.р.", "Цена: от 15 т.р.",
                                "Цена: от 100 т.р."]]
                },
            )

            # Добавление товаров всех категорий, кроме расходных материалов
            if category != "Расходные материалы":
                data.append(
                    {
                        "range": f"{category}!A2:F{len(json_file[category]) + 1}",
                        "values": [[product["name"], f'=IMAGE("{product["image"]}")', product["retailPrice"] / 100,
                                    product["priceFrom5k"] / 100, product["priceFrom15k"] / 100,
                                    product["priceFrom100k"] / 100] for product in json_file[category]]
                    }
                )
            else:
                # Товары категории расходных материалов имеют немного другую структуру,
                # а именно распределены по подкатегориям, из-за чего
                # их обработка происходит иначе

                start_idx = 2  # вспомогательная переменная для правильной расстановки товаров по строкам таблицы
                group_ids_list = []  # вспомогательный список, хранящий номера первой и последней строки товара,
                # и эти номера необходимы для создания раскрывающихся групп

                for subcategory in json_file[category]:
                    group_ids_list.append((start_idx, len(json_file[category][subcategory]) + start_idx))
                    data.append(
                        {
                            "range": f"{category}!A{start_idx}:A{start_idx}",
                            "values": [[subcategory]]
                        }
                    )
                    data.append(
                        {
                            "range":
                                f"{category}!A{start_idx + 1}:F{len(json_file[category][subcategory]) + start_idx}",
                            "values": [[product["name"], f'=IMAGE("{product["image"]}")', product["retailPrice"] / 100,
                                        product["priceFrom5k"] / 100, product["priceFrom15k"] / 100,
                                        product["priceFrom100k"] / 100] for product in json_file[category][subcategory]]
                        }
                    )
                    start_idx += len(json_file[category][subcategory]) + 1

                # Группировка товаров категории расходных материалов
                group_consumables(spreadsheet_id, tab_id, group_ids_list)

        # Формирование тела запроса из тела запроса и типа ячеек
        # При параметре USER_ENTERED Google таблица попытается интерпретировать содержимое ячейки как функцию
        # Например, при тексте =IMAGE() будет вставлена картинка
        body = {
            'valueInputOption': "USER_ENTERED",
            'data': data
        }
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        print(f"{(result.get('totalUpdatedCells'))} ячеек обновлено.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    products = merge_json_files(ASSORTMENT_FILE_PATH, STOCKS_FILE_PATH)
    service = authorize()
    sheet_id = create("LaOne")
    transfer_to_sheets(sheet_id, products)
    delete_first_tab(sheet_id)
