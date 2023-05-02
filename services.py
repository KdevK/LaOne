import json


def merge_json_files(assortment_path: str, stocks_path: str) -> dict:
    """
    Объединить два json файла, составив из них словарь следующего вида:
    {
        "Брови": [...],
        "Инструменты": [...],
        ...
        "Расходные материалы": {
            "Аппликаторы для губ": [...],
            "Камни для клея": [...],
            ...
        }
        ...
        "Ресницы_Barbara": [...],
        "Ресницы_Enigma": [...],
        ...
    }
    :param assortment_path:
    :param stocks_path:
    :return:
    """
    with open(assortment_path, encoding='utf-8') as assortment, \
            open(stocks_path, encoding='utf-8') as stocks:
        assortment_json = json.load(assortment)
        stocks_json = json.load(stocks)

    result_json = {}

    for item in stocks_json["rows"]:
        # Соответствующий объект из файла с ассортиментом для получения полной информации о товаре
        assortment_item = [obj for obj in assortment_json if obj["externalCode"] == item["externalCode"]][0]

        # Определение категории и подкатегории, т.к. структура json объектов неоднородная
        if len(item["folder"]["pathName"].split("/")) == 1:
            category = item["folder"]["name"]
            subcategory = ""
        elif len(item["folder"]["pathName"].split("/")) == 2:
            category = item["folder"]["pathName"].split("/")[1]
            subcategory = item["folder"]["name"]
        else:
            category = item["folder"]["pathName"].split("/")[1]
            subcategory = item["folder"]["pathName"].split("/")[2]

        # Обработка полученной информации для создания объекта товаров, с которым удобнее работать
        product = {
            "name": item["name"],
            "image": item["image"]["miniature"]["downloadHref"] if "image" in item else None,
            "retailPrice": assortment_item["salePrices"][0]["value"],
            "priceFrom5k": assortment_item["salePrices"][1]["value"],
            "priceFrom15k": assortment_item["salePrices"][2]["value"],
            "priceFrom100k": assortment_item["salePrices"][3]["value"],
        }

        # Создание категорий и подкатегорий в словаре и добавление в них соответствующих объектов
        if category == "Ресницы":
            if category + "_" + subcategory not in result_json:
                result_json[category + "_" + subcategory] = []
            result_json[category + "_" + subcategory].append(product)
        elif category == "Расходные материалы":
            if category not in result_json:
                result_json[category] = {}
            if subcategory not in result_json[category]:
                result_json[category][subcategory] = []
            result_json[category][subcategory].append(product)
        else:
            if category not in result_json:
                result_json[category] = []
            result_json[category].append(product)

    return result_json


if __name__ == '__main__':
    # Для дебага: посмотреть, какого вида будет объединённый json файл
    from config import ASSORTMENT_FILE_PATH, STOCKS_FILE_PATH
    refined_json = merge_json_files(ASSORTMENT_FILE_PATH, STOCKS_FILE_PATH)
    with open("data/refined.json", "w", encoding='utf-8') as file:
        json.dump(refined_json, file, ensure_ascii=False)
