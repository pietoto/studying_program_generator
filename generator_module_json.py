import json
import os

def generate_module_config():
    input_file = 'result_parsing_plan.json'
    output_file = 'module_config.json'

    # проверка, есть ли файл с результатами парсинга
    if not os.path.exists(input_file):
        print(f"Файл {input_file} не найден")
        return
    # загружаем дисциплины из result_parsing_plan.json
    with open(input_file, 'r', encoding='UTF-8') as f:
        data = json.load(f)

    all_disciplines = data.get('all_disciplines', {})

    if not all_disciplines:
        print("В файле отсутствуют данные о дисциплинах")
        return

    # загружаем старый конфиг (если он существует)
    old_disciplines = {}
    old_modules = {}
    old_config = {}
    if os.path.exists(output_file):
        try:
            with open(output_file, 'r', encoding='UTF-8') as f:
                old_config = json.load(f)
                old_disciplines = old_config.get('disciplines', {})
                old_modules = old_config.get('modules', {})
            print(f"Был найден старый файл {output_file}, сохраняем настройки модулей")

        except Exception as e:
            print(f"Не удалось прочитать старый файл: {e}")
    if old_modules:
        modules = old_modules
        print(f"    Загружено модулей из старого файла: {len(modules)}")
    else:
        modules = {
            "1": {"name": "Модуль 1"},
            "2": {"name": "Модуль 2"},
            "3": {"name": "Модуль 3"},
            "4": {"name": "Модуль 4"},
            "5": {"name": "Модуль 5"},
            "6": {"name": "Модуль 6"},
            "7": {"name": "Модуль 7"},
            "8": {"name": "Модуль 8"},
            "9": {"name": "Модуль 9"},
            "10": {"name": "Модуль 10"},
            "11": {"name": "Модуль 11"},
            "12": {"name": "Модуль 12"},
            "13": {"name": "Модуль 13"},
            "14": {"name": "Модуль 14"},
            "15": {"name": "Модуль 15"},
            "16": {"name": "Модуль 16"}
        }
        print(f"    Создано стандартных модулей: {len(modules)}")
    # формируем новую структуру для module_config.json
    module_config = {
        "modules": modules,
        "disciplines": {}
    }

    # заполняем дисииплины с подстановкой старых значений
    new_count = 0
    preserved_count = 0
    disciplines_list = [] # временный список, который нужен для сортировки
    for code, disc in all_disciplines.items():
        old_module = None

        if code in old_disciplines:
            old_module = old_disciplines[code].get('module')
            if old_module is not None:
                preserved_count += 1
        # сохраняем во временный список
        disciplines_list.append({
            "code": code,
            "name": disc.get("name", ""),
            "disc_code": disc.get("code", ""),
            "credits": disc.get("total_credits", 0.0),
            "is_facultative": disc.get("is_facultative", False),
            "module": old_module,
        })
        if old_module is None:
            new_count += 1
    # сортировка по модулю (сначала none, потом по возрастанию)
    disciplines_list.sort(key=lambda x: (x["module"] is None, x["module"] if x["module"] is not None else 999))
    # записываем отсортированные дисциплины в словарь
    for item in disciplines_list:
        module_config["disciplines"][item["code"]] = {
            "name": item["name"],
            "code": item["disc_code"],
            "credits": item["credits"],
            "is_facultative": item["is_facultative"],
            "module": item["module"]
        }
    # сохраняем новый конфиг
    with open(output_file, 'w', encoding='UTF-8') as f:
        json.dump(module_config, f, ensure_ascii=False, indent=2)

    print(f"Файл {output_file} успешно создан")
    print(f"    Найдено дисциплин: {len(module_config['disciplines'])}")
    print(f"    Сохранено старых настроек: {preserved_count}")
    print(f"    Новых дисциплин (требуют настройки): {new_count}")
    print(f"    Откройте файл {output_file} и внесите название модуля вместо 'Модуль n'")
    print(f"    Внесите в 'module' номер от 1 до {len(modules)} для каждой дисциплины")

if __name__ == '__main__':
    generate_module_config()