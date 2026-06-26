# импорт библиотек
# встроенная библиотека для работы с xml
import os.path
import xml.etree.cElementTree as ET
# библиотека для работы с json
import json
# библиотека для группировки данных
from collections import defaultdict

# загружаем XML-файл (указываем путь к файлу)
def load_plx_file(file_path):
    # попытка выгрузить данные из файла
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        print(f"Файл {os.path.basename(file_path)} был успешно загружен")
        return root
    except Exception as e:
        print(f"Файл не удалось загрузить: {e}")
        return None

# извлекаем информацию о плане
def extract_plan_info(root, ns):
    # извлекаем общую информацию о плане
    plan = root.find('.//ds:Планы', namespaces=ns)
    # словарь для сохранения данных о плане
    plan_info = {}
    # если найден тег Планы, то извлекаем из него атрибуты
    if plan is not None:
        # title - название направления
        # qualification - квалификация: бакалавр, специалист, магистр
        # duration - срок обучения в годах
        # year - год обучения (например, 2026-2027)
        plan_info = {
            "title": plan.get('Титул', '').replace('&#xD;&#xA;', ' '),
            "qualification": plan.get('Квалификация', ''),
            "duration": plan.get('СрокОбучения', ''),
            "year": plan.get('УчебныйГод', '')
        }
        print("Информация об учебном плане:")
        print(f"    Название: {plan_info['title']}")
        print(f"    Квалификация: {plan_info['qualification']}")
        print(f"    Срок обучения: {plan_info['duration']}")
        print(f"    Учебный год: {plan_info['year']}")
    else:
        print(f"Учебный план не найден")
    return plan_info

# собираем ключевые данные о дисциплинах
def extract_disciplines(root, ns):
    disciplines = {}
    for item in root.findall('.//ds:ПланыСтроки', namespaces=ns):
        code = item.get('Код')
        # если есть код, то обрабатыаем элемент
        if code:
            # ТипОбъекта определяет, что за элемент перед нами:
            # 1 - модуль из дисциплин
            # 2 - дисциплина
            # 3 - практика
            # 5 - блок по выбору
            # 6 - ГИА (государственная итоговая аттестация)
            object_type = item.get('ТипОбъекта', '')
            # имеет смысл рассматривать обязательные дисциплины, практику и ГИА
            if object_type in ['2', '3', '6']:
                # задаём флаг для дисциплин, которые считаются без ЗЕТ
                count_without_zet = item.get('СчитатьБезЗЕТ', 'false') == 'true'
                # name - название дисциплины
                # code - код дисциплины
                # total_credits - общая трудоёмкость в зет
                # type - тип объекта (2 - дисциплина, 3 - практика, 6 - ГИА)
                # block - код блока
                # is_facultative - факультативный ли предмет
                # count_without_zet - проверяем, считается ли дисциплина без зетов
                disciplines[code] = {
                    "name": item.get('Дисциплина', ''),
                    "code": item.get('ДисциплинаКод', ''),
                    "total_credits": float(item.get('ТрудоемкостьКредитов', '0')),
                    "type": object_type,
                    "block": item.get('КодБлока', ''),
                    "is_facultative": item.get('Факультатив', 'false') == 'true',
                    "count_without_zet": count_without_zet
                }
    print(f"Было найдено: {len(disciplines)} дисциплин")
    return disciplines

# сбор часов по семестрам для каждой дисциплины
def extract_semestr_hours(root, ns, disciplines):
    semestr_hours = defaultdict(list)
    # ищем все теги ПланыНовыеЧасы
    for hours in root.findall('.//ds:ПланыНовыеЧасы', namespaces=ns):
        # код объекта, к которому относятся часы
        object_code = hours.get('КодОбъекта')
        # номер курма
        kurs = hours.get('Курс')
        # номер семестра внутри курса (1 или 2)
        semestr = hours.get('Семестр')
        # вида работы определяется из тега СправочникВидыРабот
        work_type = hours.get('КодВидаРаботы')
        # количество часов
        count = float(hours.get('Количество', '0'))

        if object_code and object_code in disciplines:
            if work_type == '1000' and kurs and semestr and count > 0:
                # преобразуем курс и семестр в числа
                kurs_int = int(kurs)
                semestr_int = int(semestr)
                pair = (kurs_int, semestr_int)

                # проверяем не была ли ранее записана такая же пара для дисциплина (нужно, чтобы избежать дублирования)
                exists = False
                for item in semestr_hours[object_code]:
                    if item['kurs'] == kurs_int and item['semestr'] == semestr_int:
                        exists = True
                        break
                if not exists:
                    semestr_hours[object_code].append(
                        {
                            'kurs': kurs_int,
                            'semestr': semestr_int,
                            'hours': count
                        }
                    )

    # выводим статистику (сколько всего записей (курс, семестр) было найдено)
    total_pairs = sum(len(v) for v in semestr_hours.values())
    print(f"Найдено записей с часами по семестрам: {total_pairs}")
    return semestr_hours

# функция для вычисления реального семестра
def get_real_semestr(kurs, sem_in_course):
    return (kurs - 1) * 2 + sem_in_course

# группируем дисциплины по родительскому коду
def merge_dv_disciplines(disciplines, semestr_hours):
    dv_groups = defaultdict(list)
    for code, disc in disciplines.items():
        disc_code = disc.get('code', '')
        # проверка есть ли в коде предмета .ДВ. (т.е. дисциплина по выбору)
        if '.ДВ.' in disc_code:
            # извлекаем родительский код (всё до последнего .xx)
            parent_code = '.'.join(disc_code.split('.')[:-1])
            dv_groups[parent_code].append(code)

    # создаём словарь для замены старого кода на новый
    merge_map = {}

    for parent_code, codes_to_merge in dv_groups.items():
        if len(codes_to_merge) > 1:
            # берём первую дисциплину за основу
            base_code = codes_to_merge[0]
            base_disc = disciplines[base_code]

            # собираем название всех дисциплин в группе
            names = [disciplines[code]['name'] for code in codes_to_merge]
            merged_name = ' / '.join(names)

            # создаём новую объединённую дисциплину
            merged_disc = {
                "name": merged_name,
                "code": parent_code,  # родительский код без .xx на конце
                "total_credits": base_disc['total_credits'],
                "type": base_disc['type'],
                "block": base_disc['block'],
                "is_facultative": base_disc['is_facultative'],
                "is_dv": True,
                "dv_children": codes_to_merge  # сохраняем исходные коды

            }
            # заменяем первую дисциплину на объединённую
            disciplines[base_code] = merged_disc

            # удаляем остальные дисциплины из словаря
            for code in codes_to_merge[1:]:
                del disciplines[code]

            # запоминаем, что нужно объяснить в семестрах
            for code in codes_to_merge[1:]:
                merge_map[code] = base_code

    print(f"Было объединено {len(dv_groups)} групп дисциплин по выбору")

    # создаём словарь для часов с объединёнными дисциплинами
    new_semestr_hours = defaultdict(list)
    for obj_code, hours_list in semestr_hours.items():
        # если код был объединён, то используем новый код
        new_code = merge_map.get(obj_code, obj_code)
        # проверяем, что новый код существует в дисциплинах
        if new_code in disciplines:
            # добавляем часы в новый словарь
            for h in hours_list:
                # проверяем, не добавлены ли они были ранее
                exists = False
                for item in new_semestr_hours[new_code]:
                    if item['kurs'] == h['kurs'] and item['semestr'] == h['semestr']:
                        exists = True
                        break
                if not exists:
                    new_semestr_hours[new_code].append(h)

    semestr_hours = new_semestr_hours
    # обновляем total_pairs
    total_pairs = sum(len(v) for v in semestr_hours.values())
    print(f"Найдено записей с часами по семестрам после объединения: {total_pairs}")
    return disciplines, new_semestr_hours

# формируем структуру по семестрам
def build_semestr_structure(disciplines, semestr_hours):
    # ключ - реальный номер семестра, значение - список дисциплин в этом семестре
    semestrs = defaultdict(list)
    for object_code, hours_list in semestr_hours.items():
        if object_code in disciplines:
            # получаем данные о дисциплине
            disc = disciplines[object_code]
            # перебираем все пары (курс, семестр) для этой дисциплины
            for h in hours_list:
                # вычисляем реальный номер семестра (1-8 для бакалавриата)
                real_semestr = get_real_semestr(h['kurs'], h['semestr'])
                # если в дисциплине нет зетов, то credits = 0
                if disc.get('count_without_zet', False):
                    credits_in_semestr = 0.0
                else:
                    # вычисляем количество зет в текущем семестре (округляем до 1 знака после запятой)
                    credits_in_semestr = round(h['hours'] / 36, 1)

                # проверяем не добавлена ли текущая дисцпиплина ранее в этот семестр
                exists = False
                for existing in semestrs[real_semestr]:
                    if existing.get('code') == disc['code']:
                        exists = True
                        break

                if not exists:
                    # name - назавние дисциплины
                    # code - код дисциплины по учебному плану
                    # credits - количество зет в текущем семестре
                    # hours - количество часов в этом семестре
                    # total_credits - суммарное количество зет по дисциплине
                    # is_facultative - факультативность
                    # type - тип элемента (2 - дисциплина, 3 - практика, 6 - ГИА)
                    # block - код блока
                    semestrs[real_semestr].append({
                        "name": disc['name'],
                        "code": disc['code'],
                        "credits": credits_in_semestr,
                        "hours": h['hours'],
                        "total_credits": disc['total_credits'],
                        "is_facultative": disc['is_facultative'],
                        "type": disc['type'],
                        "block": disc['block']
                    })
    return dict(sorted(semestrs.items()))

# сохранение в JSON
def save_to_json(result, output_file):
    # заполняем созданный файл
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"JSON был успешно создан!")
    print(f"Файл был сохранён, как: {output_file}")

# вывод статистики
def print_statistics(sorted_semestrs):
    print("Распределение по семестрам:")
    total_credits_all = 0
    for sem, discs in sorted_semestrs.items():
        total_credits = sum(d['credits'] for d in discs)
        total_credits_all += total_credits
        print(f"    Семестр {sem}: {len(discs)} дисциплин, {total_credits:.1f} зет")

    print(f"Итого зет по всем семестрам: {total_credits_all:.1f}")


def main():
    # указываем путь к файлу
    file_path = 'b09.03.03_03_ИКНК_2026.plx'

    # указываем пространтсов имён
    ns = {'ds': 'http://tempuri.org/dsMMISDB.xsd'}

    # загрузка файла
    root = load_plx_file(file_path)
    if root is None:
        return

    # извлечение информации о плане
    plan_info = extract_plan_info(root, ns)

    # извлечение дисциплин
    disciplines = extract_disciplines(root, ns)

    # извлечение часов по семестрам
    semestr_hours = extract_semestr_hours(root, ns, disciplines)

    # объединение дисциплин по выбору
    disciplines, semestr_hours = merge_dv_disciplines(disciplines, semestr_hours)

    # формирование структуры по семестрам
    sorted_semestrs = build_semestr_structure(disciplines, semestr_hours)

    # формирование итоговой структуры для json-файла
    result = {
        "plan_info": plan_info,
        "total_disciplines": len(disciplines),
        "semestrs": sorted_semestrs,
        "all_disciplines": disciplines
    }
    # сохранение в JSON
    output_file = 'result_parsing_plan.json'
    save_to_json(result, output_file)

    # выбор статистики
    print_statistics(sorted_semestrs)

    print(f"Найдено {len(disciplines)} дисциплин, {len(sorted_semestrs)} семестров")

if __name__ == '__main__':
    main()