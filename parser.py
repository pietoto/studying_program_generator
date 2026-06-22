# импорт библиотек
# встроенная библиотека для работы с xml
import os.path
import xml.etree.cElementTree as ET
# библиотека для работы с json
import json
# библиотека для группировки данных
from collections import defaultdict

# загружаем XML-файл (указываем путь к файлу)
file_path = 'plan.xml'
# попытка выгрузить данные из файла
try:
    tree = ET.parse(file_path)
    root = tree.getroot()
    print(f"Файл {os.path.basename(file_path)} был успешно загружен")
except Exception as e:
    print(f"Файл не удалось загрузить: {e}")
    exit()

# извлекаем пространство имён. ds - главное пространство, где находятся все данные учебного плана
ns = {'ds': 'http://tempuri.org/dsMMISDB.xsd'}

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
    print ("Информация об учебном плане:")
    print(f"    Название: {plan_info['title']}")
    print(f"    Квалификация: {plan_info['qualification']}")
    print(f"    Срок обучения: {plan_info['duration']}")
    print(f"    Учебный год: {plan_info['year']}")
else:
    print(f"Учебный план не найден в {os.path.basename(file_path)}")

# собираем ключевые данные о дисциплинах
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
            # name - название дисциплины
            # code - код дисциплины
            # total_credits - общая трудоёмкость в зет
            # type - тип объекта (2 - дисциплина, 3 - практика, 6 - ГИА)
            # block - код блока
            # is_facultative - факультативный ли предмет
            disciplines[code] = {
                "name": item.get('Дисциплина', ''),
                "code": item.get('ДисциплинаКод', ''),
                "total_credits": float(item.get('ТрудоемкостьКредитов', '0')),
                "type": object_type,
                "block": item.get('КодБлока', ''),
                "is_facultative": item.get('Факультатив', 'false') == 'true'
            }
print(f"Было найдено: {len(disciplines)} дисциплин")

# сбор часов по семестрам для каждой дисциплины
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

# функция для вычисления реального семестра
def get_real_semestr(kurs, sem_in_course):
    return (kurs - 1) * 2 + sem_in_course

# формируем структуру по семестрам
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

# сортировка семестров
sorted_semestrs = dict(sorted(semestrs.items()))

# формирование итоговой структуры для json-файла
result = {
    "plan_info": plan_info,
    "total_disciplines": len(disciplines),
    "semestrs": sorted_semestrs,
    "all_disciplines": disciplines
}
# адрес итогового json-файла
output_file = 'result_plan.json'

# заполняем созданный файл
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"JSON был успешно создан! Найдено {len(disciplines)} дисциплин, {len(sorted_semestrs)} семестров.")
print(f"Файл был сохранён, как: {output_file}")

print("Распределение по семестрам:")
total_credits_all = 0
for sem, discs in sorted_semestrs.items():
    total_credits = sum(d['credits'] for d in discs)
    total_credits_all += total_credits
    print(f"    Семестр {sem}: {len(discs)} дисциплин, {total_credits:.1f} зет")

print(f"Итого зет по всем семестрам: {total_credits_all:.1f}")