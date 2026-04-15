import json
import re
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter

MONTHS = {
    1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
    7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
}

def format_date(date_str):
    """Преобразует 'DD.MM.YYYY' в 'DD месяц' (например, '13 апреля')"""
    try:
        dt = datetime.strptime(date_str.strip(), '%d.%m.%Y')
        return f"{dt.day} {MONTHS[dt.month]}"
    except:
        return date_str

def clean_value(val):
    """Очищает значение ячейки от лишних пробелов и None."""
    if val is None:
        return ""
    if isinstance(val, str):
        return val.strip()
    return str(val).strip()

def parse_schedule(file_path, sheet_name=None, target_group="Б24-142", week="верхняя"):
    """
    Парсит Excel-файл с расписанием и возвращает словарь в формате,
    идентичном текущему data.json.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    if sheet_name is None:
        # По умолчанию берём второй лист (индекс 1), как у вас было sheet_name=1
        sheet = wb.worksheets[1]
    else:
        sheet = wb[sheet_name]
    
    # Определяем заголовок недели (строка 5, индекс 4 в pandas)
    week_title = clean_value(sheet.cell(row=5, column=1).value)
    is_upper = 'верхняя' in week_title.lower()
    
    # Находим строку с заголовками ("Д/н", "Время")
    header_row = None
    for r in range(1, sheet.max_row + 1):
        row_vals = [clean_value(sheet.cell(row=r, column=c).value) for c in range(1, 10)]
        if any("Д/н" in v or "Время" in v for v in row_vals):
            header_row = r
            break
    if header_row is None:
        raise ValueError("Не найден заголовок расписания.")
    
    # Определяем имена групп из строк заголовка
    group1_name = None
    group2_name = None
    for r in (header_row, header_row + 1):
        d_val = clean_value(sheet.cell(row=r, column=4).value)
        g_val = clean_value(sheet.cell(row=r, column=7).value)
        if "Б24" in d_val:
            group1_name = d_val.split()[0]  # "Б24-141"
        if "Б24" in g_val:
            group2_name = g_val.split()[0]  # "Б24-142"
        if group1_name and group2_name:
            break
    
    if not group1_name or not group2_name:
        raise ValueError("Не удалось определить названия групп.")
    
    # Определяем, какие колонки соответствуют целевой группе
    if target_group.startswith(group1_name):
        group_cols = (4, 5, 6)  # D, E, F
    elif target_group.startswith(group2_name):
        group_cols = (7, 8, 9)  # G, H, I
    else:
        raise ValueError(f"Группа {target_group} не найдена на листе.")
    
    # Словарь для сбора пар: (день, время) -> список предметов
    pairs_by_day_time = defaultdict(list)
    day_date_map = {}
    
    current_day = None
    current_time = None
    current_pair_num = None
    
    # Начинаем со строки после заголовка
    for r in range(header_row + 2, sheet.max_row + 1):
        a_val = clean_value(sheet.cell(row=r, column=1).value)
        b_val = sheet.cell(row=r, column=2).value
        c_val = clean_value(sheet.cell(row=r, column=3).value)
        
        # Если столбец A содержит день недели
        if a_val and any(d in a_val for d in ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]):
            current_day = a_val.split()[0]
            # Извлекаем дату
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', a_val)
            if date_match:
                day_date_map[current_day] = format_date(date_match.group(1))
        
        # Если нет дня, но есть продолжение (объединённые ячейки), пропускаем служебные строки
        if current_day is None:
            continue
        
        # Номер пары и время
        if b_val is not None:
            current_pair_num = b_val
            current_time = c_val if c_val else ""
        # Если b_val пусто, значит это продолжение той же пары (объединённая ячейка)
        if not current_time:
            continue
        
        # Читаем предмет, преподавателя, аудиторию для целевой группы
        subj_col, teach_col, room_col = group_cols
        subject = clean_value(sheet.cell(row=r, column=subj_col).value)
        teacher = clean_value(sheet.cell(row=r, column=teach_col).value)
        room = clean_value(sheet.cell(row=r, column=room_col).value)
        
        if not subject:
            # Если у нашей группы пусто, проверяем, не общая ли это лекция
            # Смотрим другую группу
            other_subj_col = 7 if group_cols[0] == 4 else 4
            other_teach_col = 8 if group_cols[0] == 4 else 5
            other_room_col = 9 if group_cols[0] == 4 else 6
            other_subject = clean_value(sheet.cell(row=r, column=other_subj_col).value)
            if other_subject and "(лек)" in other_subject.lower():
                subject = other_subject
                teacher = clean_value(sheet.cell(row=r, column=other_teach_col).value)
                room = clean_value(sheet.cell(row=r, column=other_room_col).value)
        
        if subject:
            pairs_by_day_time[(current_day, current_time)].append({
                "name": subject,
                "teacher": teacher,
                "room": room
            })
    
    # Теперь формируем итоговое расписание, выбирая предметы по неделям
    days_dict = defaultdict(list)
    for (day, time_str), lessons in pairs_by_day_time.items():
        # Если в одной временной ячейке несколько предметов (из-за переносов строк в Excel)
        # В нашей структуре lessons может содержать несколько записей для одного времени
        # из-за объединённых ячеек, но мы соберём их в один список для дальнейшей фильтрации.
        # Фактически, для каждой пары времени у нас есть список предметов,
        # которые могут относиться к разным неделям или подгруппам.
        
        # Объединяем все предметы для этого времени
        all_subjects = []
        for les in lessons:
            # Проверяем, есть ли перенос строки (верхняя/нижняя неделя)
            name = les["name"]
            if "\n" in name:
                parts = name.split("\n")
                upper = parts[0].strip()
                lower = parts[1].strip() if len(parts) > 1 else ""
                if week == "верхняя":
                    if upper:
                        all_subjects.append({"name": upper, "teacher": les["teacher"], "room": les["room"]})
                else:
                    if lower:
                        all_subjects.append({"name": lower, "teacher": les["teacher"], "room": les["room"]})
                    elif upper:  # если нижней нет, берём верхнюю (на всякий случай)
                        all_subjects.append({"name": upper, "teacher": les["teacher"], "room": les["room"]})
            else:
                all_subjects.append(les)
        
        # Теперь фильтруем: убираем дубликаты и оставляем только те, что соответствуют неделе
        # Если есть физика лаб, то для верхней недели исключаем, для нижней оставляем только её
        filtered = []
        for subj in all_subjects:
            name_lower = subj["name"].lower()
            is_physics_lab = "физика" in name_lower and ("лаб" in name_lower or "лабораторная" in name_lower)
            if is_upper and is_physics_lab:
                continue
            if not is_upper and is_physics_lab:
                # оставляем только физику лаб, остальные для этого времени могут быть верхней недели
                filtered = [subj]
                break
            filtered.append(subj)
        
        # Если после фильтрации осталось несколько предметов (например, подгруппы), добавляем все
        for subj in filtered:
            days_dict[day].append({
                "time": time_str,
                "name": subj["name"],
                "teacher": subj["teacher"],
                "room": subj["room"]
            })
    
    # Удаляем возможные дубликаты с одинаковым временем и названием
    for day in days_dict:
        unique_lessons = []
        seen = set()
        for lesson in days_dict[day]:
            key = (lesson["time"], lesson["name"])
            if key not in seen:
                seen.add(key)
                unique_lessons.append(lesson)
        days_dict[day] = unique_lessons
    
    # Сортируем дни по порядку
    day_order = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    result_days = []
    for day in day_order:
        if day in days_dict:
            lessons = sorted(days_dict[day], key=lambda x: time_to_minutes(x["time"]))
            result_days.append({
                "day": day,
                "date": day_date_map.get(day, ""),
                "lessons": lessons
            })
    
    return {
        "week_title": week_title,
        "original_link": "schedule.xlsx",
        "days": result_days
    }

def time_to_minutes(time_str):
    """Переводит '9.00-10.20' в количество минут от начала дня для сортировки."""
    try:
        start = time_str.split('-')[0].strip()
        parts = start.split('.')
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m
    except:
        return 0

if __name__ == "__main__":
    # Здесь можно указать нужную группу и неделю.
    # По умолчанию берём группу Б24-142 и неделю из названия листа.
    data = parse_schedule('schedule.xlsx', target_group='Б24-142')
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print("Расписание успешно сохранено в data.json")