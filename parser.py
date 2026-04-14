import pandas as pd
import json
import re
from datetime import datetime

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

def parse_schedule(file_path):
    df = pd.read_excel(file_path, sheet_name=1, header=None)
    
    # Заголовок недели (строка 5, индекс 4)
    week_info = str(df.iloc[4, 0]) if len(df) > 4 and not pd.isna(df.iloc[4, 0]) else ""
    is_upper = 'верхняя' in week_info.lower()
    
    # Forward fill для дней и времени
    df[0] = df[0].ffill()
    df[2] = df[2].ffill()
    
    # Собираем пары в словарь: (день, время) -> список пар
    pairs_by_day_time = {}
    day_date_map = {}
    
    for _, row in df.iterrows():
        day_str = str(row[0]).strip() if not pd.isna(row[0]) else ""
        time_str = str(row[2]).strip() if not pd.isna(row[2]) else ""
        if not day_str or not time_str:
            continue
        
        # Проверяем, что это строка с днём недели (содержит название дня)
        if not any(d in day_str for d in ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]):
            continue
        
        # Извлекаем чистый день и дату
        day_name = day_str.split()[0].strip()
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', day_str)
        date_str = date_match.group(1) if date_match else ""
        if date_str:
            day_date_map[day_name] = format_date(date_str)
        
        key = (day_name, time_str)
        if key not in pairs_by_day_time:
            pairs_by_day_time[key] = []
        
        # Сначала смотрим колонки группы Б24-142 (6,7,8)
        if not pd.isna(row[6]) and str(row[6]).strip():
            name = str(row[6]).strip()
            teacher = str(row[7]).strip() if not pd.isna(row[7]) else ""
            room = str(row[8]).strip() if not pd.isna(row[8]) else ""
            pairs_by_day_time[key].append({"name": name, "teacher": teacher, "room": room})
        # Fallback: предмет в колонке 3, преподаватель/аудитория в 7/8
        elif not pd.isna(row[3]) and str(row[3]).strip() and not pd.isna(row[7]) and str(row[7]).strip():
            name = str(row[3]).strip()
            teacher = str(row[7]).strip()
            room = str(row[8]).strip() if not pd.isna(row[8]) else ""
            pairs_by_day_time[key].append({"name": name, "teacher": teacher, "room": room})
    
    # Строим расписание по дням, выбирая нужную пару для недели
    schedule_data = {
        "week_title": week_info,
        "original_link": "schedule.xlsx",
        "days": []
    }
    
    # Группируем по дням
    days_dict = {}
    for (day_name, time_str), lessons_list in pairs_by_day_time.items():
        if day_name not in days_dict:
            days_dict[day_name] = []
        # Выбор пары в зависимости от недели
        if is_upper:
            selected = lessons_list[0] if lessons_list else None
        else:
            selected = lessons_list[1] if len(lessons_list) > 1 else (lessons_list[0] if lessons_list else None)
        if selected:
            days_dict[day_name].append({
                "time": time_str,
                "name": selected["name"],
                "teacher": selected["teacher"],
                "room": selected["room"]
            })
    
    # Сортируем дни в порядке недели
    day_order = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    for day in day_order:
        if day in days_dict:
            schedule_data["days"].append({
                "day": day,
                "date": day_date_map.get(day, ""),
                "lessons": sorted(days_dict[day], key=lambda x: x["time"])  # сортировка по времени
            })
    
    return schedule_data

if __name__ == "__main__":
    data = parse_schedule('schedule.xlsx')
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print("Расписание успешно сохранено в data.json")