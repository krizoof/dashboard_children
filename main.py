import pandas as pd
import re
from datetime import datetime
from unidecode import unidecode
from difflib import SequenceMatcher

def clean_phone(phone):
    if pd.isna(phone):
        return None
    # Преобразуем в строку и удаляем все нецифровые символы
    phone = str(phone)
    digits = re.sub(r'\D', '', phone)
    
    # Проверяем длину номера
    if len(digits) == 11:
        return f"+7{digits[1:]}"
    elif len(digits) == 10:
        return f"+7{digits}"
    return None

def clean_date(date):
    """Стандартизирует формат даты"""
    if pd.isna(date):
        return None
    
    # Если дата уже в формате datetime, преобразуем в строку
    if isinstance(date, datetime):
        return date.strftime('%d.%m.%Y')
    
    # Преобразуем в строку
    date = str(date).strip()
    
    # Пробуем разные форматы даты
    formats = [
        '%d.%m.%Y',  # DD.MM.YYYY
        '%d/%m/%Y',  # DD/MM/YYYY
        '%d-%m-%Y',  # DD-MM-YYYY
        '%Y.%m.%d',  # YYYY.MM.DD
        '%Y/%m/%d',  # YYYY/MM/DD
        '%Y-%m-%d',  # YYYY-MM-DD
        '%d.%m.%y',  # DD.MM.YY
        '%d/%m/%y',  # DD/MM/YY
        '%d-%m-%y',  # DD-MM-YY
    ]
    
    # Пробуем распарсить дату в разных форматах
    for fmt in formats:
        try:
            parsed_date = datetime.strptime(date, fmt)
            return parsed_date.strftime('%d.%m.%Y')
        except ValueError:
            continue
    
    return None

def split_fio(fio):
    """Разбивает ФИО на компоненты и возвращает словарь с фамилией, именем и отчеством"""
    parts = fio.split()
    result = {'фамилия': None, 'имя': None, 'отчество': None}
    
    if len(parts) >= 1:
        result['фамилия'] = parts[0]
    if len(parts) >= 2:
        result['имя'] = parts[1]
    if len(parts) >= 3:
        result['отчество'] = parts[2]
    
    return result

def compare_fio_parts(fio1, fio2):
    """Сравнивает части ФИО с учетом возможных пропусков"""
    parts1 = split_fio(fio1)
    parts2 = split_fio(fio2)
    
    # Если одна из частей пустая, а другая нет, и они достаточно похожи
    for part in ['фамилия', 'имя', 'отчество']:
        if (parts1[part] and not parts2[part]) or (not parts1[part] and parts2[part]):
            continue
        if parts1[part] and parts2[part]:
            similarity = SequenceMatcher(None, parts1[part], parts2[part]).ratio()
            if similarity < 0.8:  # Порог схожести
                return False
    
    return True

def clean_fio(fio):
    if pd.isna(fio):
        return None
    
    # Преобразуем в строку
    fio = str(fio).strip()
    
    # Удаляем точки в конце
    fio = fio.rstrip('.')
    
    # Удаляем скобки и их содержимое
    fio = re.sub(r'\([^)]*\)', '', fio)
    
    # Удаляем лишние пробелы
    fio = ' '.join(fio.split())
    
    return fio

def clean_city(city):
    if pd.isna(city):
        return None
    
    city = str(city).strip()
    
    # Словарь замен для стандартизации названий городов
    city_replacements = {
        'спб': 'Санкт-Петербург',
        'спб.': 'Санкт-Петербург',
        'питер': 'Санкт-Петербург',
        'мск': 'Москва',
        'мск.': 'Москва',
        'нск': 'Новосибирск',
        'нск.': 'Новосибирск',
        'екат': 'Екатеринбург',
        'екат.': 'Екатеринбург',
    }
    
    # Удаляем сокращения типа "г.", "п.", "пос.", "с.", "д."
    city = re.sub(r'^(г\.|п\.|пос\.|с\.|д\.|город|поселок|село|деревня)\s*', '', city, flags=re.IGNORECASE)
    
    # Удаляем лишние пробелы
    city = ' '.join(city.split())
    
    # Приводим к нижнему регистру для сравнения
    city_lower = city.lower()
    
    # Проверяем наличие в словаре замен
    for key, value in city_replacements.items():
        if city_lower == key:
            return value
    
    # Если город не найден в словаре замен, приводим к формату "Первая буква заглавная"
    return city.title()

def clean_region(region):
    if pd.isna(region):
        return None
    
    region = str(region).strip()
    
    # Словарь замен для стандартизации названий регионов
    region_replacements = {
        'мо': 'Московская область',
        'моск. обл.': 'Московская область',
        'моск.область': 'Московская область',
        'ленинградская': 'Ленинградская область',
        'ленинград. обл.': 'Ленинградская область',
        'ленинград.область': 'Ленинградская область',
        'спб': 'Санкт-Петербург',
        'спб.': 'Санкт-Петербург',
        'питер': 'Санкт-Петербург',
    }
    
    # Удаляем сокращения
    region = re.sub(r'(обл\.|респ\.|АО|край)\s*', '', region, flags=re.IGNORECASE)
    
    # Удаляем лишние пробелы
    region = ' '.join(region.split())
    
    # Приводим к нижнему регистру для сравнения
    region_lower = region.lower()
    
    # Проверяем наличие в словаре замен
    for key, value in region_replacements.items():
        if region_lower == key:
            return value
    
    # Если регион не найден в словаре замен, приводим к формату "Первая буква заглавная"
    return region.title()

def find_duplicates(df):
    # Создаем копию DataFrame для работы
    df_clean = df.copy()
    
    # Очищаем все поля
    df_clean['ФИО'] = df_clean['ФИО'].apply(clean_fio)
    df_clean['ТЕЛЕФОН'] = df_clean['ТЕЛЕФОН'].apply(clean_phone)
    df_clean['ДАТА РОЖДЕНИЯ'] = df_clean['ДАТА РОЖДЕНИЯ'].apply(clean_date)
    
    # Создаем список для хранения групп дубликатов
    duplicate_groups = []
    processed_indices = set()
    
    # Проходим по всем строкам
    for idx, row in df_clean.iterrows():
        if idx in processed_indices:
            continue
        
        # Находим потенциальные дубликаты по ФИО
        potential_duplicates = []
        for other_idx, other_row in df_clean.iterrows():
            if other_idx != idx and other_idx not in processed_indices:
                # Проверяем схожесть ФИО
                if compare_fio_parts(row['ФИО'], other_row['ФИО']):
                    # Проверяем совпадение хотя бы одного из полей
                    if ((row['ТЕЛЕФОН'] == other_row['ТЕЛЕФОН'] and row['ТЕЛЕФОН'] is not None) or
                        (row['ДАТА РОЖДЕНИЯ'] == other_row['ДАТА РОЖДЕНИЯ'] and row['ДАТА РОЖДЕНИЯ'] is not None) or
                        (row['ЭЛ.ПОЧТА'] == other_row['ЭЛ.ПОЧТА'] and row['ЭЛ.ПОЧТА'] is not None) or
                        (row['ТЕЛЕГРАМ'] == other_row['ТЕЛЕГРАМ'] and row['ТЕЛЕГРАМ'] is not None)):
                        potential_duplicates.append(other_idx)
        
        if potential_duplicates:
            duplicate_groups.append([idx] + potential_duplicates)
            processed_indices.update([idx] + potential_duplicates)
    
    return duplicate_groups

def find_duplicates_by_fields(df):
    """Находит дубликаты по совпадению минимум двух полей из: телефон, дата рождения, почта, телеграм"""
    # Создаем копию DataFrame для работы
    df_clean = df.copy()
    
    # Очищаем поля
    df_clean['ТЕЛЕФОН'] = df_clean['ТЕЛЕФОН'].apply(clean_phone)
    df_clean['ДАТА РОЖДЕНИЯ'] = df_clean['ДАТА РОЖДЕНИЯ'].apply(clean_date)
    
    # Создаем список для хранения групп дубликатов
    duplicate_groups = []
    processed_indices = set()
    
    # Проходим по всем строкам
    for idx, row in df_clean.iterrows():
        if idx in processed_indices:
            continue
        
        # Находим потенциальные дубликаты по полям
        potential_duplicates = []
        for other_idx, other_row in df_clean.iterrows():
            if other_idx != idx and other_idx not in processed_indices:
                # Считаем количество совпадающих полей
                matching_fields = 0
                if row['ТЕЛЕФОН'] == other_row['ТЕЛЕФОН'] and row['ТЕЛЕФОН'] is not None:
                    matching_fields += 1
                if row['ДАТА РОЖДЕНИЯ'] == other_row['ДАТА РОЖДЕНИЯ'] and row['ДАТА РОЖДЕНИЯ'] is not None:
                    matching_fields += 1
                if row['ЭЛ.ПОЧТА'] == other_row['ЭЛ.ПОЧТА'] and row['ЭЛ.ПОЧТА'] is not None:
                    matching_fields += 1
                if row['ТЕЛЕГРАМ'] == other_row['ТЕЛЕГРАМ'] and row['ТЕЛЕГРАМ'] is not None:
                    matching_fields += 1
                
                # Если совпадает минимум 2 поля, считаем записи дубликатами
                if matching_fields >= 2:
                    potential_duplicates.append(other_idx)
        
        if potential_duplicates:
            duplicate_groups.append([idx] + potential_duplicates)
            processed_indices.update([idx] + potential_duplicates)
    
    return duplicate_groups

def merge_duplicate_rows(group):
    # Сортируем по году в обратном порядке (новые записи первыми)
    group = group.sort_values('Год', ascending=False)
    
    # Берем первую строку как основу
    result = group.iloc[0].copy()
    
    # Проходим по остальным строкам и заполняем пустые значения
    for _, row in group.iloc[1:].iterrows():
        for column in group.columns:
            if pd.isna(result[column]) and not pd.isna(row[column]):
                result[column] = row[column]
    
    return result

def find_similar_names(names, threshold=0.8):
    """Находит похожие названия и объединяет их"""
    from difflib import SequenceMatcher
    
    # Создаем словарь для хранения групп похожих названий
    similar_groups = {}
    processed = set()
    
    # Преобразуем список в множество для уникальных значений
    unique_names = set(names)
    
    for name1 in unique_names:
        if name1 in processed:
            continue
            
        group = [name1]
        processed.add(name1)
        
        for name2 in unique_names:
            if name2 in processed:
                continue
                
            # Сравниваем названия
            similarity = SequenceMatcher(None, name1.lower(), name2.lower()).ratio()
            if similarity >= threshold:
                group.append(name2)
                processed.add(name2)
        
        if len(group) > 1:
            # Выбираем самое длинное название как основное
            main_name = max(group, key=len)
            similar_groups[main_name] = group
    
    return similar_groups

def parse_competition_place(status):
    """Определяет место в соревновании на основе статуса"""
    if pd.isna(status):
        return None
    
    status = str(status).lower().strip()
    
    # Ключевые слова для определения места (в нижнем регистре)
    winner_keywords = ['1 место', 'первое место', 'победитель', 'победила', 'победил', 'победители', 'победителей']
    prize_keywords = ['2 место', '3 место', 'второе место', 'третье место', 'призер', 'призерка', 'полуфинал', 'финал', 'призеры', 'призеров']
    participant_keywords = ['участник', 'участница', 'участвовал', 'участвовала', 'участники', 'участников', 'уастник']
    
    # Проверяем на победителя
    if any(keyword.lower() in status for keyword in winner_keywords):
        return 'Победитель'
    
    # Проверяем на призера
    if any(keyword.lower() in status for keyword in prize_keywords):
        return 'Призер'
    
    # Проверяем на участника
    if any(keyword.lower() in status for keyword in participant_keywords):
        return 'Участник'
    
    return None

def process_students():
    # Читаем исходный файл
    df = pd.read_excel('База.xlsx')
    
    # Создаем новые DataFrame'ы
    students_df = pd.DataFrame(columns=[
        'id', 'ФИО', 'ТЕЛЕФОН', 'РЕГИОН', 'ГОРОД', 
        'ШКОЛА', 'ДАТА РОЖДЕНИЯ', 'ЭЛ.ПОЧТА', 'ТЕЛЕГРАМ'
    ])
    
    events_df = pd.DataFrame(columns=['id', 'Мероприятие', 'Тип мероприятия', 'Год'])
    relations_df = pd.DataFrame(columns=['id', 'id_студента', 'id_мероприятия', 'Место'])
    
    # Очищаем и стандартизируем данные
    df['ТЕЛЕФОН'] = df['ТЕЛЕФОН'].apply(clean_phone)
    df['ФИО'] = df['ФИО'].apply(clean_fio)
    df['ГОРОД'] = df['ГОРОД'].apply(clean_city)
    df['РЕГИОН'] = df['РЕГИОН'].apply(clean_region)
    df['ДАТА РОЖДЕНИЯ'] = df['ДАТА РОЖДЕНИЯ'].apply(clean_date)
    
    # Объединяем похожие названия регионов и городов
    region_groups = find_similar_names(df['РЕГИОН'].dropna().unique())
    city_groups = find_similar_names(df['ГОРОД'].dropna().unique())
    
    # Применяем объединение
    for main_region, group in region_groups.items():
        df.loc[df['РЕГИОН'].isin(group), 'РЕГИОН'] = main_region
    
    for main_city, group in city_groups.items():
        df.loc[df['ГОРОД'].isin(group), 'ГОРОД'] = main_city
    
    # Заполняем отсутствующие годы текущим годом
    current_year = datetime.now().year
    df['Год'] = df['Год'].fillna(current_year)
    
    # Удаляем строки, где нет ни телефона, ни ФИО
    df = df.dropna(subset=['ТЕЛЕФОН', 'ФИО'], how='all')
    
    # Создаем словарь для хранения мероприятий
    events_dict = {}
    event_id = 1
    
    # Создаем словарь для хранения связей
    relations_dict = {}
    relation_id = 1
    
    # Первый проход: находим дубликаты по ФИО
    duplicate_groups = find_duplicates(df)
    processed_indices = set()
    
    # Обрабатываем группы с дубликатами
    for group_indices in duplicate_groups:
        group = df.loc[group_indices]
        merged_row = merge_duplicate_rows(group)
        
        # Добавляем студента
        student_id = len(students_df) + 1
        students_df.loc[len(students_df)] = {
            'id': student_id,
            'ФИО': merged_row['ФИО'],
            'ТЕЛЕФОН': merged_row['ТЕЛЕФОН'],
            'РЕГИОН': merged_row['РЕГИОН'],
            'ГОРОД': merged_row['ГОРОД'],
            'ШКОЛА': merged_row['ШКОЛА'],
            'ДАТА РОЖДЕНИЯ': merged_row['ДАТА РОЖДЕНИЯ'],
            'ЭЛ.ПОЧТА': merged_row['ЭЛ.ПОЧТА'],
            'ТЕЛЕГРАМ': merged_row['ТЕЛЕГРАМ']
        }
        
        # Обрабатываем мероприятия для этой группы
        for _, row in group.iterrows():
            event_type = 'Соревнование' if pd.notna(row['Статус']) and str(row['Статус']).strip() != '' else 'Курс'
            event_key = (row['Мероприятие'], row['Год'], event_type)
            
            # Добавляем мероприятие, если его еще нет
            if event_key not in events_dict:
                events_dict[event_key] = event_id
                events_df.loc[len(events_df)] = {
                    'id': event_id,
                    'Мероприятие': row['Мероприятие'],
                    'Тип мероприятия': event_type,
                    'Год': row['Год']
                }
                event_id += 1
            
            # Добавляем связь
            relation_key = (student_id, events_dict[event_key])
            if relation_key not in relations_dict:
                relations_dict[relation_key] = relation_id
                relations_df.loc[len(relations_df)] = {
                    'id': relation_id,
                    'id_студента': student_id,
                    'id_мероприятия': events_dict[event_key],
                    'Место': parse_competition_place(row['Статус']) if event_type == 'Соревнование' else None
                }
                relation_id += 1
        
        processed_indices.update(group_indices)
    
    # Второй проход: находим дубликаты по полям
    duplicate_groups = find_duplicates_by_fields(df)
    
    # Обрабатываем оставшиеся группы с дубликатами
    for group_indices in duplicate_groups:
        if not any(idx in processed_indices for idx in group_indices):
            group = df.loc[group_indices]
            merged_row = merge_duplicate_rows(group)
            
            # Добавляем студента
            student_id = len(students_df) + 1
            students_df.loc[len(students_df)] = {
                'id': student_id,
                'ФИО': merged_row['ФИО'],
                'ТЕЛЕФОН': merged_row['ТЕЛЕФОН'],
                'РЕГИОН': merged_row['РЕГИОН'],
                'ГОРОД': merged_row['ГОРОД'],
                'ШКОЛА': merged_row['ШКОЛА'],
                'ДАТА РОЖДЕНИЯ': merged_row['ДАТА РОЖДЕНИЯ'],
                'ЭЛ.ПОЧТА': merged_row['ЭЛ.ПОЧТА'],
                'ТЕЛЕГРАМ': merged_row['ТЕЛЕГРАМ']
            }
            
            # Обрабатываем мероприятия для этой группы
            for _, row in group.iterrows():
                event_type = 'Соревнование' if pd.notna(row['Статус']) and str(row['Статус']).strip() != '' else 'Курс'
                event_key = (row['Мероприятие'], row['Год'], event_type)
                
                # Добавляем мероприятие, если его еще нет
                if event_key not in events_dict:
                    events_dict[event_key] = event_id
                    events_df.loc[len(events_df)] = {
                        'id': event_id,
                        'Мероприятие': row['Мероприятие'],
                        'Тип мероприятия': event_type,
                        'Год': row['Год']
                    }
                    event_id += 1
                
                # Добавляем связь
                relation_key = (student_id, events_dict[event_key])
                if relation_key not in relations_dict:
                    relations_dict[relation_key] = relation_id
                    relations_df.loc[len(relations_df)] = {
                        'id': relation_id,
                        'id_студента': student_id,
                        'id_мероприятия': events_dict[event_key],
                        'Место': parse_competition_place(row['Статус']) if event_type == 'Соревнование' else None
                    }
                    relation_id += 1
            
            processed_indices.update(group_indices)
    
    # Обрабатываем оставшихся студентов без дубликатов
    for idx, row in df.iterrows():
        if idx not in processed_indices:
            # Добавляем студента
            student_id = len(students_df) + 1
            students_df.loc[len(students_df)] = {
                'id': student_id,
                'ФИО': row['ФИО'],
                'ТЕЛЕФОН': row['ТЕЛЕФОН'],
                'РЕГИОН': row['РЕГИОН'],
                'ГОРОД': row['ГОРОД'],
                'ШКОЛА': row['ШКОЛА'],
                'ДАТА РОЖДЕНИЯ': row['ДАТА РОЖДЕНИЯ'],
                'ЭЛ.ПОЧТА': row['ЭЛ.ПОЧТА'],
                'ТЕЛЕГРАМ': row['ТЕЛЕГРАМ']
            }
            
            # Добавляем мероприятие и связь
            event_type = 'Соревнование' if pd.notna(row['Статус']) and str(row['Статус']).strip() != '' else 'Курс'
            event_key = (row['Мероприятие'], row['Год'], event_type)
            
            if event_key not in events_dict:
                events_dict[event_key] = event_id
                events_df.loc[len(events_df)] = {
                    'id': event_id,
                    'Мероприятие': row['Мероприятие'],
                    'Тип мероприятия': event_type,
                    'Год': row['Год']
                }
                event_id += 1
            
            relation_key = (student_id, events_dict[event_key])
            if relation_key not in relations_dict:
                relations_dict[relation_key] = relation_id
                relations_df.loc[len(relations_df)] = {
                    'id': relation_id,
                    'id_студента': student_id,
                    'id_мероприятия': events_dict[event_key],
                    'Место': parse_competition_place(row['Статус']) if event_type == 'Соревнование' else None
                }
                relation_id += 1
    
    # Сохраняем результаты
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Сохраняем студентов
    students_filename = f'students_clean_{timestamp}.xlsx'
    students_df.to_excel(students_filename, index=False)
    print(f'Обработка студентов завершена. Результат сохранен в файл: {students_filename}')
    print(f'Всего уникальных студентов: {len(students_df)}')
    
    # Сохраняем мероприятия
    events_filename = f'events_clean_{timestamp}.xlsx'
    events_df.to_excel(events_filename, index=False)
    print(f'Обработка мероприятий завершена. Результат сохранен в файл: {events_filename}')
    print(f'Всего уникальных мероприятий: {len(events_df)}')
    
    # Сохраняем связи
    relations_filename = f'student_event_relations_{timestamp}.xlsx'
    relations_df.to_excel(relations_filename, index=False)
    print(f'Создание связей завершено. Результат сохранен в файл: {relations_filename}')
    print(f'Всего связей: {len(relations_df)}')

if __name__ == '__main__':
    process_students()
