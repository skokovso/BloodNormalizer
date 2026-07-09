#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Препроцессор данных трансфузий крови
=====================================

Модуль для нормализации и предобработки Excel-файлов с трансфузиями крови.

Назначение:
    Подготовка данных для расчета оптимального запаса компонентов крови.
    Нормализует группы крови с учетом подгрупп A2 и A2B, определяет типы
    компонентов, заполняет пропуски из данных среды, обрабатывает время
    трансфузий и фенотип.

Автор: Скоков С.О.
Дата создания: Июнь 2026
Версия: 8.0

Изменения в v8.0:
    - Добавлен расширенный алгоритм нормализации групп крови (распознавание A2, A2B, римских цифр)
    - Автоматическое определение имени нормализатора для суффикса выходного файла
    - Улучшено распознавание резус-фактора (Rh+, Rh-, положительная, отрицательная)
    - Сохранены все функции: лист проблем, обработка времени, фенотип
"""

import pandas as pd
import numpy as np
import re
import os
from pathlib import Path
from typing import Optional, Tuple, List, Dict
import tkinter as tk
from tkinter import filedialog

# ===================================================================
# ОПРЕДЕЛЕНИЕ ИМЕНИ ФАЙЛА НОРМАЛИЗАТОРА (автоматически)
# ===================================================================

def get_normalizer_name() -> str:
    """
    Возвращает имя текущего файла нормализатора без расширения .py
    Пример: 'blood_normalizer_v_8_stable'
    """
    current_file = os.path.basename(__file__)
    name_without_ext = os.path.splitext(current_file)[0]
    return name_without_ext


NORMALIZER_NAME = get_normalizer_name()
print(f"📌 Имя нормализатора: {NORMALIZER_NAME}")

# ===================================================================
# КОНФИГУРАЦИЯ
# ===================================================================

CONFIG = {
    # Пути и файлы
    'input_folder': 'enter_data',              # Папка с исходными файлами
    'output_suffix': '_normalize',             # Суффикс для выходного файла (БУДЕТ ЗАМЕНЕН)
    'reports_folder': 'reports',               # Папка для отчетов
    
    # Параметры поиска
    'search_rows': range(0, 20),               # В каких строках искать заголовки
    'required_columns': [                      # Обязательные столбцы
        'Дата', 
        'Объем', 
        'Трансфузионная среда'
    ],
    
    # Режимы работы
    'debug': True,                             # Выводить отладочную информацию
    'create_reports_folder': True,             # Создавать папку для отчетов
    
    # Словари для нормализации (расширенные)
    'rh_positive_keywords': ['+', 'ПОЛОЖ', 'POS', 'ПОЗ', '1', 'ДА', 'RH+', 'РЕЗУС+', 'ПОЛОЖИТЕЛЬНЫЙ', 'ПОЗИТИВ'],
    'rh_negative_keywords': ['-', 'ОТРИЦ', 'NEG', 'НЕГ', '0', 'НЕТ', 'RH-', 'РЕЗУС-', 'ОТРИЦАТЕЛЬНЫЙ', 'НЕГАТИВ'],
    
    # Ключевые слова для определения компонентов
    'component_keywords': {
        'Эритроциты': [
            'эритроцит', 'эм', 'erythrocyte', 
            'эритроцитарная взвесь', 'эритроцитная масса'
        ],
        'Плазма': [
            'плазм', 'plasma', 'свежезамороженная плазма', 
            'сзп', 'свежезамороженная', 'замороженная плазма'
        ],
        'Тромбоциты': [
            'тромбоцит', 'platelet', 'тромбоконцентрат', 
            'тромбоцитарный', 'тромбоцитная масса'
        ],
        'Криопреципитат': [
            'криопреципитат', 'cryoprecipitate', 'крио',
            'криопреципитит'
        ],
    }
}

# ===================================================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ
# ===================================================================

def select_file_from_data_folder() -> Optional[Path]:
    """Открывает диалоговое окно для выбора файла из папки 'enter_data'"""
    root = tk.Tk()
    root.withdraw()
    
    data_folder = Path(CONFIG['input_folder'])
    if not data_folder.exists():
        print(f"📁 Создаю папку '{CONFIG['input_folder']}'...")
        data_folder.mkdir(exist_ok=True)
        print(f"✅ Папка создана. Положите файлы в папку '{CONFIG['input_folder']}'")
        input("Нажмите Enter после добавления файлов...")
    
    print(f"\n📂 Открываю папку: {data_folder.absolute()}")
    
    file_path = filedialog.askopenfilename(
        title="Выберите файл с трансфузиями",
        initialdir=data_folder.absolute(),
        filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
    )
    
    root.destroy()
    
    if not file_path:
        print("❌ Файл не выбран. Выход...")
        return None
    
    return Path(file_path)


def find_header_row(df: pd.DataFrame, required_cols: List[str]) -> Tuple[Optional[int], Optional[Dict]]:
    """Находит строку с заголовками в DataFrame"""
    for idx in CONFIG['search_rows']:
        if idx >= len(df):
            break
        
        row_values = []
        for val in df.iloc[idx]:
            if pd.isna(val):
                row_values.append('')
            else:
                row_values.append(str(val).strip())
        
        found_cols = {}
        for req in required_cols:
            matches = [col for col in row_values if req.lower() in col.lower()]
            if matches:
                found_cols[req] = matches[0]
            else:
                break
        else:
            col_mapping = {}
            for i, val in enumerate(row_values):
                for req, col_name in found_cols.items():
                    if val == col_name:
                        col_mapping[req] = i
            return idx, col_mapping
    
    return None, None


# ===================================================================
# ФУНКЦИИ ДЛЯ ПРОВЕРКИ И ОЧИСТКИ
# ===================================================================

def roman_to_abo(roman: str) -> Optional[str]:
    """Преобразует римскую цифру в группу крови"""
    roman = roman.upper().strip()
    mapping = {
        'I': 'O',
        'II': 'A',
        'III': 'B',
        'IV': 'AB'
    }
    return mapping.get(roman)


def extract_from_parentheses(text: str) -> Optional[str]:
    """Извлекает римскую цифру из скобок и преобразует в группу"""
    match = re.search(r'\(([IVX]+)\)', text, re.IGNORECASE)
    if match:
        return roman_to_abo(match.group(1))
    return None


def is_garbage_blood_group(abo_str: any) -> bool:
    """
    Проверяет, является ли строка "мусором" (невалидной группой крови)
    
    Признаки мусора:
        - Более 3 цифр подряд (например, "1310")
        - Длинные буквенно-цифровые коды
    """
    if pd.isna(abo_str):
        return True
    
    abo = str(abo_str).upper().strip()
    
    # Более 3 цифр подряд - мусор
    if re.search(r'\d{4,}', abo):
        return True
    
    # Длиннее 15 символов - мусор
    if len(abo) > 15:
        return True
    
    return False


def clean_blood_group_string(abo_str: str) -> str:
    """Очищает строку группы крови от шума (скобки, пробелы, римские цифры)"""
    abo = str(abo_str).upper()
    # Убираем скобки и пробелы
    abo = re.sub(r'[\(\)\[\]\{\}\s]', '', abo)
    # Убираем римские цифры
    abo = re.sub(r'\b[IVX]+\b', '', abo)
    # Убираем Rh и резус
    abo = re.sub(r'RH[+\-]?', '', abo)
    abo = re.sub(r'РЕЗУС[+\-]?', '', abo)
    return abo


# ===================================================================
# РАСШИРЕННЫЕ ФУНКЦИИ НОРМАЛИЗАЦИИ ГРУПП КРОВИ
# ===================================================================

def normalize_blood_group_advanced(abo_str: any, verification_value: any = None) -> Optional[str]:
    """
    Расширенная нормализация группы крови
    
    Распознает:
        - A2B, А2В, A2B (IV)
        - A2, А2, A2 (II)
        - AB, АВ, AB (IV)
        - A, А, A (II)
        - B, В, B (III)
        - O, О, 0, O (I), 0 (I)
        - Римские цифры в скобках
        - Цифра 2 → A2
        - Мусор → берем из verification_value
    """
    if pd.isna(abo_str):
        return None
    
    value_str = str(abo_str).strip()
    
    # ========== 1. Прямое сопоставление ==========
    direct_mapping = {
        'a2b': 'A2B', 'а2в': 'A2B', 'a2b(iv)': 'A2B', 'а2в(iv)': 'A2B',
        'a2b (iv)': 'A2B', 'а2в (iv)': 'A2B',
        'a2': 'A2', 'а2': 'A2', 'a2(ii)': 'A2', 'а2(ii)': 'A2',
        'a2 (ii)': 'A2', 'а2 (ii)': 'A2',
        '2': 'A2', '4': 'O',
        'ab': 'AB', 'ав': 'AB', 'ab(iv)': 'AB', 'ав(iv)': 'AB',
        'ab (iv)': 'AB', 'ав (iv)': 'AB',
        'a': 'A', 'а': 'A', 'a(ii)': 'A', 'а(ii)': 'A',
        'a (ii)': 'A', 'а (ii)': 'A',
        'b': 'B', 'в': 'B', 'b(iii)': 'B', 'в(iii)': 'B',
        'b (iii)': 'B', 'в (iii)': 'B',
        'o': 'O', 'о': 'O', '0': 'O', 'o(i)': 'O', 'о(i)': 'O',
        '0(i)': 'O', 'o (i)': 'O', 'о (i)': 'O', '0 (i)': 'O',
    }
    
    value_lower = value_str.lower()
    for pattern, result in direct_mapping.items():
        if value_lower == pattern or value_lower.startswith(pattern):
            return result
    
    # ========== 2. Извлечение из скобок ==========
    abo_from_parentheses = extract_from_parentheses(value_str)
    if abo_from_parentheses:
        return abo_from_parentheses
    
    # ========== 3. Поиск ключевых слов ==========
    value_upper = value_str.upper()
    
    if 'A2B' in value_upper or 'А2В' in value_upper:
        return 'A2B'
    if 'A2' in value_upper or 'А2' in value_upper:
        return 'A2'
    if 'AB' in value_upper or 'АВ' in value_upper:
        if 'A2B' not in value_upper and 'А2В' not in value_upper:
            return 'AB'
    if 'A' in value_upper or 'А' in value_upper:
        if 'AB' not in value_upper and 'АВ' not in value_upper:
            return 'A'
    if 'B' in value_upper or 'В' in value_upper:
        if 'AB' not in value_upper and 'АВ' not in value_upper:
            return 'B'
    if 'O' in value_upper or 'О' in value_upper or '0' in value_str:
        return 'O'
    
    # ========== 4. Мусор → берем из проверочного столбца ==========
    if is_garbage_blood_group(value_str):
        if verification_value and not pd.isna(verification_value):
            result = normalize_blood_group_advanced(verification_value)
            if result:
                return result
    
    return None


def normalize_rh_advanced(rh_str: any) -> Optional[str]:
    """
    Расширенная нормализация резус-фактора
    
    Распознает:
        - +, Rh+, Rh +, (+)
        - -, Rh-, Rh -, (-)
        - положительная, положительный
        - отрицательная, отрицательный
    """
    if pd.isna(rh_str):
        return None
    
    rh = str(rh_str).upper().strip()
    
    # Убираем скобки и пробелы
    clean_rh = re.sub(r'[\(\s\)]', '', rh)
    
    # Положительный резус
    positive_patterns = ['+', 'RH+', 'ПОЛОЖ', 'ПОЛОЖИТЕЛЬН', 'POS', 'ПОЗ']
    for pattern in positive_patterns:
        if pattern in clean_rh:
            return '+'
    
    # Отрицательный резус
    negative_patterns = ['-', 'RH-', 'ОТРИЦ', 'ОТРИЦАТЕЛЬН', 'NEG', 'НЕГ']
    for pattern in negative_patterns:
        if pattern in clean_rh:
            return '-'
    
    return None


def normalize_blood_group_only(abo_str: any, verification_value: any = None) -> Optional[str]:
    """
    Нормализует группу крови (обертка над расширенной функцией)
    """
    return normalize_blood_group_advanced(abo_str, verification_value)


def normalize_rh_only(rh_str: any) -> Optional[str]:
    """
    Нормализует резус-фактор (обертка над расширенной функцией)
    """
    return normalize_rh_advanced(rh_str)


# ===================================================================
# ФУНКЦИИ ОПРЕДЕЛЕНИЯ КОМПОНЕНТОВ
# ===================================================================

def get_component_type(component_str: any) -> str:
    """Определяет тип компонента крови"""
    if pd.isna(component_str):
        return 'Не определено'
    
    comp = str(component_str).lower()
    
    if comp == 'сзп':
        return 'Плазма'
    
    for comp_type, keywords in CONFIG['component_keywords'].items():
        for keyword in keywords:
            if keyword in comp:
                return comp_type
    
    return 'Не определено'


# ===================================================================
# ФУНКЦИИ ДЛЯ ОБРАБОТКИ ВРЕМЕНИ
# ===================================================================

def excel_time_to_time_str(excel_time: any) -> str:
    """Преобразует время Excel в строку ЧЧ:ММ:СС"""
    if pd.isna(excel_time):
        return ''
    
    try:
        time_float = float(excel_time)
        total_seconds = int(time_float * 24 * 3600)
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    except:
        return str(excel_time)


def calculate_duration_minutes(start_time: any, end_time: any) -> Optional[float]:
    """Рассчитывает длительность трансфузии в минутах"""
    if pd.isna(start_time) or pd.isna(end_time):
        return None
    
    try:
        start_float = float(start_time)
        end_float = float(end_time)
        
        if end_float < start_float:
            end_float += 1.0
        
        duration_days = end_float - start_float
        return round(duration_days * 24 * 60, 1)
    except:
        return None


# ===================================================================
# ФУНКЦИИ ДЛЯ ОБРАБОТКИ ФЕНОТИПА
# ===================================================================

def extract_short_phenotype(phenotype_str: any) -> str:
    """Извлекает краткий фенотип (паттерны C/c и E/e)"""
    if pd.isna(phenotype_str):
        return ''
    
    pheno = str(phenotype_str)
    
    c_match = re.search(r'[Cc]{2}', pheno)
    e_match = re.search(r'[Ee]{2}', pheno)
    
    c_part = c_match.group(0) if c_match else ''
    e_part = e_match.group(0) if e_match else ''
    
    return f"{c_part}{e_part}" if (c_part or e_part) else ''


# ===================================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ===================================================================

def main() -> None:
    """Главная функция скрипта."""
    print("🩸 ЗАПУСК ПРЕПРОЦЕССОРА ТРАНСФУЗИЙ v8.0")
    print("="*60)
    print("\n📌 Новое в v8.0:")
    print("   • Расширенный алгоритм нормализации групп крови (распознавание A2, A2B, римских цифр)")
    print("   • Автоматическое определение имени нормализатора для суффикса файла")
    print("   • Улучшено распознавание резус-фактора")
    print(f"   • Имя нормализатора: {NORMALIZER_NAME}")
    
    # Создаем папку для отчетов
    if CONFIG['create_reports_folder']:
        reports_dir = Path(CONFIG['reports_folder'])
        reports_dir.mkdir(exist_ok=True)
        print(f"\n📁 Папка для отчетов: {reports_dir}")
    
    # Выбираем файл
    input_path = select_file_from_data_folder()
    if not input_path:
        return
    
    print(f"\n📂 Загружаем файл: {input_path.name}")
    
    # Загружаем Excel файл
    try:
        df_raw = pd.read_excel(input_path, header=None, dtype=str)
    except Exception as e:
        print(f"❌ Ошибка загрузки файла: {e}")
        print("   Проверьте установку: pip install openpyxl")
        return
    
    print(f"   Размер: {df_raw.shape[0]} строк × {df_raw.shape[1]} столбцов")
    
    # Ищем строку с заголовками
    print("\n🔍 Ищем строку с заголовками...")
    header_row, col_mapping = find_header_row(df_raw, CONFIG['required_columns'])
    
    if header_row is None:
        print("❌ Не удалось найти строку с заголовками!")
        print(f"   Искали столбцы: {CONFIG['required_columns']}")
        return
    
    print(f"✅ Заголовки найдены в строке {header_row + 1}")
    
    # Создаем DataFrame с правильными заголовками
    headers = []
    for val in df_raw.iloc[header_row]:
        if pd.isna(val):
            headers.append('')
        else:
            headers.append(str(val).strip())
    
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = headers
    
    print(f"📊 Загружено {len(data)} записей")
    
    # Находим нужные колонки
    print("\n🔎 Поиск необходимых колонок...")
    
    abo_patient_col = None
    rh_patient_col = None
    abo_env_col = None
    rh_env_col = None
    component_col = None
    time_from_col = None
    time_to_col = None
    phenotype_col = None
    
    for col in data.columns:
        if not col:
            continue
        col_lower = col.lower()
        if ('ab0' in col_lower or 'abo' in col_lower):
            if 'пац' in col_lower:
                abo_patient_col = col
            elif 'сред' in col_lower:
                abo_env_col = col
        elif 'rh' in col_lower:
            if 'пац' in col_lower:
                rh_patient_col = col
            elif 'сред' in col_lower:
                rh_env_col = col
        elif 'трансфузионная среда' in col_lower:
            component_col = col
        elif col_lower == 'с':
            time_from_col = col
        elif col_lower == 'по':
            time_to_col = col
        elif 'фенотип' in col_lower:
            phenotype_col = col
    
    if CONFIG['debug']:
        print(f"   Пациент: ABO={abo_patient_col}, RH={rh_patient_col}")
        print(f"   Среда: ABO={abo_env_col}, RH={rh_env_col}")
        print(f"   Компонент: {component_col}")
        print(f"   Время начала (с): {time_from_col}")
        print(f"   Время окончания (по): {time_to_col}")
        print(f"   Фенотип: {phenotype_col}")
    
    # ========== НОРМАЛИЗАЦИЯ ДАННЫХ ПАЦИЕНТА (с расширенным алгоритмом) ==========
    print("\n🩸 Нормализация групп крови пациента (расширенный алгоритм)...")
    
    patient_blood_norm = []
    patient_rh_norm = []
    patient_full = []
    patient_problems = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_patient_col) if abo_patient_col else None
        rh = row.get(rh_patient_col) if rh_patient_col else None
        donor_abo = row.get(abo_env_col) if abo_env_col else None
        
        problem = ''
        blood_norm = None
        
        # Проверяем на мусор в группе пациента
        if is_garbage_blood_group(abo):
            # Мусор - пробуем взять из среды
            blood_norm = normalize_blood_group_only(donor_abo)
            if blood_norm:
                problem = f"Мусор в группе пациента ('{abo}'), группа взята из среды: {blood_norm}"
            else:
                blood_norm = normalize_blood_group_only(abo)
                if not blood_norm:
                    problem = f"Невалидная группа пациента (мусор): '{abo}'"
        else:
            # Расширенная нормализация с возможностью использования проверочного столбца
            blood_norm = normalize_blood_group_advanced(abo, donor_abo)
            if not blood_norm and donor_abo:
                blood_norm = normalize_blood_group_advanced(donor_abo)
                if blood_norm:
                    problem = f"Группа не определена, взята из среды: {blood_norm}"
            elif not blood_norm:
                problem = f"Не определена группа пациента: '{abo}'"
        
        rh_norm = normalize_rh_advanced(rh)
        if rh_norm is None and rh:
            problem += f" [не определен резус: '{rh}']"
        
        patient_blood_norm.append(blood_norm)
        patient_rh_norm.append(rh_norm)
        
        if blood_norm and rh_norm:
            patient_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm:
            patient_full.append(f"{blood_norm}?")
            if not problem:
                problem = "Не определен резус пациента"
        else:
            patient_full.append(None)
        
        patient_problems.append(problem if problem else '')
    
    # ========== НОРМАЛИЗАЦИЯ ДАННЫХ СРЕДЫ ==========
    print("🩸 Нормализация групп крови среды...")
    
    env_blood_norm = []
    env_rh_norm = []
    env_full = []
    env_problems = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_env_col) if abo_env_col else None
        rh = row.get(rh_env_col) if rh_env_col else None
        
        problem = ''
        blood_norm = normalize_blood_group_advanced(abo)
        if not blood_norm and abo:
            problem = f"Не определена группа донора: '{abo}'"
        
        rh_norm = normalize_rh_advanced(rh)
        if rh_norm is None and rh:
            problem += f" [не определен резус донора: '{rh}']"
        
        env_blood_norm.append(blood_norm)
        env_rh_norm.append(rh_norm)
        
        if blood_norm and rh_norm:
            env_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm:
            env_full.append(f"{blood_norm}?")
        else:
            env_full.append(None)
        
        env_problems.append(problem if problem else '')
    
    # ========== ЗАПОЛНЕНИЕ ПРОПУСКОВ ==========
    print("\n🔄 Заполнение пропусков из данных среды...")
    
    patient_final_full = []
    patient_final_source = []
    final_problems = []
    
    for i in range(len(data)):
        patient_val = patient_full[i]
        source = 'patient'
        problem = patient_problems[i]
        
        if patient_val is None or '?' in str(patient_val):
            env_val = env_full[i]
            if env_val and '?' not in str(env_val):
                patient_val = env_val
                source = 'environment'
                problem = f"ЗАМЕНА: {problem}" if problem else "ЗАМЕНА: группа взята из среды"
        
        patient_final_full.append(patient_val)
        patient_final_source.append(source)
        final_problems.append(problem if problem else '')
    
    # ========== ОПРЕДЕЛЕНИЕ ТИПОВ КОМПОНЕНТОВ ==========
    print("🔬 Определение типов компонентов...")
    
    component_types = []
    for idx, row in data.iterrows():
        comp = row.get(component_col) if component_col else None
        comp_type = get_component_type(comp)
        component_types.append(comp_type)
    
    # ========== ОБРАБОТКА ВРЕМЕНИ ==========
    print("⏰ Обработка временных меток...")
    
    time_from_str = []
    time_to_str = []
    durations = []
    
    if time_from_col and time_to_col:
        for idx, row in data.iterrows():
            start = row.get(time_from_col)
            end = row.get(time_to_col)
            
            from_str = excel_time_to_time_str(start)
            to_str = excel_time_to_time_str(end)
            duration = calculate_duration_minutes(start, end)
            
            time_from_str.append(from_str)
            time_to_str.append(to_str)
            durations.append(duration)
        
        data['Время_начала'] = time_from_str
        data['Время_окончания'] = time_to_str
        data['Длительность_минуты'] = durations
        print(f"   ✅ Добавлены столбцы: Время_начала, Время_окончания, Длительность_минуты")
    else:
        print("   ⚠️ Столбцы со временем не найдены ('с' и 'по')")
    
    # ========== ОБРАБОТКА ФЕНОТИПА ==========
    print("🧬 Нормализация фенотипа...")
    
    if phenotype_col:
        short_phenotypes = []
        for idx, row in data.iterrows():
            pheno = row.get(phenotype_col)
            short_pheno = extract_short_phenotype(pheno)
            short_phenotypes.append(short_pheno)
        
        data['Краткий_фенотип'] = short_phenotypes
        print(f"   ✅ Добавлен столбец: Краткий_фенотип")
    
    # ========== ДОБАВЛЕНИЕ НОВЫХ СТОЛБЦОВ ==========
    print("\n📝 Добавление новых столбцов...")
    
    data['Blood_Group_Patient_Norm'] = patient_blood_norm
    data['Rh_Patient_Norm'] = patient_rh_norm
    data['Blood_Group_Patient_Full'] = patient_final_full
    data['Blood_Group_Source'] = patient_final_source
    data['Проблема_пациента'] = final_problems
    
    data['Blood_Group_Env_Norm'] = env_blood_norm
    data['Rh_Env_Norm'] = env_rh_norm
    data['Blood_Group_Env_Full'] = env_full
    data['Проблема_среды'] = env_problems
    
    data['Component_Type'] = component_types
    
    # ========== СОХРАНЕНИЕ РЕЗУЛЬТАТОВ (с суффиксом от имени нормализатора) ==========
    # Используем имя нормализатора для суффикса
    output_path = input_path.parent / f"{input_path.stem}_by_{NORMALIZER_NAME}.xlsx"
    
    print(f"\n💾 Сохраняем обработанный файл: {output_path}")
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Основной лист
            data.to_excel(writer, sheet_name='Все_трансфузии', index=False)
            
            # Лист с проблемами нормализации
            problems_all = data[(data['Проблема_пациента'] != '') | (data['Проблема_среды'] != '')]
            if len(problems_all) > 0:
                problem_cols = ['Проблема_пациента', 'Проблема_среды', 
                               'Blood_Group_Patient_Full', 'Blood_Group_Env_Full',
                               'Component_Type']
                existing_cols = [c for c in problem_cols if c in problems_all.columns]
                problems_all[existing_cols].to_excel(writer, sheet_name='Проблемы_нормализации', index=False)
                print(f"   ⚠️ Записей с проблемами: {len(problems_all)}")
            
            # Отдельный лист только с эритроцитами
            erythro_mask = data['Component_Type'] == 'Эритроциты'
            if erythro_mask.any():
                data[erythro_mask].to_excel(writer, sheet_name='Эритроциты', index=False)
    except Exception as e:
        print(f"❌ Ошибка при сохранении: {e}")
        return
    
    # ========== ФИНАЛЬНАЯ СТАТИСТИКА ==========
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ СТАТИСТИКА:")
    print(f"   Всего записей: {len(data)}")
    
    problems_count = len(data[(data['Проблема_пациента'] != '') | (data['Проблема_среды'] != '')])
    print(f"   Проблемных записей: {problems_count} ({problems_count/len(data)*100:.1f}%)")
    
    print(f"\n   📦 Типы компонентов:")
    for comp_type, count in data['Component_Type'].value_counts().items():
        print(f"      {comp_type}: {count} ({count/len(data)*100:.1f}%)")
    
    # Статистика по эритроцитам
    erythro_data = data[data['Component_Type'] == 'Эритроциты']
    if len(erythro_data) > 0:
        print(f"\n   🩸 Группы крови (только эритроциты, n={len(erythro_data)}):")
        for blood, count in erythro_data['Blood_Group_Patient_Full'].value_counts().sort_index().items():
            if blood and '?' not in str(blood):
                print(f"      {blood}: {count} ({count/len(erythro_data)*100:.1f}%)")
    
    print(f"\n   📍 Источники групп крови:")
    for source, count in data['Blood_Group_Source'].value_counts().items():
        print(f"      {source}: {count} ({count/len(data)*100:.1f}%)")
    
    print("\n" + "="*60)
    print("✅ ПРЕПРОЦЕССИНГ ЗАВЕРШЕН")
    print(f"📁 Результат: {output_path.name}")
    print(f"📁 Папка с отчетами: {CONFIG['reports_folder']}/")


if __name__ == "__main__":
    main()