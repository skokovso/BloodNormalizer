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
Версия: 7.0

Изменения в v2.7:
    - Добавлена проверка на цифровой мусор (более 3 цифр подряд)
    - При цифровом мусоре в группе пациента - берем группу из среды
    - Добавлен лист 'Проблемы_нормализации' в Excel отчет
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import Optional, Tuple, List, Dict
import tkinter as tk
from tkinter import filedialog

# ===================================================================
# КОНФИГУРАЦИЯ
# ===================================================================

CONFIG = {
    # Пути и файлы
    'input_folder': 'enter_data',              # Папка с исходными файлами
    'output_suffix': '_normalize',          # Суффикс для выходного файла
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
    
    # Словари для нормализации
    'rh_positive_keywords': ['+', 'ПОЛОЖ', 'POS', 'ПОЗ', '1', 'ДА', 'RH+', 'РЕЗУС+'],
    'rh_negative_keywords': ['-', 'ОТРИЦ', 'NEG', 'НЕГ', '0', 'НЕТ', 'RH-', 'РЕЗУС-'],
    
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

def is_garbage_blood_group(abo_str: any) -> bool:
    """
    Проверяет, является ли строка "мусором" (невалидной группой крови)
    
    Признаки мусора:
        - Более 3 цифр подряд (например, "1310")
        - Длинные буквенно-цифровые коды
        - Содержит недопустимые символы (кроме A,B,O,0,2,+,-,/,А,В,О)
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
    
    # Разрешенные символы для группы крови
    allowed_pattern = r'^[ABO02АВО2\+\-\/\(\)\sIVX]+$'
    if not re.match(allowed_pattern, abo):
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
# ФУНКЦИИ НОРМАЛИЗАЦИИ ГРУПП КРОВИ
# ===================================================================

def normalize_blood_group_only(abo_str: any) -> Optional[str]:
    """
    Нормализует только группу крови (без резус-фактора)
    
    Правила:
        - O: символы O, О, 0 (ноль)
        - A: символ A/А без B и без 2
        - B: символ B/В без A и без 2
        - AB: символы A и B без 2
        - A2: символ A с цифрой 2 после него
        - A2B: символы A и B с цифрой 2 после A
    """
    if pd.isna(abo_str):
        return None
    
    abo = clean_blood_group_string(str(abo_str))
    
    if not abo:
        return None
    
    has_A = 'А' in abo or 'A' in abo
    has_B = 'В' in abo or 'B' in abo
    has_2 = '2' in abo
    
    if not has_2:
        a2_pattern = r'[АA]2'
        has_2 = bool(re.search(a2_pattern, abo))
    
    if has_2 and has_A and has_B:
        return 'A2B'
    elif has_2 and has_A and not has_B:
        return 'A2'
    elif has_A and has_B and not has_2:
        return 'AB'
    elif has_A and not has_B and not has_2:
        return 'A'
    elif has_B and not has_A and not has_2:
        return 'B'
    elif (not has_A and not has_B) or ('О' in abo or 'O' in abo or '0' in abo):
        return 'O'
    else:
        return None


def normalize_rh_only(rh_str: any) -> Optional[str]:
    """Нормализует резус-фактор"""
    if pd.isna(rh_str):
        return None
    
    rh = str(rh_str).upper().strip()
    
    if '+' in rh:
        return '+'
    elif '-' in rh:
        return '-'
    
    if any(word in rh for word in CONFIG['rh_positive_keywords']):
        return '+'
    elif any(word in rh for word in CONFIG['rh_negative_keywords']):
        return '-'
    
    return None


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
    print("🩸 ЗАПУСК ПРЕПРОЦЕССОРА ТРАНСФУЗИЙ v2.7")
    print("="*60)
    print("\n📌 Новое в v2.7:")
    print("   • Проверка на цифровой мусор (>3 цифр) - берем группу из среды")
    print("   • Отдельный лист 'Проблемы_нормализации'")
    
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
    
    # ========== НОРМАЛИЗАЦИЯ ДАННЫХ ПАЦИЕНТА ==========
    print("\n🩸 Нормализация групп крови пациента...")
    
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
            blood_norm = normalize_blood_group_only(abo)
            if not blood_norm and donor_abo:
                blood_norm = normalize_blood_group_only(donor_abo)
                if blood_norm:
                    problem = f"Группа не определена, взята из среды: {blood_norm}"
            elif not blood_norm:
                problem = f"Не определена группа пациента: '{abo}'"
        
        rh_norm = normalize_rh_only(rh)
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
        blood_norm = normalize_blood_group_only(abo)
        if not blood_norm and abo:
            problem = f"Не определена группа донора: '{abo}'"
        
        rh_norm = normalize_rh_only(rh)
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
    
    # ========== СОХРАНЕНИЕ РЕЗУЛЬТАТОВ ==========
    output_path = input_path.parent / f"{input_path.stem}{CONFIG['output_suffix']}.xlsx"
    
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