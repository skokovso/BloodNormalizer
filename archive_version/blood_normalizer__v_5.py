#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль нормализации данных трансфузий крови v5.0
=================================================

Модуль для предобработки Excel-файлов с трансфузиями крови.
Может использоваться как самостоятельный скрипт или импортироваться в другие модули.

Назначение:
    Подготовка данных для расчета оптимального запаса компонентов крови.
    Нормализует группы крови с учетом подгрупп A2 и A2B, определяет типы
    компонентов, заполняет пропуски из данных среды, обрабатывает время
    трансфузий и фенотип.

Автор: Скоков С.О.
Дата создания: Июнь 2026
Версия: 5.0

Изменения в v5.0:
    - Улучшена проверка валидности группы крови (римские цифры, скобки)
    - Добавлен fallback на группу донора при невалидных данных
    - Расширены ключевые слова для распознавания резус-фактора
    - Добавлен столбец 'Проблема_нормализации' для отслеживания ошибок
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
    'input_folder': 'Enter_data',
    'output_suffix': '_preprocessed',
    'reports_folder': 'reports',
    'search_rows': range(0, 20),
    'required_columns': ['Дата', 'Объем', 'Трансфузионная среда'],
    'debug': True,
    'create_reports_folder': True,
    
    # Ключевые слова для резус-фактора
    'rh_positive_keywords': [
        '+', 'ПОЛОЖ', 'POS', 'ПОЗ', '1', 'ДА',
        'RH+', 'RH ПОЛОЖ', 'РЕЗУС+', 'РЕЗУС ПОЛОЖ',
        'ПОЛОЖИТЕЛЬНЫЙ', 'POSITIVE', 'ПОЗИТИВ'
    ],
    'rh_negative_keywords': [
        '-', 'ОТРИЦ', 'NEG', 'НЕГ', '0', 'НЕТ',
        'RH-', 'RH ОТРИЦ', 'РЕЗУС-', 'РЕЗУС ОТРИЦ',
        'ОТРИЦАТЕЛЬНЫЙ', 'NEGATIVE', 'НЕГАТИВ'
    ],
    
    # Ключевые слова для компонентов
    'component_keywords': {
        'Эритроциты': ['эритроцит', 'эм', 'erythrocyte', 'эритроцитарная взвесь', 'эритроцитная масса'],
        'Плазма': ['плазм', 'plasma', 'свежезамороженная плазма', 'сзп', 'свежезамороженная', 'замороженная плазма'],
        'Тромбоциты': ['тромбоцит', 'platelet', 'тромбоконцентрат', 'тромбоцитарный', 'тромбоцитная масса'],
        'Криопреципитат': ['криопреципитат', 'cryoprecipitate', 'крио', 'криопреципитит'],
    }
}


# ===================================================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ
# ===================================================================

def select_file_from_data_folder() -> Optional[Path]:
    """Открывает диалоговое окно для выбора файла из папки Enter_data"""
    root = tk.Tk()
    root.withdraw()
    
    data_folder = Path(CONFIG['input_folder'])
    if not data_folder.exists():
        data_folder.mkdir(exist_ok=True)
        print(f"📁 Создана папка '{CONFIG['input_folder']}'")
        input("Нажмите Enter после добавления файлов...")
    
    print(f"\n📂 Открываю папку: {data_folder.absolute()}")
    
    file_path = filedialog.askopenfilename(
        title="Выберите файл с трансфузиями",
        initialdir=data_folder.absolute(),
        filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
    )
    
    root.destroy()
    
    if not file_path:
        print("❌ Файл не выбран.")
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
# ФУНКЦИИ НОРМАЛИЗАЦИИ ГРУПП КРОВИ
# ===================================================================

def clean_blood_group_string(abo_str: str) -> str:
    """Очищает строку группы крови от шума (скобки, пробелы, римские цифры)"""
    # Убираем скобки и пробелы
    abo = re.sub(r'[\(\)\[\]\{\}\s]', '', abo_str)
    # Убираем римские цифры (I, II, III, IV, V, VI и т.д.)
    abo = re.sub(r'\b[IVX]+\b', '', abo, flags=re.IGNORECASE)
    return abo.upper().strip()


def is_valid_blood_group(abo_str: any) -> bool:
    """
    Проверяет, является ли строка допустимым обозначением группы крови
    
    Допустимые форматы:
        - O, A, B, AB
        - A2, A2B
        - 0 (ноль)
        - С резусом: O+, A-
        - Русские буквы: А, В, О
    
    Недопустимые (мусор):
        - Длинные цифровые коды (1310, 123456)
        - Длинные буквенно-цифровые коды
        - Символы кроме A,B,O,0,2,+,-,/,А,В,О
    """
    if pd.isna(abo_str):
        return False
    
    abo = str(abo_str).upper().strip()
    
    # Слишком длинная строка - мусор
    if len(abo) > 10:
        return False
    
    # Только цифры и длиннее 2 символов - мусор
    if abo.isdigit() and len(abo) > 2:
        return False
    
    # Очищаем от скобок и римских цифр
    abo_clean = clean_blood_group_string(abo)
    
    if not abo_clean:
        return False
    
    # Разрешенные символы
    allowed_pattern = r'^[ABO02АВО2\+\-\/]+$'
    if not re.match(allowed_pattern, abo_clean):
        return False
    
    return True


def normalize_blood_group_only(abo_str: any) -> Optional[str]:
    """Нормализует только группу крови (без резус-фактора)"""
    if pd.isna(abo_str):
        return None
    
    abo = str(abo_str).upper().strip()
    abo = clean_blood_group_string(abo)
    
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


def normalize_blood_group_with_fallback(abo_str: any, fallback_abo_str: any = None) -> Optional[str]:
    """Нормализует группу крови с подстановкой из запасного поля при невалидных данных"""
    if is_valid_blood_group(abo_str):
        result = normalize_blood_group_only(abo_str)
        if result:
            return result
    
    if fallback_abo_str and is_valid_blood_group(fallback_abo_str):
        result = normalize_blood_group_only(fallback_abo_str)
        if result:
            return result
    
    return None


def normalize_rh_only(rh_str: any) -> Optional[str]:
    """Нормализует резус-фактор"""
    if pd.isna(rh_str):
        return None
    
    rh = str(rh_str).upper().strip()
    
    if '+' in rh:
        return '+'
    if '-' in rh:
        return '-'
    
    for keyword in CONFIG['rh_positive_keywords']:
        if keyword in rh:
            return '+'
    
    for keyword in CONFIG['rh_negative_keywords']:
        if keyword in rh:
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
# ФУНКЦИИ ДЛЯ ОБРАБОТКИ ВРЕМЕНИ И ФЕНОТИПА
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
    """Рассчитывает длительность в минутах"""
    if pd.isna(start_time) or pd.isna(end_time):
        return None
    try:
        start_float = float(start_time)
        end_float = float(end_time)
        if end_float < start_float:
            end_float += 1.0
        return round((end_float - start_float) * 24 * 60, 1)
    except:
        return None


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

def main():
    print("🩸 ЗАПУСК ПРЕПРОЦЕССОРА ТРАНСФУЗИЙ v5.0")
    print("="*60)
    print("\n📌 Улучшения v5.0:")
    print("   • Улучшена проверка валидности группы крови")
    print("   • Fallback на группу донора при невалидных данных")
    print("   • Расширены ключевые слова для резус-фактора")
    
    if CONFIG['create_reports_folder']:
        reports_dir = Path(CONFIG['reports_folder'])
        reports_dir.mkdir(exist_ok=True)
        print(f"\n📁 Папка для отчетов: {reports_dir}")
    
    input_path = select_file_from_data_folder()
    if not input_path:
        return
    
    print(f"\n📂 Загружаем файл: {input_path.name}")
    
    try:
        df_raw = pd.read_excel(input_path, header=None, dtype=str)
    except Exception as e:
        print(f"❌ Ошибка загрузки: {e}")
        return
    
    print(f"   Размер: {df_raw.shape[0]} строк × {df_raw.shape[1]} столбцов")
    
    print("\n🔍 Ищем строку с заголовками...")
    header_row, col_mapping = find_header_row(df_raw, CONFIG['required_columns'])
    
    if header_row is None:
        print("❌ Не удалось найти строку с заголовками!")
        return
    
    print(f"✅ Заголовки найдены в строке {header_row + 1}")
    
    headers = []
    for val in df_raw.iloc[header_row]:
        headers.append(str(val).strip() if not pd.isna(val) else '')
    
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = headers
    
    print(f"📊 Загружено {len(data)} записей")
    
    # Поиск колонок
    print("\n🔎 Поиск необходимых колонок...")
    
    abo_patient_col = rh_patient_col = None
    abo_env_col = rh_env_col = None
    component_col = time_from_col = time_to_col = None
    phenotype_col = patient_name_col = patient_id_col = None
    
    for col in data.columns:
        if not col:
            continue
        col_lower = col.lower()
        if 'ab0' in col_lower or 'abo' in col_lower:
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
        elif 'фио' in col_lower and 'пациент' in col_lower:
            patient_name_col = col
        elif 'номер' in col_lower and 'пациент' in col_lower:
            patient_id_col = col
    
    if CONFIG['debug']:
        print(f"   Пациент: ABO={abo_patient_col}, RH={rh_patient_col}")
        print(f"   Среда: ABO={abo_env_col}, RH={rh_env_col}")
        print(f"   Компонент: {component_col}")
    
    # ========== НОРМАЛИЗАЦИЯ ==========
    print("\n🩸 Нормализация групп крови...")
    
    patient_blood_norm = []
    patient_rh_norm = []
    patient_full = []
    patient_problems = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_patient_col) if abo_patient_col else None
        rh = row.get(rh_patient_col) if rh_patient_col else None
        donor_abo = row.get(abo_env_col) if abo_env_col else None
        
        blood_norm = normalize_blood_group_with_fallback(abo, donor_abo)
        rh_norm = normalize_rh_only(rh)
        
        patient_blood_norm.append(blood_norm)
        patient_rh_norm.append(rh_norm)
        
        problem = ''
        if blood_norm and rh_norm:
            patient_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm and rh_norm is None:
            patient_full.append(f"{blood_norm}?")
            problem = 'Не определен резус пациента'
        elif blood_norm:
            patient_full.append(f"{blood_norm}?")
            problem = 'Не определен резус пациента'
        else:
            patient_full.append(None)
            if abo and not is_valid_blood_group(abo):
                problem = f"Невалидная группа пациента (мусор): {str(abo)[:50]}"
            else:
                problem = 'Не определена группа пациента'
        
        patient_problems.append(problem)
    
    # Нормализация среды
    print("🩸 Нормализация групп крови среды...")
    
    env_blood_norm = []
    env_rh_norm = []
    env_full = []
    env_problems = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_env_col) if abo_env_col else None
        rh = row.get(rh_env_col) if rh_env_col else None
        
        blood_norm = normalize_blood_group_only(abo)
        rh_norm = normalize_rh_only(rh)
        
        env_blood_norm.append(blood_norm)
        env_rh_norm.append(rh_norm)
        
        problem = ''
        if blood_norm and rh_norm:
            env_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm and rh_norm is None:
            env_full.append(f"{blood_norm}?")
            problem = 'Не определен резус донора'
        elif blood_norm:
            env_full.append(f"{blood_norm}?")
            problem = 'Не определен резус донора'
        else:
            env_full.append(None)
            problem = 'Не определена группа донора'
        
        env_problems.append(problem)
    
    # Заполнение пропусков
    print("\n🔄 Заполнение пропусков...")
    
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
                problem = f"ЗАМЕНА: группа взята из среды (было: {problem})" if problem else 'ЗАМЕНА: группа взята из среды'
        
        patient_final_full.append(patient_val)
        patient_final_source.append(source)
        final_problems.append(problem if problem else '')
    
    # Типы компонентов
    print("🔬 Определение типов компонентов...")
    component_types = [get_component_type(row.get(component_col)) for _, row in data.iterrows()]
    
    # Время
    print("⏰ Обработка времени...")
    if time_from_col and time_to_col:
        data['Время_начала'] = [excel_time_to_time_str(row.get(time_from_col)) for _, row in data.iterrows()]
        data['Время_окончания'] = [excel_time_to_time_str(row.get(time_to_col)) for _, row in data.iterrows()]
        data['Длительность_минуты'] = [calculate_duration_minutes(row.get(time_from_col), row.get(time_to_col)) for _, row in data.iterrows()]
    
    # Фенотип
    print("🧬 Нормализация фенотипа...")
    if phenotype_col:
        data['Краткий_фенотип'] = [extract_short_phenotype(row.get(phenotype_col)) for _, row in data.iterrows()]
    
    # Добавление столбцов
    print("\n📝 Добавление новых столбцов...")
    data['Blood_Group_Patient_Norm'] = patient_blood_norm
    data['Rh_Patient_Norm'] = patient_rh_norm
    data['Blood_Group_Patient_Full'] = patient_final_full
    data['Blood_Group_Source'] = patient_final_source
    data['Blood_Group_Env_Norm'] = env_blood_norm
    data['Rh_Env_Norm'] = env_rh_norm
    data['Blood_Group_Env_Full'] = env_full
    data['Component_Type'] = component_types
    data['Проблема_нормализации'] = final_problems
    
    if patient_name_col:
        data['ФИО_пациента'] = data[patient_name_col]
    if patient_id_col:
        data['Номер_пациента'] = data[patient_id_col]
    
    # Сохранение
    output_path = input_path.parent / f"{input_path.stem}{CONFIG['output_suffix']}.xlsx"
    print(f"\n💾 Сохраняем: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='Все_трансфузии', index=False)
        
        problems_data = data[data['Проблема_нормализации'] != '']
        if len(problems_data) > 0:
            problems_data.to_excel(writer, sheet_name='Проблемы_нормализации', index=False)
        
        erythro_mask = data['Component_Type'] == 'Эритроциты'
        if erythro_mask.any():
            data[erythro_mask].to_excel(writer, sheet_name='Эритроциты', index=False)
    
    # Статистика
    print("\n" + "="*60)
    print("📊 СТАТИСТИКА:")
    print(f"   Всего записей: {len(data)}")
    
    problems_count = data[data['Проблема_нормализации'] != '']
    if len(problems_count) > 0:
        print(f"\n⚠️ Проблемы нормализации: {len(problems_count)} записей")
        for problem, count in problems_count['Проблема_нормализации'].value_counts().head(5).items():
            print(f"      {problem[:60]}: {count}")
    else:
        print("\n✅ Проблем нормализации не выявлено")
    
    print(f"\n   Типы компонентов:")
    for comp, count in data['Component_Type'].value_counts().items():
        print(f"      {comp}: {count} ({count/len(data)*100:.1f}%)")
    
    print("\n✅ ГОТОВО")
    print(f"📁 Результат: {output_path}")


if __name__ == "__main__":
    main()