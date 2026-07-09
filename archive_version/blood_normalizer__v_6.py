#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль нормализации данных трансфузий крови v5.1
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import Optional, Tuple, List, Dict
import tkinter as tk
from tkinter import filedialog

CONFIG = {
    'input_folder': 'Enter_data',
    'output_suffix': '_preprocessed',
    'reports_folder': 'reports',
    'search_rows': range(0, 20),
    'required_columns': ['Дата', 'Объем', 'Трансфузионная среда'],
    'debug': True,
    'create_reports_folder': True,
    
    'rh_positive_keywords': ['+', 'ПОЛОЖ', 'POS', 'ПОЗ', '1', 'ДА', 'RH+', 'РЕЗУС+'],
    'rh_negative_keywords': ['-', 'ОТРИЦ', 'NEG', 'НЕГ', '0', 'НЕТ', 'RH-', 'РЕЗУС-'],
    
    'component_keywords': {
        'Эритроциты': ['эритроцит', 'эм', 'erythrocyte', 'эритроцитарная взвесь'],
        'Плазма': ['плазм', 'plasma', 'свежезамороженная плазма', 'сзп'],
        'Тромбоциты': ['тромбоцит', 'platelet', 'тромбоконцентрат'],
        'Криопреципитат': ['криопреципитат', 'cryoprecipitate', 'крио'],
    }
}


def select_file_from_data_folder() -> Optional[Path]:
    root = tk.Tk()
    root.withdraw()
    data_folder = Path(CONFIG['input_folder'])
    if not data_folder.exists():
        data_folder.mkdir(exist_ok=True)
        print(f"📁 Создана папка '{CONFIG['input_folder']}'")
        input("Нажмите Enter после добавления файлов...")
    file_path = filedialog.askopenfilename(
        title="Выберите файл с трансфузиями",
        initialdir=data_folder.absolute(),
        filetypes=[("Excel файлы", "*.xlsx *.xls")]
    )
    root.destroy()
    return Path(file_path) if file_path else None


def find_header_row(df: pd.DataFrame, required_cols: List[str]):
    for idx in CONFIG['search_rows']:
        if idx >= len(df):
            break
        row_values = [str(val).strip() if not pd.isna(val) else '' for val in df.iloc[idx]]
        found_cols = {}
        for req in required_cols:
            matches = [col for col in row_values if req.lower() in col.lower()]
            if matches:
                found_cols[req] = matches[0]
            else:
                break
        else:
            col_mapping = {req: list(row_values).index(col_name) for req, col_name in found_cols.items()}
            return idx, col_mapping
    return None, None


def clean_blood_group_string(abo_str: str) -> str:
    """Очищает строку группы крови от шума"""
    abo = str(abo_str).upper()
    # Убираем скобки, пробелы
    abo = re.sub(r'[\(\)\[\]\{\}\s]', '', abo)
    # Убираем римские цифры
    abo = re.sub(r'\b[IVX]+\b', '', abo)
    # Убираем "Rh" и "резус"
    abo = re.sub(r'RH[+\-]?', '', abo)
    abo = re.sub(r'РЕЗУС[+\-]?', '', abo)
    # Оставляем только нужные символы
    abo = re.sub(r'[^ABO02АВО2\+\-]', '', abo)
    return abo


def is_valid_blood_group(abo_str: any) -> bool:
    if pd.isna(abo_str):
        return False
    abo = str(abo_str).upper().strip()
    if len(abo) > 15:
        return False
    if abo.isdigit() and len(abo) > 2:
        return False
    abo_clean = clean_blood_group_string(abo)
    if not abo_clean:
        return False
    return bool(re.match(r'^[ABO02АВО2\+\-]+$', abo_clean))


def normalize_blood_group_only(abo_str: any) -> Optional[str]:
    if pd.isna(abo_str):
        return None
    abo = clean_blood_group_string(str(abo_str))
    if not abo:
        return None
    has_A = 'А' in abo or 'A' in abo
    has_B = 'В' in abo or 'B' in abo
    has_2 = '2' in abo
    if not has_2:
        has_2 = bool(re.search(r'[АA]2', abo))
    if has_2 and has_A and has_B:
        return 'A2B'
    if has_2 and has_A and not has_B:
        return 'A2'
    if has_A and has_B and not has_2:
        return 'AB'
    if has_A and not has_B and not has_2:
        return 'A'
    if has_B and not has_A and not has_2:
        return 'B'
    if (not has_A and not has_B) or ('О' in abo or 'O' in abo or '0' in abo):
        return 'O'
    return None


def normalize_blood_group_with_fallback(abo_str: any, fallback_str: any = None) -> Optional[str]:
    if is_valid_blood_group(abo_str):
        result = normalize_blood_group_only(abo_str)
        if result:
            return result
    if fallback_str and is_valid_blood_group(fallback_str):
        return normalize_blood_group_only(fallback_str)
    return None


def normalize_rh_only(rh_str: any) -> Optional[str]:
    if pd.isna(rh_str):
        return None
    rh = str(rh_str).upper().strip()
    if '+' in rh:
        return '+'
    if '-' in rh:
        return '-'
    for kw in CONFIG['rh_positive_keywords']:
        if kw in rh:
            return '+'
    for kw in CONFIG['rh_negative_keywords']:
        if kw in rh:
            return '-'
    return None


def get_component_type(comp_str: any) -> str:
    if pd.isna(comp_str):
        return 'Не определено'
    comp = str(comp_str).lower()
    if comp == 'сзп':
        return 'Плазма'
    for comp_type, keywords in CONFIG['component_keywords'].items():
        for kw in keywords:
            if kw in comp:
                return comp_type
    return 'Не определено'


def excel_time_to_str(t: any) -> str:
    if pd.isna(t):
        return ''
    try:
        sec = int(float(t) * 24 * 3600)
        return f"{sec//3600:02d}:{(sec%3600)//60:02d}:{sec%60:02d}"
    except:
        return str(t)


def duration_minutes(start: any, end: any) -> Optional[float]:
    if pd.isna(start) or pd.isna(end):
        return None
    try:
        s, e = float(start), float(end)
        if e < s:
            e += 1
        return round((e - s) * 24 * 60, 1)
    except:
        return None


def extract_phenotype(pheno: any) -> str:
    if pd.isna(pheno):
        return ''
    c = re.search(r'[Cc]{2}', str(pheno))
    e = re.search(r'[Ee]{2}', str(pheno))
    return f"{c.group(0) if c else ''}{e.group(0) if e else ''}"


def main():
    print("🩸 ПРЕПРОЦЕССОР ТРАНСФУЗИЙ v5.1")
    print("="*60)
    
    Path(CONFIG['reports_folder']).mkdir(exist_ok=True)
    
    input_path = select_file_from_data_folder()
    if not input_path:
        return
    
    print(f"\n📂 Загрузка: {input_path.name}")
    df_raw = pd.read_excel(input_path, header=None, dtype=str)
    print(f"   Размер: {df_raw.shape}")
    
    header_row, _ = find_header_row(df_raw, CONFIG['required_columns'])
    if header_row is None:
        print("❌ Заголовки не найдены")
        return
    
    headers = [str(v).strip() if not pd.isna(v) else '' for v in df_raw.iloc[header_row]]
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = headers
    print(f"   Записей: {len(data)}")
    
    # Поиск колонок
    abo_p = rh_p = abo_e = rh_e = comp_col = None
    for col in data.columns:
        cl = col.lower()
        if 'ab0' in cl or 'abo' in cl:
            if 'пац' in cl:
                abo_p = col
            elif 'сред' in cl:
                abo_e = col
        elif 'rh' in cl:
            if 'пац' in cl:
                rh_p = col
            elif 'сред' in cl:
                rh_e = col
        elif 'трансфузионная среда' in cl:
            comp_col = col
    
    print(f"   ABO пациент: {abo_p}, RH пациент: {rh_p}")
    print(f"   ABO среда: {abo_e}, RH среда: {rh_e}")
    
    # Нормализация
    print("\n🩸 Нормализация...")
    
    patient_full = []
    patient_source = []
    problems = []
    
    for _, row in data.iterrows():
        abo = row.get(abo_p)
        rh = row.get(rh_p)
        donor_abo = row.get(abo_e)
        
        blood = normalize_blood_group_with_fallback(abo, donor_abo)
        r = normalize_rh_only(rh)
        
        problem = ''
        if blood and r:
            patient_full.append(f"{blood}{r}")
        elif blood:
            patient_full.append(f"{blood}?")
            problem = 'Не определен резус'
        else:
            # Пробуем взять из донора
            donor_blood = normalize_blood_group_only(donor_abo) if donor_abo else None
            if donor_blood:
                patient_full.append(f"{donor_blood}?")
                problem = f'Группа взята из донора (исходная: {abo})'
            else:
                patient_full.append(None)
                problem = f'Не определена группа (исходная: {abo})'
        
        patient_source.append('patient' if 'ЗАМЕНА' not in problem else 'environment')
        problems.append(problem)
    
    # Типы компонентов
    component_types = [get_component_type(row.get(comp_col)) for _, row in data.iterrows()]
    
    # Сохранение
    data['Blood_Group_Patient_Full'] = patient_full
    data['Blood_Group_Source'] = patient_source
    data['Проблема_нормализации'] = problems
    data['Component_Type'] = component_types
    
    output_path = input_path.parent / f"{input_path.stem}{CONFIG['output_suffix']}.xlsx"
    data.to_excel(output_path, index=False)
    
    # Статистика
    print("\n" + "="*60)
    print("📊 РЕЗУЛЬТАТ:")
    print(f"   Всего: {len(data)}")
    
    problem_count = len([p for p in problems if p])
    print(f"   Проблем: {problem_count}")
    
    if problem_count > 0:
        print("\n   Примеры проблем:")
        for p in problems[:5]:
            if p:
                print(f"      {p[:80]}")
    
    print(f"\n   Типы компонентов:")
    for comp, cnt in data['Component_Type'].value_counts().items():
        print(f"      {comp}: {cnt}")
    
    print(f"\n✅ Сохранено: {output_path}")


if __name__ == "__main__":
    main()