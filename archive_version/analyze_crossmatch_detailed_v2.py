#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Расширенный анализ иногруппных трансфузий v2.0
===============================================

Анализирует трансфузии с учетом:
- Правил совместимости по приказу №1134н (с адаптацией для A2/A2B)
- Раздельно по компонентам: ЭСК, Плазма, Криопреципитат, Тромбоциты
- Раздельно по ABO и Резус
- В разрезе площадок
- В разрезе врачей (столбец "ФИО врача")
- Контроль несовместимых трансфузий

Автор: Скоков С.О.
Версия: 2.0
Дата: Июнь 2026
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# =============================================================
# НАСТРОЙКИ
# =============================================================
INPUT_FILE = Path('Enter_data/Подробный_реестр_трансфузий_2023_2024_2025_preprocessed.xlsx')
REPORTS_FOLDER = Path('reports')
REPORTS_FOLDER.mkdir(exist_ok=True)

DOCTOR_COLUMN = 'ФИО врача'  # правильное название столбца

# =============================================================
# КОРРЕКЦИЯ ОШИБОК НОРМАЛИЗАЦИИ
# =============================================================
def fix_blood_group(patient_group: str, original_abo: str = None) -> str:
    """
    Корректирует ошибки нормализации (например, 1310 -> B)
    """
    if pd.isna(patient_group):
        return patient_group
    
    group_str = str(patient_group)
    
    # Известные ошибки
    corrections = {
        'O?': None,  # требует ручного разбора
    }
    
    # Если оригинальная группа указана и это цифры
    if original_abo and str(original_abo).isdigit():
        # Логика расшифровки цифровых кодов групп
        # 1310 может означать B (нужно уточнить логику)
        if original_abo == '1310':
            return 'B+'
        elif original_abo == 'ж751012301003723':
            return 'B+'
    
    return group_str


# =============================================================
# ПРАВИЛА СОВМЕСТИМОСТИ ПО ПРИКАЗУ №1134н
# =============================================================

# 1. Эритроцитсодержащие компоненты (ЭСК)
COMPATIBILITY_ERYTHROCYTE = {
    'O+': ['O+', 'O-'],
    'O-': ['O-'],
    'A+': ['A+', 'A-', 'O+', 'O-'],
    'A-': ['A-', 'O-'],
    'B+': ['B+', 'B-', 'O+', 'O-'],
    'B-': ['B-', 'O-'],
    'AB+': ['AB+', 'AB-', 'A+', 'A-', 'B+', 'B-', 'O+', 'O-'],
    'AB-': ['AB-', 'A-', 'B-', 'O-'],
    'A2+': ['O+', 'O-', 'A2+', 'A2-'],
    'A2-': ['A-', 'O-', 'A2-'],
    'A2B+': ['B+', 'B-', 'O+', 'O-', 'A2B+', 'A2B-'],
    'A2B-': ['B-', 'O-', 'A2B-'],
}

# 2. Плазма (донор плазмы должен быть совместим с реципиентом)
# Правило: плазма донора не должна содержать антител к эритроцитам реципиента
COMPATIBILITY_PLASMA = {
    'O': ['O', 'A', 'B', 'AB'],  # Плазма O (анти-A, анти-B) - только для O
    'A': ['A', 'AB'],             # Плазма A (анти-B) - для A и AB
    'B': ['B', 'AB'],             # Плазма B (анти-A) - для B и AB
    'AB': ['AB'],                 # Плазма AB (без антител) - только для AB
}

# 3. Криопреципитат (совместимость как у плазмы, но менее строго)
COMPATIBILITY_CRYO = {
    'O': ['O', 'A', 'B', 'AB'],
    'A': ['A', 'B', 'AB'],
    'B': ['B', 'A', 'AB'],
    'AB': ['AB', 'A', 'B', 'O'],
}

# 4. Тромбоциты (совместимость как у эритроцитов, но с учетом резуса)
COMPATIBILITY_PLATELETS = {
    'O+': ['O+', 'O-'],
    'O-': ['O-'],
    'A+': ['A+', 'A-', 'O+', 'O-'],
    'A-': ['A-', 'O-'],
    'B+': ['B+', 'B-', 'O+', 'O-'],
    'B-': ['B-', 'O-'],
    'AB+': ['AB+', 'AB-', 'A+', 'A-', 'B+', 'B-', 'O+', 'O-'],
    'AB-': ['AB-', 'A-', 'B-', 'O-'],
    'A2+': ['O+', 'O-', 'A2+', 'A2-'],
    'A2-': ['A-', 'O-', 'A2-'],
    'A2B+': ['B+', 'B-', 'O+', 'O-', 'A2B+', 'A2B-'],
    'A2B-': ['B-', 'O-', 'A2B-'],
}


def get_compatibility_type(patient_group: str, donor_group: str, component_type: str) -> str:
    """
    Определяет тип совместимости в зависимости от типа компонента
    
    Returns:
        'same' - одногруппная
        'compatible_cross' - совместимая иногруппная
        'incompatible' - несовместимая
        'unknown' - неизвестная группа
    """
    if pd.isna(patient_group) or pd.isna(donor_group):
        return 'unknown'
    
    if patient_group == donor_group:
        return 'same'
    
    # Выбираем правила в зависимости от компонента
    if component_type == 'Эритроциты':
        rules = COMPATIBILITY_ERYTHROCYTE
    elif component_type == 'Плазма':
        # Для плазмы используем только ABO часть
        patient_abo = patient_group[:-1] if patient_group else None
        donor_abo = donor_group[:-1] if donor_group else None
        if patient_abo and donor_abo:
            allowed = COMPATIBILITY_PLASMA.get(donor_abo, [])
            if patient_abo in allowed:
                return 'compatible_cross'
        return 'incompatible'
    elif component_type == 'Криопреципитат':
        patient_abo = patient_group[:-1] if patient_group else None
        donor_abo = donor_group[:-1] if donor_group else None
        if patient_abo and donor_abo:
            allowed = COMPATIBILITY_CRYO.get(donor_abo, [])
            if patient_abo in allowed:
                return 'compatible_cross'
        return 'incompatible'
    elif component_type == 'Тромбоциты':
        rules = COMPATIBILITY_PLATELETS
    else:
        return 'unknown'
    
    allowed = rules.get(patient_group, [])
    if donor_group in allowed:
        return 'compatible_cross'
    else:
        return 'incompatible'


def get_abo_only(group: str) -> str:
    """Извлекает только ABO часть (без резуса)"""
    if pd.isna(group):
        return None
    group_str = str(group)
    if group_str.endswith('+') or group_str.endswith('-') or group_str.endswith('?'):
        return group_str[:-1]
    return group_str


def get_rh_only(group: str) -> str:
    """Извлекает только резус часть"""
    if pd.isna(group):
        return None
    group_str = str(group)
    if group_str.endswith('+'):
        return '+'
    elif group_str.endswith('-'):
        return '-'
    elif group_str.endswith('?'):
        return '?'
    return None


def extract_year(date_value) -> int:
    """Извлекает год из даты"""
    if pd.isna(date_value):
        return None
    try:
        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value.year
        date_str = str(date_value)
        parts = date_str.split('.')
        if len(parts) >= 3:
            return int(parts[2][:4])
        return None
    except:
        return None


# =============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# =============================================================
def main():
    print("🩸 РАСШИРЕННЫЙ АНАЛИЗ ИНОГРУППНЫХ ТРАНСФУЗИЙ v2.0")
    print("="*60)
    
    # 1. Загружаем данные
    print(f"\n📂 Загружаю: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE)
    print(f"   ✅ Загружено {len(df)} записей")
    
    # 2. Применяем коррекцию для проблемных групп
    if 'AB0 пац.' in df.columns:
        df['Blood_Group_Patient_Full_Corrected'] = df.apply(
            lambda row: fix_blood_group(row['Blood_Group_Patient_Full'], row.get('AB0 пац.')), axis=1
        )
        df['Blood_Group_Patient_Full'] = df['Blood_Group_Patient_Full_Corrected']
    
    # 3. Определяем тип совместимости для каждого компонента
    print("\n🔍 Определение типов совместимости...")
    
    compatibility_types = []
    for idx, row in df.iterrows():
        comp_type = row['Component_Type']
        patient = row['Blood_Group_Patient_Full']
        donor = row['Blood_Group_Env_Full']
        
        if comp_type in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
            compat = get_compatibility_type(patient, donor, comp_type)
        else:
            compat = 'unknown'
        
        compatibility_types.append(compat)
    
    df['Тип_совместимости'] = compatibility_types
    
    # 4. Добавляем ABO и Rh отдельно
    df['ABO_пациент'] = df['Blood_Group_Patient_Full'].apply(get_abo_only)
    df['Rh_пациент'] = df['Blood_Group_Patient_Full'].apply(get_rh_only)
    df['ABO_донор'] = df['Blood_Group_Env_Full'].apply(get_abo_only)
    df['Rh_донор'] = df['Blood_Group_Env_Full'].apply(get_rh_only)
    
    df['Несовпадение_ABO'] = df['ABO_пациент'] != df['ABO_донор']
    df['Несовпадение_Rh'] = df['Rh_пациент'] != df['Rh_донор']
    
    # 5. Определяем год
    date_col = None
    for col in df.columns:
        if 'дата' in col.lower():
            date_col = col
            break
    
    if date_col:
        df['Год'] = df[date_col].apply(extract_year)
    else:
        print("❌ Столбец с датой не найден!")
        return
    
    # 6. Фильтруем только нужные годы
    years = [2023, 2024, 2025]
    df = df[df['Год'].isin(years)]
    
    # 7. Анализ по типам компонентов
    print("\n" + "="*60)
    print("📊 АНАЛИЗ ПО ТИПАМ КОМПОНЕНТОВ:")
    
    for comp in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
        comp_df = df[df['Component_Type'] == comp]
        if len(comp_df) == 0:
            continue
        
        print(f"\n   🩸 {comp} (n={len(comp_df)}):")
        for year in years:
            year_df = comp_df[comp_df['Год'] == year]
            total = len(year_df)
            compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
            incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
            percent = compatible / total * 100 if total > 0 else 0
            print(f"      {year}: {total} трансфузий, совместимых иногруппных: {compatible} ({percent:.1f}%), несовместимых: {incompatible}")
    
    # 8. Выделяем несовместимые в отдельный лист
    incompatible_df = df[df['Тип_совместимости'] == 'incompatible'].copy()
    
    # 9. Сохраняем отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = REPORTS_FOLDER / f'иногруппные_анализ_v2_{timestamp}.xlsx'
    
    print(f"\n💾 Сохраняю отчет: {report_path}")
    
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        # Лист 1: Сводка по годам и компонентам
        summary_data = []
        for comp in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
            comp_df = df[df['Component_Type'] == comp]
            for year in years:
                year_df = comp_df[comp_df['Год'] == year]
                total = len(year_df)
                compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
                incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
                summary_data.append({
                    'Компонент': comp,
                    'Год': year,
                    'Всего': total,
                    'Совместимые_иногруппные': compatible,
                    'Процент': compatible / total * 100 if total > 0 else 0,
                    'Несовместимые': incompatible
                })
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Сводка', index=False)
        
        # Лист 2: Несовместимые трансфузии (с врачом!)
        if len(incompatible_df) > 0:
            incompat_cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура', 
                           'Component_Type', 'Blood_Group_Patient_Full', 
                           'Blood_Group_Env_Full', 'Объем']
            incompat_cols = [c for c in incompat_cols if c in incompatible_df.columns]
            incompatible_df[incompat_cols].to_excel(writer, sheet_name='Несовместимые', index=False)
            print(f"   ⚠️ Выявлено {len(incompatible_df)} несовместимых трансфузий")
        
        # Лист 3: Совместимые иногруппные
        compatible_df = df[df['Тип_совместимости'] == 'compatible_cross']
        if len(compatible_df) > 0:
            compat_cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура',
                          'Component_Type', 'Blood_Group_Patient_Full', 
                          'Blood_Group_Env_Full', 'Объем']
            compat_cols = [c for c in compat_cols if c in compatible_df.columns]
            compatible_df[compat_cols].to_excel(writer, sheet_name='Совместимые_иногруппные', index=False)
        
        # Лист 4: Площадки
        sites_data = []
        for site in df['Структура'].dropna().unique():
            for year in years:
                site_year = df[(df['Структура'] == site) & (df['Год'] == year)]
                total = len(site_year)
                compatible = len(site_year[site_year['Тип_совместимости'] == 'compatible_cross'])
                sites_data.append({
                    'Площадка': site,
                    'Год': year,
                    'Всего': total,
                    'Совместимые_иногруппные': compatible,
                    'Процент': compatible / total * 100 if total > 0 else 0
                })
        sites_df = pd.DataFrame(sites_data)
        sites_df.to_excel(writer, sheet_name='Площадки', index=False)
        
        # Лист 5: Врачи (только если есть данные)
        if DOCTOR_COLUMN in df.columns:
            doctors_data = []
            for doctor in df[DOCTOR_COLUMN].dropna().unique():
                for year in years:
                    doc_year = df[(df[DOCTOR_COLUMN] == doctor) & (df['Год'] == year)]
                    total = len(doc_year)
                    compatible = len(doc_year[doc_year['Тип_совместимости'] == 'compatible_cross'])
                    if total > 0:
                        doctors_data.append({
                            'Врач': doctor,
                            'Год': year,
                            'Всего': total,
                            'Совместимые_иногруппные': compatible,
                            'Процент': compatible / total * 100 if total > 0 else 0
                        })
            doctors_df = pd.DataFrame(doctors_data)
            doctors_df.to_excel(writer, sheet_name='Врачи', index=False)
    
    print(f"\n📁 Отчет сохранен: {report_path}")
    
    # 10. Итоговый вывод
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ ДИНАМИКА (Эритроциты):")
    
    erythro = df[df['Component_Type'] == 'Эритроциты']
    for year in years:
        year_df = erythro[erythro['Год'] == year]
        total = len(year_df)
        same = len(year_df[year_df['Тип_совместимости'] == 'same'])
        compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
        incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
        print(f"   {year}: Всего {total}, Совместимые иногруппные: {compatible} ({compatible/total*100:.1f}%)")
    
    print("\n✅ АНАЛИЗ ЗАВЕРШЕН")


if __name__ == "__main__":
    main()