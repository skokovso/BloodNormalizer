#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Расширенный анализ иногруппных трансфузий v3.0
===============================================

Анализирует трансфузии с учетом правил совместимости по приказу №1134н:
- Эритроциты: ABO + Резус (с учетом A2, A2B)
- Плазма: только ABO (донор→реципиент)
- Криопреципитат: любые группы (все совместимы)
- Тромбоциты: любые группы (все совместимы)

Автор: Скоков С.О.
Версия: 3.0
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

DOCTOR_COLUMN = 'ФИО врача'

# =============================================================
# ПРАВИЛА СОВМЕСТИМОСТИ ПО ПРИКАЗУ №1134н
# =============================================================

# 1. Эритроциты (ЭСК) - ABO + Резус (с учетом A2, A2B)
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

# 2. Плазма (только ABO, резус не важен)
# Донор плазмы → Реципиент
PLASMA_COMPATIBILITY = {
    '0': ['0', 'A', 'B', 'AB'],   # плазма 0 подходит всем
    'A': ['A', 'AB'],              # плазма A подходит A и AB
    'B': ['B', 'AB'],              # плазма B подходит B и AB
    'AB': ['AB'],                  # плазма AB только AB
}


# =============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================

def get_abo_only(group: str) -> str:
    """Извлекает только ABO часть (без резуса)"""
    if pd.isna(group):
        return None
    group_str = str(group)
    if group_str.endswith('+') or group_str.endswith('-') or group_str.endswith('?'):
        return group_str[:-1]
    # Для A2, A2B
    if group_str.startswith('A2'):
        return 'A2' if 'B' not in group_str else 'A2B'
    return group_str


def get_compatibility_type(patient_group: str, donor_group: str, component_type: str) -> str:
    """
    Определяет тип совместимости по таблице приказа №1134н
    
    Returns:
        'same' - одногруппная
        'compatible_cross' - совместимая иногруппная (разрешено)
        'incompatible' - несовместимая (нарушение)
    """
    if pd.isna(patient_group) or pd.isna(donor_group):
        return 'unknown'
    
    if patient_group == donor_group:
        return 'same'
    
    # Извлекаем ABO
    patient_abo = get_abo_only(patient_group)
    donor_abo = get_abo_only(donor_group)
    
    # ========== ЭРИТРОЦИТЫ (ЭСК) ==========
    if component_type == 'Эритроциты':
        allowed = COMPATIBILITY_ERYTHROCYTE.get(patient_group, [])
        return 'compatible_cross' if donor_group in allowed else 'incompatible'
    
    # ========== ПЛАЗМА ==========
    elif component_type == 'Плазма':
        # Донор плазмы → Реципиент
        allowed = PLASMA_COMPATIBILITY.get(donor_abo, [])
        return 'compatible_cross' if patient_abo in allowed else 'incompatible'
    
    # ========== КРИОПРЕЦИПИТАТ ==========
    elif component_type == 'Криопреципитат':
        # Все группы совместимы (кроме одинаковых, уже обработано)
        return 'compatible_cross'
    
    # ========== ТРОМБОЦИТЫ ==========
    elif component_type == 'Тромбоциты':
        # Все группы совместимы (кроме одинаковых, уже обработано)
        return 'compatible_cross'
    
    else:
        return 'unknown'


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
    print("🩸 РАСШИРЕННЫЙ АНАЛИЗ ИНОГРУППНЫХ ТРАНСФУЗИЙ v3.0")
    print("="*60)
    
    # 1. Загружаем данные
    print(f"\n📂 Загружаю: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE)
    print(f"   ✅ Загружено {len(df)} записей")
    
    # 2. Определяем тип совместимости для каждого компонента
    print("\n🔍 Определение типов совместимости (по приказу №1134н)...")
    
    compatibility_types = []
    for idx, row in df.iterrows():
        comp_type = row['Component_Type']
        patient = row['Blood_Group_Patient_Full']
        donor = row['Blood_Group_Env_Full']
        compat = get_compatibility_type(patient, donor, comp_type)
        compatibility_types.append(compat)
    
    df['Тип_совместимости'] = compatibility_types
    
    # 3. Определяем год
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
    
    # 4. Фильтруем годы
    years = [2023, 2024, 2025]
    df = df[df['Год'].isin(years)]
    
    # 5. Анализ по типам компонентов
    print("\n" + "="*60)
    print("📊 АНАЛИЗ ПО ТИПАМ КОМПОНЕНТОВ (по приказу №1134н):")
    
    for comp in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
        comp_df = df[df['Component_Type'] == comp]
        if len(comp_df) == 0:
            continue
        
        print(f"\n   🩸 {comp} (n={len(comp_df)}):")
        for year in years:
            year_df = comp_df[comp_df['Год'] == year]
            total = len(year_df)
            same = len(year_df[year_df['Тип_совместимости'] == 'same'])
            compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
            incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
            print(f"      {year}: Всего {total}, одногруппных: {same}, совместимых иногруппных: {compatible}, несовместимых: {incompatible}")
    
    # 6. Несовместимые
    incompatible_df = df[df['Тип_совместимости'] == 'incompatible']
    
    # 7. Сохраняем отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = REPORTS_FOLDER / f'иногруппные_анализ_v3_{timestamp}.xlsx'
    
    print(f"\n💾 Сохраняю отчет: {report_path}")
    
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        # Сводка
        summary_data = []
        for comp in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
            comp_df = df[df['Component_Type'] == comp]
            for year in years:
                year_df = comp_df[comp_df['Год'] == year]
                total = len(year_df)
                same = len(year_df[year_df['Тип_совместимости'] == 'same'])
                compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
                incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
                summary_data.append({
                    'Компонент': comp,
                    'Год': year,
                    'Всего': total,
                    'Одногруппные': same,
                    'Совместимые_иногруппные': compatible,
                    'Процент_иногруппных': compatible / total * 100 if total > 0 else 0,
                    'Несовместимые': incompatible
                })
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Сводка', index=False)
        
        # Несовместимые
        if len(incompatible_df) > 0:
            cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура', 'Component_Type',
                   'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Объем']
            cols = [c for c in cols if c in incompatible_df.columns]
            incompatible_df[cols].to_excel(writer, sheet_name='Несовместимые', index=False)
        
        # Совместимые иногруппные
        compatible_df = df[df['Тип_совместимости'] == 'compatible_cross']
        if len(compatible_df) > 0:
            cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура', 'Component_Type',
                   'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Объем']
            cols = [c for c in cols if c in compatible_df.columns]
            compatible_df[cols].to_excel(writer, sheet_name='Совместимые_иногруппные', index=False)
    
    # 8. Итоговый вывод
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ ДИНАМИКА (Эритроциты):")
    erythro = df[df['Component_Type'] == 'Эритроциты']
    for year in years:
        year_df = erythro[erythro['Год'] == year]
        total = len(year_df)
        compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
        print(f"   {year}: {total} трансфузий, совместимых иногруппных: {compatible} ({compatible/total*100:.1f}%)")
    
    if len(incompatible_df) > 0:
        print(f"\n⚠️ Выявлено {len(incompatible_df)} несовместимых трансфузий (только эритроциты)")
    
    print(f"\n📁 Отчет сохранен: {report_path}")
    print("\n✅ АНАЛИЗ ЗАВЕРШЕН")


if __name__ == "__main__":
    main()