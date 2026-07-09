#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Расширенный анализ иногруппных трансфузий v4.0
===============================================

Анализирует трансфузии с учетом правил совместимости по приказу №1134н:
- Эритроциты: ABO + Резус (с учетом A2, A2B)
- Плазма: только ABO (по таблице: реципиент 0 → любая плазма, A → A/AB, B → B/AB, AB → AB)
- Криопреципитат: любые группы (все совместимы)
- Тромбоциты: любые группы (все совместимы)

Автор: Скоков С.О.
Версия: 4.0
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
# Реципиент → разрешенная плазма донора (по таблице приказа)
PLASMA_COMPATIBILITY = {
    '0': ['0', 'A', 'B', 'AB'],   # реципиенту 0 → любая плазма
    'A': ['A', 'AB'],              # реципиенту A → плазма A и AB
    'B': ['B', 'AB'],              # реципиенту B → плазма B и AB
    'AB': ['AB'],                  # реципиенту AB → только плазма AB
    'A2': ['A', 'AB'],             # A2 реципиенту → как A
    'A2B': ['AB'],                 # A2B реципиенту → как AB
}


# =============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================

def get_abo_only(group: str) -> str:
    """
    Извлекает только ABO часть (без резуса)
    
    Примеры:
        'A+' -> 'A'
        'A2-' -> 'A2'
        'A2B+' -> 'A2B'
        'AB?' -> 'AB'
    """
    if pd.isna(group):
        return None
    group_str = str(group)
    # Убираем резус в конце (+, -, ?)
    if group_str.endswith('+') or group_str.endswith('-') or group_str.endswith('?'):
        group_str = group_str[:-1]
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


def get_compatibility_type(patient_group: str, donor_group: str, component_type: str) -> str:
    """
    Определяет тип совместимости по таблице приказа №1134н
    
    Returns:
        'same' - одногруппная
        'compatible_cross' - совместимая иногруппная (разрешено)
        'incompatible' - несовместимая (нарушение)
        'unknown' - неизвестная группа
    """
    if pd.isna(patient_group) or pd.isna(donor_group):
        return 'unknown'
    
    # Одногруппная трансфузия
    if patient_group == donor_group:
        return 'same'
    
    # Извлекаем ABO часть
    patient_abo = get_abo_only(patient_group)
    donor_abo = get_abo_only(donor_group)
    
    # ========== 1. ЭРИТРОЦИТЫ (ЭСК) ==========
    if component_type == 'Эритроциты':
        allowed = COMPATIBILITY_ERYTHROCYTE.get(patient_group, [])
        if donor_group in allowed:
            return 'compatible_cross'
        else:
            return 'incompatible'
    
    # ========== 2. ПЛАЗМА ==========
    elif component_type == 'Плазма':
        # Реципиент → разрешенная плазма донора
        allowed_donor_plasmas = PLASMA_COMPATIBILITY.get(patient_abo, [])
        if donor_abo in allowed_donor_plasmas:
            return 'compatible_cross'
        else:
            return 'incompatible'
    
    # ========== 3. КРИОПРЕЦИПИТАТ ==========
    elif component_type == 'Криопреципитат':
        # По приказу: любые группы совместимы
        return 'compatible_cross'
    
    # ========== 4. ТРОМБОЦИТЫ ==========
    elif component_type == 'Тромбоциты':
        # По приказу: любые группы совместимы
        return 'compatible_cross'
    
    else:
        return 'unknown'


def extract_year(date_value) -> int:
    """Извлекает год из даты (поддерживает разные форматы)"""
    if pd.isna(date_value):
        return None
    try:
        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value.year
        date_str = str(date_value)
        # Формат ДД.ММ.ГГГГ
        if '.' in date_str:
            parts = date_str.split('.')
            if len(parts) >= 3:
                return int(parts[2][:4])
        # Формат ГГГГ-ММ-ДД
        elif '-' in date_str:
            return int(date_str[:4])
        # Просто ищем 4 цифры подряд
        import re
        match = re.search(r'\b(20\d{2})\b', date_str)
        if match:
            return int(match.group(1))
        return None
    except:
        return None


def print_summary_table(df, years, component_type):
    """Выводит красиво отформатированную таблицу для одного компонента"""
    print(f"\n   📊 {component_type}:")
    print(f"   {'─'*65}")
    print(f"   {'Год':<6} {'Всего':<8} {'Одногруппные':<14} {'Совместимые':<14} {'Несовместимые':<14}")
    print(f"   {'─'*65}")
    
    comp_df = df[df['Component_Type'] == component_type]
    for year in years:
        year_df = comp_df[comp_df['Год'] == year]
        total = len(year_df)
        same = len(year_df[year_df['Тип_совместимости'] == 'same'])
        compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
        incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
        
        if total > 0:
            print(f"   {year:<6} {total:<8} {same:<14} {compatible:<14} {incompatible:<14}")
        else:
            print(f"   {year:<6} {'нет данных':<8}")
    
    # Итоговая строка
    total_all = len(comp_df[comp_df['Год'].isin(years)])
    if total_all > 0:
        same_all = len(comp_df[comp_df['Тип_совместимости'] == 'same'])
        compat_all = len(comp_df[comp_df['Тип_совместимости'] == 'compatible_cross'])
        incompat_all = len(comp_df[comp_df['Тип_совместимости'] == 'incompatible'])
        print(f"   {'─'*65}")
        print(f"   {'ИТОГО':<6} {total_all:<8} {same_all:<14} {compat_all:<14} {incompat_all:<14}")


def print_site_analysis(df, years, sites):
    """Выводит анализ по площадкам"""
    print("\n" + "="*70)
    print("📊 ДИНАМИКА ПО ПЛОЩАДКАМ (Эритроциты):")
    print("="*70)
    
    erythro = df[df['Component_Type'] == 'Эритроциты']
    
    for site in sorted(sites):
        site_df = erythro[erythro['Структура'] == site]
        if len(site_df) == 0:
            continue
        
        print(f"\n   🏥 {site}:")
        print(f"   {'─'*55}")
        print(f"   {'Год':<6} {'Всего':<8} {'Одногруппные':<14} {'Совместимые':<14} {'%':<8}")
        print(f"   {'─'*55}")
        
        for year in years:
            year_site = site_df[site_df['Год'] == year]
            total = len(year_site)
            if total == 0:
                continue
            same = len(year_site[year_site['Тип_совместимости'] == 'same'])
            compatible = len(year_site[year_site['Тип_совместимости'] == 'compatible_cross'])
            percent = compatible / total * 100 if total > 0 else 0
            print(f"   {year:<6} {total:<8} {same:<14} {compatible:<14} {percent:.1f}%")


def print_doctor_analysis(df, years, doctor_col):
    """Выводит анализ по врачам (только с положительной динамикой)"""
    print("\n" + "="*70)
    print("📊 ДИНАМИКА ПО ВРАЧАМ (Эритроциты, рост совместимых):")
    print("="*70)
    
    erythro = df[df['Component_Type'] == 'Эритроциты']
    
    # Анализируем врачей, у которых есть данные за несколько лет
    doctor_stats = []
    for doctor in erythro[doctor_col].dropna().unique():
        doctor_df = erythro[erythro[doctor_col] == doctor]
        years_data = {}
        for year in years:
            year_df = doctor_df[doctor_df['Год'] == year]
            total = len(year_df)
            compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
            percent = compatible / total * 100 if total > 0 else 0
            if total > 0:
                years_data[year] = {'total': total, 'compatible': compatible, 'percent': percent}
        
        if len(years_data) >= 2:
            first_year = min(years_data.keys())
            last_year = max(years_data.keys())
            if years_data[first_year]['total'] >= 5:  # только врачи с достаточной статистикой
                doctor_stats.append({
                    'doctor': doctor,
                    'first_year': first_year,
                    'first_percent': years_data[first_year]['percent'],
                    'last_year': last_year,
                    'last_percent': years_data[last_year]['percent'],
                    'change': years_data[last_year]['percent'] - years_data[first_year]['percent'],
                    'total_transfusions': years_data[last_year]['total']
                })
    
    # Сортируем по изменению процента (кто больше вырос)
    doctor_stats.sort(key=lambda x: x['change'], reverse=True)
    
    print(f"\n   {'Врач':<35} {'Период':<12} {'% (начало)':<12} {'% (конец)':<12} {'Изменение':<10}")
    print(f"   {'─'*85}")
    
    for stat in doctor_stats[:15]:  # Топ-15
        change_symbol = '▲' if stat['change'] > 0 else '▼' if stat['change'] < 0 else '●'
        print(f"   {stat['doctor']:<35} {stat['first_year']}-{stat['last_year']:<6} "
              f"{stat['first_percent']:>5.1f}%{'':<6} {stat['last_percent']:>5.1f}%{'':<6} "
              f"{change_symbol} {stat['change']:>+5.1f}%")


# =============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# =============================================================

def main():
    print("🩸 РАСШИРЕННЫЙ АНАЛИЗ ИНОГРУППНЫХ ТРАНСФУЗИЙ v4.0")
    print("="*60)
    print("\nПравила совместимости по приказу №1134н от 20.10.2020:")
    print("   • Эритроциты: ABO + Резус (с учетом A2, A2B)")
    print("   • Плазма: только ABO (реципиент 0→любая, A→A/AB, B→B/AB, AB→AB)")
    print("   • Криопреципитат: любые группы")
    print("   • Тромбоциты: любые группы")
    
    # 1. Загружаем данные
    print(f"\n📂 Загружаю: {INPUT_FILE.name}")
    
    if not INPUT_FILE.exists():
        print(f"❌ Файл не найден: {INPUT_FILE}")
        print("   Проверьте путь к файлу")
        return
    
    df = pd.read_excel(INPUT_FILE)
    print(f"   ✅ Загружено {len(df)} записей")
    
    # 2. Определяем тип совместимости для каждого компонента
    print("\n🔍 Определение типов совместимости...")
    
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
    
    if date_col is None:
        print("❌ Столбец с датой не найден!")
        print(f"   Доступные столбцы: {list(df.columns[:10])}...")
        return
    
    print(f"   📅 Столбец с датой: {date_col}")
    df['Год'] = df[date_col].apply(extract_year)
    
    # 4. Фильтруем нужные годы
    years = [2023, 2024, 2025]
    df = df[df['Год'].isin(years)]
    print(f"   📆 Годы в данных: {sorted(df['Год'].unique())}")
    
    # 5. Анализ по типам компонентов
    print("\n" + "="*60)
    print("📊 АНАЛИЗ ПО ТИПАМ КОМПОНЕНТОВ")
    print("="*60)
    
    for comp in ['Эритроциты', 'Плазма', 'Криопреципитат', 'Тромбоциты']:
        if len(df[df['Component_Type'] == comp]) > 0:
            print_summary_table(df, years, comp)
    
    # 6. Несовместимые трансфузии
    incompatible_df = df[df['Тип_совместимости'] == 'incompatible']
    
    if len(incompatible_df) > 0:
        print("\n" + "="*60)
        print(f"⚠️ ВЫЯВЛЕНО {len(incompatible_df)} НЕСОВМЕСТИМЫХ ТРАНСФУЗИЙ")
        print("="*60)
        print("\n   Распределение по компонентам:")
        for comp, count in incompatible_df['Component_Type'].value_counts().items():
            print(f"      {comp}: {count}")
        
        # Показываем примеры несовместимых
        print("\n   📋 Примеры несовместимых трансфузий:")
        sample_cols = ['Год', 'Component_Type', 'Blood_Group_Patient_Full', 
                       'Blood_Group_Env_Full', DOCTOR_COLUMN, 'Структура']
        sample_cols = [c for c in sample_cols if c in incompatible_df.columns]
        print(incompatible_df[sample_cols].head(10).to_string(index=False))
    else:
        print("\n" + "="*60)
        print("✅ НЕСОВМЕСТИМЫХ ТРАНСФУЗИЙ НЕ ВЫЯВЛЕНО")
        print("="*60)
    
    # 7. Анализ по площадкам (только эритроциты)
    sites = df[df['Component_Type'] == 'Эритроциты']['Структура'].dropna().unique()
    if len(sites) > 0:
        print_site_analysis(df, years, sites)
    
    # 8. Анализ по врачам
    if DOCTOR_COLUMN in df.columns:
        print_doctor_analysis(df, years, DOCTOR_COLUMN)
    else:
        print(f"\n⚠️ Столбец '{DOCTOR_COLUMN}' не найден в данных")
    
    # 9. Сохраняем Excel отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = REPORTS_FOLDER / f'иногруппные_анализ_v4_{timestamp}.xlsx'
    
    print(f"\n💾 Сохраняю Excel отчет: {report_path}")
    
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        # Лист 1: Сводка по годам и компонентам
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
                    'Процент_иногруппных': round(compatible / total * 100, 1) if total > 0 else 0,
                    'Несовместимые': incompatible
                })
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Сводка', index=False)
        
        # Лист 2: Несовместимые трансфузии
        if len(incompatible_df) > 0:
            cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура', 'Component_Type',
                   'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Объем']
            cols = [c for c in cols if c in incompatible_df.columns]
            incompatible_df[cols].to_excel(writer, sheet_name='Несовместимые', index=False)
        
        # Лист 3: Совместимые иногруппные
        compatible_df = df[df['Тип_совместимости'] == 'compatible_cross']
        if len(compatible_df) > 0:
            cols = ['Год', date_col, DOCTOR_COLUMN, 'Структура', 'Component_Type',
                   'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Объем']
            cols = [c for c in cols if c in compatible_df.columns]
            compatible_df[cols].to_excel(writer, sheet_name='Совместимые_иногруппные', index=False)
        
        # Лист 4: Площадки
        if len(sites) > 0:
            sites_data = []
            erythro = df[df['Component_Type'] == 'Эритроциты']
            for site in sites:
                for year in years:
                    site_year = erythro[(erythro['Структура'] == site) & (erythro['Год'] == year)]
                    total = len(site_year)
                    compatible = len(site_year[site_year['Тип_совместимости'] == 'compatible_cross'])
                    sites_data.append({
                        'Площадка': site,
                        'Год': year,
                        'Всего': total,
                        'Совместимые_иногруппные': compatible,
                        'Процент': round(compatible / total * 100, 1) if total > 0 else 0
                    })
            sites_df = pd.DataFrame(sites_data)
            sites_df.to_excel(writer, sheet_name='Площадки', index=False)
    
    # 10. Итоговый вывод по эритроцитам
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ ДИНАМИКА (Эритроциты):")
    print("="*60)
    
    erythro = df[df['Component_Type'] == 'Эритроциты']
    for year in years:
        year_df = erythro[erythro['Год'] == year]
        total = len(year_df)
        compatible = len(year_df[year_df['Тип_совместимости'] == 'compatible_cross'])
        incompatible = len(year_df[year_df['Тип_совместимости'] == 'incompatible'])
        print(f"\n   {year}:")
        print(f"      Всего: {total}")
        print(f"      Совместимые иногруппные: {compatible} ({compatible/total*100:.1f}%)")
        print(f"      Несовместимые: {incompatible}")
    
    # Динамика
    if len(erythro[erythro['Год'] == 2023]) > 0 and len(erythro[erythro['Год'] == 2025]) > 0:
        y2023 = len(erythro[(erythro['Год'] == 2023) & (erythro['Тип_совместимости'] == 'compatible_cross')])
        y2025 = len(erythro[(erythro['Год'] == 2025) & (erythro['Тип_совместимости'] == 'compatible_cross')])
        total2023 = len(erythro[erythro['Год'] == 2023])
        total2025 = len(erythro[erythro['Год'] == 2025])
        
        pct2023 = y2023 / total2023 * 100 if total2023 > 0 else 0
        pct2025 = y2025 / total2025 * 100 if total2025 > 0 else 0
        change = pct2025 - pct2023
        count_change = y2025 - y2023
        
        print(f"\n📈 ДИНАМИКА 2023 → 2025:")
        print(f"   Рост доли: {pct2023:.1f}% → {pct2025:.1f}% (+{change:.1f}%)")
        print(f"   Рост количества: {y2023} → {y2025} (+{count_change} трансфузий)")
    
    print(f"\n📁 Excel отчет сохранен: {report_path}")
    print("\n✅ АНАЛИЗ ЗАВЕРШЕН")


if __name__ == "__main__":
    main()