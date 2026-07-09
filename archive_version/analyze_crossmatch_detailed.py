#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Расширенный анализ иногруппных трансфузий
==========================================

Анализирует трансфузии с учетом:
- Правил совместимости по приказу №1134н (с адаптацией для A2/A2B)
- Раздельно по ABO и Резус
- В разрезе площадок
- В разрезе врачей
- Контроль несовместимых трансфузий

Автор: Скоков С.О.
Дата: Июнь 2026
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

# =============================================================
# НАСТРОЙКИ
# =============================================================
INPUT_FILE = Path('Enter_data/Подробный_реестр_трансфузий_2023_2024_2025_preprocessed.xlsx')
REPORTS_FOLDER = Path('reports')
REPORTS_FOLDER.mkdir(exist_ok=True)

# =============================================================
# ПРАВИЛА СОВМЕСТИМОСТИ ЭРИТРОЦИТОВ (по приказу №1134н с учетом A2/A2B)
# =============================================================
# Для эритроцитсодержащих компонентов (ЭСК)
# Ключ: группа реципиента, Значение: разрешенные группы донора

COMPATIBILITY_ERYTHROCYTE = {
    # Стандартные группы (по приказу)
    'O+': ['O+', 'O-'],
    'O-': ['O-'],
    'A+': ['A+', 'A-', 'O+', 'O-'],
    'A-': ['A-', 'O-'],
    'B+': ['B+', 'B-', 'O+', 'O-'],
    'B-': ['B-', 'O-'],
    'AB+': ['AB+', 'AB-', 'A+', 'A-', 'B+', 'B-', 'O+', 'O-'],
    'AB-': ['AB-', 'A-', 'B-', 'O-'],
    
    # Подгруппа A2 (по вашим правилам)
    'A2+': ['O+', 'O-', 'A2+', 'A2-'],
    'A2-': ['A-', 'O-', 'A2-'],
    
    # Подгруппа A2B (по вашим правилам)
    'A2B+': ['B+', 'B-', 'O+', 'O-', 'A2B+', 'A2B-'],
    'A2B-': ['B-', 'O-', 'A2B-'],
    
    # Для страховки: группы с ? (неопределенный резус)
    'O?': ['O+', 'O-', 'O?'],
    'A?': ['A+', 'A-', 'O+', 'O-', 'A?'],
    'B?': ['B+', 'B-', 'O+', 'O-', 'B?'],
    'AB?': ['AB+', 'AB-', 'A+', 'A-', 'B+', 'B-', 'O+', 'O-', 'AB?'],
    'A2?': ['A-', 'O-', 'A2+', 'A2-', 'A2?'],
    'A2B?': ['B-', 'O-', 'A2B+', 'A2B-', 'A2B?'],
}


def get_compatibility_type_erythrocyte(patient_group: str, donor_group: str) -> str:
    """
    Определяет тип совместимости для эритроцитов
    
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
    
    allowed = COMPATIBILITY_ERYTHROCYTE.get(patient_group, [])
    if donor_group in allowed:
        return 'compatible_cross'
    else:
        return 'incompatible'


def get_abo_only(group: str) -> str:
    """Извлекает только ABO часть (без резуса)"""
    if pd.isna(group):
        return None
    group_str = str(group)
    # Убираем резус в конце
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
    print("🩸 РАСШИРЕННЫЙ АНАЛИЗ ИНОГРУППНЫХ ТРАНСФУЗИЙ")
    print("="*60)
    
    # 1. Загружаем данные
    print(f"\n📂 Загружаю: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE)
    print(f"   ✅ Загружено {len(df)} записей")
    
    # 2. Оставляем только эритроциты
    erythro = df[df['Component_Type'] == 'Эритроциты'].copy()
    print(f"   🩸 Эритроцитарных трансфузий: {len(erythro)}")
    
    # 3. Определяем год
    date_col = None
    for col in erythro.columns:
        if 'дата' in col.lower():
            date_col = col
            break
    
    if date_col is None:
        print("❌ Столбец с датой не найден!")
        return
    
    erythro['Год'] = erythro[date_col].apply(extract_year)
    
    # 4. Определяем тип совместимости
    print("\n🔍 Определение типов совместимости...")
    erythro['Тип_совместимости'] = erythro.apply(
        lambda row: get_compatibility_type_erythrocyte(
            row['Blood_Group_Patient_Full'], 
            row['Blood_Group_Env_Full']
        ), axis=1
    )
    
    # 5. Добавляем ABO и Rh отдельно
    erythro['ABO_пациент'] = erythro['Blood_Group_Patient_Full'].apply(get_abo_only)
    erythro['Rh_пациент'] = erythro['Blood_Group_Patient_Full'].apply(get_rh_only)
    erythro['ABO_донор'] = erythro['Blood_Group_Env_Full'].apply(get_abo_only)
    erythro['Rh_донор'] = erythro['Blood_Group_Env_Full'].apply(get_rh_only)
    
    # 6. Определяем тип несовпадения
    erythro['Несовпадение_ABO'] = erythro['ABO_пациент'] != erythro['ABO_донор']
    erythro['Несовпадение_Rh'] = erythro['Rh_пациент'] != erythro['Rh_донор']
    
    # 7. Анализ по годам
    years = sorted(erythro['Год'].dropna().unique())
    print(f"\n📆 Годы в данных: {years}")
    
    results = []
    all_data = []
    
    for year in years:
        year_data = erythro[erythro['Год'] == year]
        total = len(year_data)
        
        same = len(year_data[year_data['Тип_совместимости'] == 'same'])
        compatible = len(year_data[year_data['Тип_совместимости'] == 'compatible_cross'])
        incompatible = len(year_data[year_data['Тип_совместимости'] == 'incompatible'])
        
        # По ABO и Rh отдельно
        abo_diff = len(year_data[year_data['Несовпадение_ABO']])
        rh_diff = len(year_data[year_data['Несовпадение_Rh']])
        
        results.append({
            'year': year,
            'total': total,
            'same': same,
            'compatible_cross': compatible,
            'incompatible': incompatible,
            'percent_compatible': compatible / total * 100 if total > 0 else 0,
            'abo_diff': abo_diff,
            'rh_diff': rh_diff,
        })
        
        # Сохраняем детали для площадок и врачей
        year_data['Год_для_отчета'] = year
        all_data.append(year_data)
    
    all_erythro = pd.concat(all_data, ignore_index=True)
    
    # 8. Вывод результатов
    print("\n" + "="*60)
    print("📊 ОБЩАЯ ДИНАМИКА:")
    print(f"\n   {'Год':<6} {'Всего':<8} {'Одногруппные':<14} {'Совместимые':<12} {'%':<6} {'Несовместимые':<14}")
    print(f"   {'-'*65}")
    
    for r in results:
        print(f"   {r['year']:<6} {r['total']:<8} {r['same']:<14} {r['compatible_cross']:<12} {r['percent_compatible']:.1f}% {r['incompatible']:<14}")
    
    # 9. Динамика
    if len(results) >= 2:
        first = results[0]
        last = results[-1]
        change = last['percent_compatible'] - first['percent_compatible']
        count_change = last['compatible_cross'] - first['compatible_cross']
        
        print(f"\n📈 ДИНАМИКА СОВМЕСТИМЫХ ИНОГРУППНЫХ:")
        print(f"   Рост доли: {first['percent_compatible']:.1f}% → {last['percent_compatible']:.1f}% (+{change:.1f}%)")
        print(f"   Рост количества: {first['compatible_cross']} → {last['compatible_cross']} (+{count_change} трансфузий, +{count_change/first['compatible_cross']*100:.0f}%)")
    
    # 10. Контроль несовместимых
    incompatible_total = sum(r['incompatible'] for r in results)
    if incompatible_total > 0:
        print(f"\n⚠️ ВНИМАНИЕ: Выявлено {incompatible_total} НЕСОВМЕСТИМЫХ трансфузий!")
        incompatible_data = all_erythro[all_erythro['Тип_совместимости'] == 'incompatible']
        print(incompatible_data[['Год', 'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Структура']].head(10))
    else:
        print(f"\n✅ НЕСОВМЕСТИМЫХ ТРАНСФУЗИЙ НЕ ВЫЯВЛЕНО (контроль пройден)")
    
    # 11. Раздельно по ABO и Rh
    print("\n" + "="*60)
    print("📊 АНАЛИЗ ПО КОМПОНЕНТАМ НЕСОВМЕСТИМОСТИ:")
    print(f"\n   {'Год':<6} {'Несовпадение ABO':<18} {'Несовпадение Rh':<18} {'Оба':<10}")
    print(f"   {'-'*55}")
    
    for r in results:
        both = len(all_erythro[(all_erythro['Год'] == r['year']) & 
                                (all_erythro['Несовпадение_ABO']) & 
                                (all_erythro['Несовпадение_Rh'])])
        print(f"   {r['year']:<6} {r['abo_diff']:<18} {r['rh_diff']:<18} {both:<10}")
    
    # 12. Разбивка по площадкам
    print("\n" + "="*60)
    print("📊 ДИНАМИКА ПО ПЛОЩАДКАМ:")
    
    sites = all_erythro['Структура'].dropna().unique()
    for site in sorted(sites):
        print(f"\n   🏥 {site}:")
        print(f"      {'Год':<6} {'Всего':<8} {'Совместимые':<12} {'%':<8}")
        print(f"      {'-'*40}")
        
        for year in years:
            site_year = all_erythro[(all_erythro['Год'] == year) & (all_erythro['Структура'] == site)]
            total = len(site_year)
            compatible = len(site_year[site_year['Тип_совместимости'] == 'compatible_cross'])
            percent = compatible / total * 100 if total > 0 else 0
            if total > 0:
                print(f"      {year:<6} {total:<8} {compatible:<12} {percent:.1f}%")
    
    # 13. Разбивка по врачам (ищем столбец с фамилией)
    print("\n" + "="*60)
    print("📊 ДИНАМИКА ПО ВРАЧАМ:")
    
    # Поиск столбца с фамилией врача
    doctor_col = None
    for col in all_erythro.columns:
        col_lower = col.lower()
        if 'врач' in col_lower or 'доктор' in col_lower or 'фио' in col_lower:
            doctor_col = col
            break
    
    if doctor_col:
        doctors = all_erythro[doctor_col].dropna().unique()
        print(f"\n   Анализ по столбцу: {doctor_col}")
        
        for doctor in sorted(doctors)[:20]:  # Топ-20 врачей
            print(f"\n   👨‍⚕️ {doctor}:")
            print(f"      {'Год':<6} {'Всего':<8} {'Совместимые':<12} {'%':<8}")
            print(f"      {'-'*40}")
            
            for year in years:
                doc_year = all_erythro[(all_erythro['Год'] == year) & (all_erythro[doctor_col] == doctor)]
                total = len(doc_year)
                compatible = len(doc_year[doc_year['Тип_совместимости'] == 'compatible_cross'])
                percent = compatible / total * 100 if total > 0 else 0
                if total > 0:
                    print(f"      {year:<6} {total:<8} {compatible:<12} {percent:.1f}%")
    else:
        print("\n   ⚠️ Столбец с фамилией врача не найден")
        print(f"   Доступные столбцы: {[c for c in all_erythro.columns if any(x in c.lower() for x in ['врач', 'доктор', 'фио'])]}")
    
    # 14. Сохраняем Excel отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = REPORTS_FOLDER / f'иногруппные_расширенный_анализ_{timestamp}.xlsx'
    
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        # Лист 1: Динамика по годам
        summary_df = pd.DataFrame(results)
        summary_df.to_excel(writer, sheet_name='Динамика_по_годам', index=False)
        
        # Лист 2: Все трансфузии с классификацией
        output_cols = ['Год', date_col, 'Структура', 'Blood_Group_Patient_Full', 'Blood_Group_Env_Full',
                       'Тип_совместимости', 'Несовпадение_ABO', 'Несовпадение_Rh', 'Объем']
        if doctor_col:
            output_cols.insert(3, doctor_col)
        output_cols = [c for c in output_cols if c in all_erythro.columns]
        all_erythro[output_cols].to_excel(writer, sheet_name='Все_трансфузии', index=False)
        
        # Лист 3: Только совместимые иногруппные
        compatible_only = all_erythro[all_erythro['Тип_совместимости'] == 'compatible_cross']
        compatible_only[output_cols].to_excel(writer, sheet_name='Совместимые_иногруппные', index=False)
        
        # Лист 4: Площадки
        sites_data = []
        for site in sorted(sites):
            for year in years:
                site_year = all_erythro[(all_erythro['Год'] == year) & (all_erythro['Структура'] == site)]
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
        
        # Лист 5: Врачи (если найдены)
        if doctor_col:
            doctors_data = []
            for doctor in doctors:
                for year in years:
                    doc_year = all_erythro[(all_erythro['Год'] == year) & (all_erythro[doctor_col] == doctor)]
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
    print("\n✅ РАСШИРЕННЫЙ АНАЛИЗ ЗАВЕРШЕН")


if __name__ == "__main__":
    main()