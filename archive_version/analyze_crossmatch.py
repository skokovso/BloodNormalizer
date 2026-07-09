#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Анализатор иногруппных трансфузий за 2023-2025 годы
(использует уже нормализованный файл)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

# =============================================================
# НАСТРОЙКИ
# =============================================================
INPUT_FILE = Path('enter_data/Подробный_реестр_трансфузий_2023_2024_2025_preprocessed.xlsx')
REPORTS_FOLDER = Path('reports')
REPORTS_FOLDER.mkdir(exist_ok=True)


# =============================================================
# ФУНКЦИЯ ИЗВЛЕЧЕНИЯ ГОДА ИЗ ДАТЫ
# =============================================================
def extract_year(date_value):
    """Извлекает год из даты"""
    if pd.isna(date_value):
        return None
    try:
        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value.year
        return int(str(date_value).split('.')[-1]) if '.' in str(date_value) else None
    except:
        return None


# =============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# =============================================================
def main():
    print("🩸 АНАЛИЗ ИНОГРУППНЫХ ТРАНСФУЗИЙ ЗА 2023-2025")
    print("="*60)
    
    # 1. Загружаем нормализованный файл
    print(f"\n📂 Загружаю: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE)
    print(f"   ✅ Загружено {len(df)} записей")
    
    # 2. Оставляем только эритроциты
    erythro = df[df['Component_Type'] == 'Эритроциты'].copy()
    print(f"   🩸 Эритроцитарных трансфузий: {len(erythro)}")
    
    # 3. Определяем год (ищем столбец с датой)
    date_col = None
    for col in erythro.columns:
        if 'дата' in col.lower():
            date_col = col
            break
    
    if date_col is None:
        print("   ❌ Столбец с датой не найден!")
        return
    
    print(f"   📅 Столбец с датой: {date_col}")
    
    # Извлекаем год
    erythro['Год'] = erythro[date_col].apply(extract_year)
    
    # 4. Анализ по годам
    years = [2023, 2024, 2025]
    results = []
    
    print("\n📊 Анализ по годам:")
    print("-" * 50)
    
    for year in years:
        year_data = erythro[erythro['Год'] == year]
        total = len(year_data)
        
        if total == 0:
            print(f"   {year}: нет данных")
            results.append({'year': year, 'total': 0, 'cross': 0, 'percent': 0})
            continue
        
        # Ищем иногруппные
        cross = year_data[year_data['Blood_Group_Patient_Full'] != year_data['Blood_Group_Env_Full']]
        cross_count = len(cross)
        percent = (cross_count / total * 100) if total > 0 else 0
        
        print(f"   {year}: {total} трансфузий, {cross_count} иногруппных ({percent:.1f}%)")
        
        results.append({
            'year': year,
            'total': total,
            'cross': cross_count,
            'percent': percent,
            'data': cross
        })
    
    # 5. Создаем отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = REPORTS_FOLDER / f'иногруппные_анализ_{timestamp}.xlsx'
    
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        # Сводка
        summary = pd.DataFrame([{
            'Год': r['year'],
            'Всего_трансфузий': r['total'],
            'Иногруппные': r['cross'],
            'Процент': f"{r['percent']:.1f}%"
        } for r in results if r['total'] > 0])
        summary.to_excel(writer, sheet_name='Сводка_по_годам', index=False)
        
        # Детали иногруппных
        all_cross = pd.concat([r['data'] for r in results if r['total'] > 0], ignore_index=True)
        if len(all_cross) > 0:
            # Выбираем важные столбцы
            detail_cols = ['Год', date_col, 'Blood_Group_Patient_Full', 'Blood_Group_Env_Full', 'Объем', 'Структура']
            detail_cols = [c for c in detail_cols if c in all_cross.columns]
            all_cross[detail_cols].to_excel(writer, sheet_name='Детали_иногруппных', index=False)
    
    # 6. Выводим динамику
    valid = [r for r in results if r['total'] > 0]
    if len(valid) >= 2:
        print("\n" + "="*60)
        print("📈 ДИНАМИКА:")
        first = valid[0]
        last = valid[-1]
        change = last['percent'] - first['percent']
        
        if change > 0:
            print(f"   ⚠️ Частота ВЫРОСЛА на {change:.1f}%")
        elif change < 0:
            print(f"   ✅ Частота СНИЗИЛАСЬ на {abs(change):.1f}%")
        else:
            print(f"   ➖ Частота не изменилась")
        
        print(f"      {first['year']}: {first['percent']:.1f}% → {last['year']}: {last['percent']:.1f}%")
    
    # 7. Топ несовместимых пар
    if len(all_cross) > 0:
        print("\n🩸 ТОП-5 НЕСОВМЕСТИМЫХ ПАР:")
        pairs = all_cross.groupby(['Blood_Group_Patient_Full', 'Blood_Group_Env_Full']).size().sort_values(ascending=False).head(5)
        for (patient, env), count in pairs.items():
            print(f"      {patient} → {env}: {count} раз")
    
    print(f"\n📁 Отчет сохранен: {report_path}")
    print("\n✅ АНАЛИЗ ЗАВЕРШЕН")


if __name__ == "__main__":
    main()