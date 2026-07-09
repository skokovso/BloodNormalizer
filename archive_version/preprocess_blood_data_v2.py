"""
Скрипт для предобработки файла трансфузий перед расчетом запасов крови
Версия 2.0 с полной нормализацией подгрупп A2 и A2B

Что делает:
1. Создает копию исходного Excel-файла с суффиксом '_preprocessed_v2'
2. Автоматически находит строку с заголовками
3. Нормализует группы крови с учетом A2 и A2B по правилам:
   - Группа определяется по наличию символов A, B, 2
   - Игнорируются римские I и скобки с пробелами
4. Нормализует резус-фактор (+ или -)
5. Добавляет столбцы:
   - Пациент: Blood_Group_Norm, Rh_Norm, Blood_Group_Full (A+, A2+, etc.)
   - Среда: Blood_Group_Env_Norm, Rh_Env_Norm, Blood_Group_Env_Full
   - Тип компонента: Component_Type
6. Заполняет пропуски из данных среды
7. Сохраняет отчет о проблемных записях

Автор: Трансфузиолог
Дата: Июнь 2026
Версия: 2.0
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
import os

# ========== НАСТРОЙКИ ==========
CONFIG = {
    'input_file': 'трансфузии_2025.xlsx',
    'output_suffix': '_preprocessed_v2',
    'search_rows': range(0, 20),
    'required_columns': ['Дата', 'Объем', 'Трансфузионная среда'],
    'debug': True,
    'create_reports_folder': True,  # Создавать папку reports/
}
# =================================

def find_header_row(df, required_cols):
    """Ищет строку с заголовками, возвращает индекс и маппинг колонок"""
    for idx in CONFIG['search_rows']:
        if idx >= len(df):
            break
        
        row = df.iloc[idx].astype(str).str.strip()
        
        found_cols = {}
        for req in required_cols:
            matches = [col for col in row if req.lower() in col.lower()]
            if matches:
                found_cols[req] = matches[0]
            else:
                break
        else:
            col_mapping = {}
            for i, val in enumerate(row):
                for req, col_name in found_cols.items():
                    if val == col_name:
                        col_mapping[req] = i
            return idx, col_mapping
    
    return None, None

def normalize_blood_group_only(abo_str):
    """
    Нормализует только группу крови (без резуса)
    Возвращает: O, A, B, AB, A2, A2B или None
    """
    if pd.isna(abo_str):
        return None
    
    # Приводим к строке и убираем лишнее
    abo = str(abo_str).upper().strip()
    
    # Убираем римские I, скобки, пробелы
    abo = re.sub(r'[I\(\)\s]', '', abo)
    
    # Определяем наличие символов
    has_A = 'А' in abo or 'A' in abo  # русская и латинская A
    has_B = 'В' in abo or 'B' in abo  # русская и латинская B
    has_2 = '2' in abo
    
    # Ищем цифру 2 после A (с учетом возможного пробела)
    if not has_2:
        # Проверяем паттерн A2, А2, A 2, А 2
        a2_pattern = r'[АA]\s*2'
        has_2 = bool(re.search(a2_pattern, abo))
    
    # Логика определения группы
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
        # Группа O (ноль)
        return 'O'
    else:
        return None

def normalize_rh_only(rh_str):
    """
    Нормализует только резус-фактор
    Возвращает: '+', '-' или None
    """
    if pd.isna(rh_str):
        return None
    
    rh = str(rh_str).upper().strip()
    
    # Ищем явный + или -
    if '+' in rh:
        return '+'
    elif '-' in rh:
        return '-'
    
    # Варианты написания
    if any(word in rh for word in ['ПОЛОЖ', 'POS', 'ПОЗ']):
        return '+'
    elif any(word in rh for word in ['ОТРИЦ', 'NEG', 'НЕГ']):
        return '-'
    
    return None

def get_component_type(component_str):
    """
    Определяет тип компонента крови
    Возвращает: 'Эритроциты', 'Плазма', 'Тромбоциты', 'Криопреципитат', 'Другое'
    """
    if pd.isna(component_str):
        return 'Не определено'
    
    comp = str(component_str).lower()
    
    # Приоритет: сначала ищем точные совпадения
    if 'эритроцит' in comp or 'эм' in comp or 'erythrocyte' in comp:
        return 'Эритроциты'
    elif 'плазм' in comp or 'plasma' in comp:
        return 'Плазма'
    elif 'тромбоцит' in comp or 'platelet' in comp or 'тромбоконцентрат' in comp:
        return 'Тромбоциты'
    elif 'криопреципитат' in comp or 'cryoprecipitate' in comp:
        return 'Криопреципитат'
    else:
        return 'Другое'

def main():
    print("🩸 ЗАПУСК ПРЕПРОЦЕССОРА ТРАНСФУЗИЙ v2.0")
    print("="*60)
    
    # Создаем папку для отчетов если нужно
    if CONFIG['create_reports_folder']:
        reports_dir = Path('reports')
        reports_dir.mkdir(exist_ok=True)
        print(f"📁 Папка для отчетов: {reports_dir}")
    
    # 1. Загружаем исходный файл
    input_path = Path(CONFIG['input_file'])
    if not input_path.exists():
        print(f"❌ Файл {input_path} не найден!")
        return
    
    print(f"📂 Загружаем файл: {input_path}")
    df_raw = pd.read_excel(input_path, header=None, dtype=str)
    print(f"   Размер: {df_raw.shape[0]} строк × {df_raw.shape[1]} столбцов")
    
    # 2. Находим строку с заголовками
    print("\n🔍 Ищем строку с заголовками...")
    header_row, col_mapping = find_header_row(df_raw, CONFIG['required_columns'])
    
    if header_row is None:
        print("❌ Не удалось найти строку с заголовками!")
        return
    
    print(f"✅ Заголовки найдены в строке {header_row + 1}")
    
    # 3. Создаем DataFrame
    headers = df_raw.iloc[header_row].astype(str).str.strip()
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = headers
    
    print(f"📊 Загружено {len(data)} записей")
    
    # 4. Находим нужные колонки
    print("\n🔎 Поиск необходимых колонок...")
    
    # Ищем колонки для пациента
    abo_patient_col = None
    rh_patient_col = None
    
    # Ищем колонки для среды
    abo_env_col = None
    rh_env_col = None
    
    for col in data.columns:
        col_lower = col.lower()
        if ('ab0' in col_lower or 'abo' in col_lower) and ('пац' in col_lower):
            abo_patient_col = col
        elif ('rh' in col_lower) and ('пац' in col_lower):
            rh_patient_col = col
        elif ('ab0' in col_lower or 'abo' in col_lower) and ('сред' in col_lower):
            abo_env_col = col
        elif ('rh' in col_lower) and ('сред' in col_lower):
            rh_env_col = col
    
    if CONFIG['debug']:
        print(f"   Пациент: ABO={abo_patient_col}, RH={rh_patient_col}")
        print(f"   Среда: ABO={abo_env_col}, RH={rh_env_col}")
    
    # 5. Нормализация данных пациента
    print("\n🩸 Нормализация групп крови пациента...")
    
    patient_blood_norm = []
    patient_rh_norm = []
    patient_full = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_patient_col) if abo_patient_col else None
        rh = row.get(rh_patient_col) if rh_patient_col else None
        
        blood_norm = normalize_blood_group_only(abo)
        rh_norm = normalize_rh_only(rh)
        
        patient_blood_norm.append(blood_norm)
        patient_rh_norm.append(rh_norm)
        
        if blood_norm and rh_norm:
            patient_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm:
            patient_full.append(f"{blood_norm}?")
        else:
            patient_full.append(None)
    
    # 6. Нормализация данных среды
    print("🩸 Нормализация групп крови среды...")
    
    env_blood_norm = []
    env_rh_norm = []
    env_full = []
    
    for idx, row in data.iterrows():
        abo = row.get(abo_env_col) if abo_env_col else None
        rh = row.get(rh_env_col) if rh_env_col else None
        
        blood_norm = normalize_blood_group_only(abo)
        rh_norm = normalize_rh_only(rh)
        
        env_blood_norm.append(blood_norm)
        env_rh_norm.append(rh_norm)
        
        if blood_norm and rh_norm:
            env_full.append(f"{blood_norm}{rh_norm}")
        elif blood_norm:
            env_full.append(f"{blood_norm}?")
        else:
            env_full.append(None)
    
    # 7. Заполняем пропуски пациента из среды
    print("\n🔄 Заполнение пропусков из данных среды...")
    
    patient_final_full = []
    patient_final_source = []
    
    for i in range(len(data)):
        # Сначала берем пациента
        patient_val = patient_full[i]
        source = 'patient'
        
        # Если нет или некорректно - берем из среды
        if patient_val is None or '?' in patient_val:
            env_val = env_full[i]
            if env_val and '?' not in env_val:
                patient_val = env_val
                source = 'environment'
        
        patient_final_full.append(patient_val)
        patient_final_source.append(source)
    
    # 8. Определяем тип компонента
    print("🔬 Определение типов компонентов крови...")
    component_col = None
    for col in data.columns:
        if 'трансфузионная среда' in col.lower():
            component_col = col
            break
    
    component_types = []
    for idx, row in data.iterrows():
        comp = row.get(component_col) if component_col else None
        comp_type = get_component_type(comp)
        component_types.append(comp_type)
    
    # 9. Добавляем новые столбцы
    print("\n📝 Добавление новых столбцов...")
    
    data['Blood_Group_Patient_Norm'] = patient_blood_norm
    data['Rh_Patient_Norm'] = patient_rh_norm
    data['Blood_Group_Patient_Full'] = patient_final_full
    data['Blood_Group_Source'] = patient_final_source
    
    data['Blood_Group_Env_Norm'] = env_blood_norm
    data['Rh_Env_Norm'] = env_rh_norm
    data['Blood_Group_Env_Full'] = env_full
    
    data['Component_Type'] = component_types
    
    # 10. Сохраняем результат
    output_filename = f"{input_path.stem}{CONFIG['output_suffix']}.xlsx"
    output_path = input_path.parent / output_filename
    
    print(f"\n💾 Сохраняем обработанный файл: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Основной лист со всеми данными
        data.to_excel(writer, sheet_name='Все_трансфузии', index=False)
        
        # Лист только с эритроцитами
        erythro_mask = data['Component_Type'] == 'Эритроциты'
        if erythro_mask.any():
            data[erythro_mask].to_excel(writer, sheet_name='Эритроциты', index=False)
        
        # Лист со статистикой
        stats_data = []
        
        # Статистика по типам компонентов
        comp_stats = data['Component_Type'].value_counts()
        for comp_type, count in comp_stats.items():
            stats_data.append(['Тип компонента', comp_type, count, ''])
        
        stats_data.append(['', '', '', ''])
        
        # Статистика по группам крови (эритроциты)
        erythro_data = data[erythro_mask] if erythro_mask.any() else pd.DataFrame()
        if len(erythro_data) > 0:
            blood_stats = erythro_data['Blood_Group_Patient_Full'].value_counts()
            stats_data.append(['Группа крови (эритроциты)', 'Количество', '', ''])
            for blood, count in blood_stats.items():
                stats_data.append([blood, count, '', ''])
        
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='Статистика', index=False, header=False)
    
    # 11. Отчет о проблемных записях
    problems = data[
        (data['Blood_Group_Patient_Full'].isna()) | 
        (data['Blood_Group_Patient_Full'].str.contains('\?', na=False))
    ]
    
    if len(problems) > 0:
        issues_path = input_path.parent / f"{input_path.stem}_issues_v2.xlsx"
        problems.to_excel(issues_path, index=False)
        print(f"\n⚠️  Найдено {len(problems)} проблемных записей")
        print(f"   Отчет сохранен: {issues_path}")
    else:
        print("\n✅ Проблемных записей не найдено")
    
    # 12. Финальная статистика
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ СТАТИСТИКА:")
    print(f"\n   Всего записей: {len(data)}")
    print(f"\n   Типы компонентов:")
    for comp_type, count in data['Component_Type'].value_counts().items():
        print(f"      {comp_type}: {count} ({count/len(data)*100:.1f}%)")
    
    print(f"\n   Источники групп крови пациента:")
    for source, count in data['Blood_Group_Source'].value_counts().items():
        print(f"      {source}: {count} ({count/len(data)*100:.1f}%)")
    
    erythro_count = len(data[data['Component_Type'] == 'Эритроциты'])
    if erythro_count > 0:
        print(f"\n   Группы крови (только эритроциты):")
        erythro_data = data[data['Component_Type'] == 'Эритроциты']
        blood_stats = erythro_data['Blood_Group_Patient_Full'].value_counts().sort_index()
        for blood, count in blood_stats.items():
            print(f"      {blood}: {count} ({count/erythro_count*100:.1f}%)")
    
    print("\n" + "="*60)
    print("✅ ПРЕПРОЦЕССИНГ ЗАВЕРШЕН")
    print(f"📁 Результат: {output_path}")
    
    if CONFIG['create_reports_folder']:
        print(f"📁 Папка с отчетами: {reports_dir}")
    
    print("\n📌 Следующие шаги:")
    print("   1. Проверьте файл *_issues_v2.xlsx (если есть)")
    print("   2. Используйте полученный файл для расчета запасов")

if __name__ == "__main__":
    main()