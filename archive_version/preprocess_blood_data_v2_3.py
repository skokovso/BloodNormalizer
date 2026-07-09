"""
Скрипт для предобработки файла трансфузий перед расчетом запасов крови
Версия 2.3 - исправлена ошибка с float в заголовках

Автор: Трансфузиолог
Дата: Июнь 2026
Версия: 2.3
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# ========== НАСТРОЙКИ ==========
CONFIG = {
    'input_folder': 'данные',
    'output_suffix': '_preprocessed_v2',
    'search_rows': range(0, 20),
    'required_columns': ['Дата', 'Объем', 'Трансфузионная среда'],
    'debug': True,
    'create_reports_folder': True,
}
# =================================

def select_file_from_data_folder():
    """Открывает диалоговое окно для выбора файла из папки 'данные'"""
    root = tk.Tk()
    root.withdraw()
    
    data_folder = Path(CONFIG['input_folder'])
    if not data_folder.exists():
        print(f"📁 Создаю папку '{CONFIG['input_folder']}'...")
        data_folder.mkdir(exist_ok=True)
        print(f"✅ Папка создана. Пожалуйста, положите файлы в папку '{CONFIG['input_folder']}'")
        input("Нажмите Enter после добавления файлов...")
    
    print(f"\n📂 Открываю папку: {data_folder.absolute()}")
    file_path = filedialog.askopenfilename(
        title="Выберите файл с трансфузиями",
        initialdir=data_folder.absolute(),
        filetypes=[
            ("Excel файлы", "*.xlsx *.xls"),
            ("Все файлы", "*.*")
        ]
    )
    
    root.destroy()
    
    if not file_path:
        print("❌ Файл не выбран. Выход...")
        return None
    
    return Path(file_path)

def find_header_row(df, required_cols):
    """Ищет строку с заголовками, возвращает индекс и маппинг колонок"""
    for idx in CONFIG['search_rows']:
        if idx >= len(df):
            break
        
        # БЕЗОПАСНО преобразуем ВСЕ значения в строки
        row_values = []
        for val in df.iloc[idx]:
            if pd.isna(val):
                row_values.append('')  # Пустая строка вместо NaN
            else:
                row_values.append(str(val).strip())
        
        found_cols = {}
        for req in required_cols:
            # Ищем совпадение (без учета регистра)
            matches = [col for col in row_values if req.lower() in col.lower()]
            if matches:
                found_cols[req] = matches[0]
            else:
                break
        else:
            # Нашли все обязательные колонки
            col_mapping = {}
            for i, val in enumerate(row_values):
                for req, col_name in found_cols.items():
                    if val == col_name:
                        col_mapping[req] = i
            return idx, col_mapping
    
    return None, None

def normalize_blood_group_only(abo_str):
    """Нормализует только группу крови"""
    if pd.isna(abo_str):
        return None
    
    abo = str(abo_str).upper().strip()
    abo = re.sub(r'[I\(\)\s]', '', abo)
    
    has_A = 'А' in abo or 'A' in abo
    has_B = 'В' in abo or 'B' in abo
    has_2 = '2' in abo
    
    if not has_2:
        a2_pattern = r'[АA]\s*2'
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

def normalize_rh_only(rh_str):
    """Нормализует резус-фактор"""
    if pd.isna(rh_str):
        return None
    
    rh = str(rh_str).upper().strip()
    
    if '+' in rh:
        return '+'
    elif '-' in rh:
        return '-'
    
    if any(word in rh for word in ['ПОЛОЖ', 'POS', 'ПОЗ']):
        return '+'
    elif any(word in rh for word in ['ОТРИЦ', 'NEG', 'НЕГ']):
        return '-'
    
    return None

def get_component_type(component_str):
    """Определяет тип компонента крови"""
    if pd.isna(component_str):
        return 'Не определено'
    
    comp = str(component_str).lower()
    
    if 'эритроцит' in comp or 'эм' in comp:
        return 'Эритроциты'
    elif 'плазм' in comp:
        return 'Плазма'
    elif 'тромбоцит' in comp:
        return 'Тромбоциты'
    elif 'криопреципитат' in comp:
        return 'Криопреципитат'
    else:
        return 'Другое'

def main():
    print("🩸 ЗАПУСК ПРЕПРОЦЕССОРА ТРАНСФУЗИЙ v2.3")
    print("="*60)
    
    # Создаем папку для отчетов
    if CONFIG['create_reports_folder']:
        reports_dir = Path('reports')
        reports_dir.mkdir(exist_ok=True)
        print(f"📁 Папка для отчетов: {reports_dir}")
    
    # Выбираем файл
    input_path = select_file_from_data_folder()
    if not input_path:
        return
    
    print(f"\n📂 Загружаем файл: {input_path}")
    
    # Загружаем файл (header=None - чтобы не искать заголовки автоматически)
    try:
        df_raw = pd.read_excel(input_path, header=None, dtype=str)
    except Exception as e:
        print(f"❌ Ошибка загрузки файла: {e}")
        return
    
    print(f"   Размер: {df_raw.shape[0]} строк × {df_raw.shape[1]} столбцов")
    
    # Ищем строку с заголовками
    print("\n🔍 Ищем строку с заголовками...")
    header_row, col_mapping = find_header_row(df_raw, CONFIG['required_columns'])
    
    if header_row is None:
        print("❌ Не удалось найти строку с заголовками!")
        print(f"   Искали столбцы: {CONFIG['required_columns']}")
        print("\n   Первые 5 строк файла:")
        for i in range(min(5, len(df_raw))):
            print(f"   Строка {i+1}: {df_raw.iloc[i, :5].tolist()}...")
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
    
    # Находим нужные колонки (игнорируем пустые заголовки)
    print("\n🔎 Поиск необходимых колонок...")
    
    abo_patient_col = None
    rh_patient_col = None
    abo_env_col = None
    rh_env_col = None
    component_col = None
    
    for col in data.columns:
        if not col:  # пропускаем пустые заголовки
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
        if 'трансфузионная среда' in col_lower:
            component_col = col
    
    if CONFIG['debug']:
        print(f"   Пациент: ABO={abo_patient_col}, RH={rh_patient_col}")
        print(f"   Среда: ABO={abo_env_col}, RH={rh_env_col}")
        print(f"   Компонент: {component_col}")
    
    # Нормализация пациента
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
    
    # Нормализация среды
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
    
    # Заполняем пропуски
    print("\n🔄 Заполнение пропусков из данных среды...")
    
    patient_final_full = []
    patient_final_source = []
    
    for i in range(len(data)):
        patient_val = patient_full[i]
        source = 'patient'
        
        if patient_val is None or '?' in patient_val:
            env_val = env_full[i]
            if env_val and '?' not in env_val:
                patient_val = env_val
                source = 'environment'
        
        patient_final_full.append(patient_val)
        patient_final_source.append(source)
    
    # Определяем тип компонента
    print("🔬 Определение типов компонентов...")
    
    component_types = []
    for idx, row in data.iterrows():
        comp = row.get(component_col) if component_col else None
        comp_type = get_component_type(comp)
        component_types.append(comp_type)
    
    # Добавляем новые столбцы
    print("\n📝 Добавление новых столбцов...")
    
    data['Blood_Group_Patient_Norm'] = patient_blood_norm
    data['Rh_Patient_Norm'] = patient_rh_norm
    data['Blood_Group_Patient_Full'] = patient_final_full
    data['Blood_Group_Source'] = patient_final_source
    
    data['Blood_Group_Env_Norm'] = env_blood_norm
    data['Rh_Env_Norm'] = env_rh_norm
    data['Blood_Group_Env_Full'] = env_full
    
    data['Component_Type'] = component_types
    
    # Сохраняем результат
    output_path = input_path.parent / f"{input_path.stem}{CONFIG['output_suffix']}.xlsx"
    
    print(f"\n💾 Сохраняем обработанный файл: {output_path}")
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Все_трансфузии', index=False)
            
            erythro_mask = data['Component_Type'] == 'Эритроциты'
            if erythro_mask.any():
                data[erythro_mask].to_excel(writer, sheet_name='Эритроциты', index=False)
    except Exception as e:
        print(f"❌ Ошибка при сохранении: {e}")
        return
    
    # Отчет о проблемах
    problematic = data['Blood_Group_Patient_Full'].isna() | data['Blood_Group_Patient_Full'].astype(str).str.contains(r'\?', na=False, regex=True)
    problems = data[problematic]
    
    if len(problems) > 0:
        issues_path = input_path.parent / f"{input_path.stem}_issues_v2.xlsx"
        problems.to_excel(issues_path, index=False)
        print(f"\n⚠️  Найдено {len(problems)} проблемных записей")
        print(f"   Отчет: {issues_path}")
    else:
        print("\n✅ Проблемных записей не найдено")
    
    # Статистика
    print("\n" + "="*60)
    print("📊 ИТОГОВАЯ СТАТИСТИКА:")
    print(f"   Всего записей: {len(data)}")
    print(f"\n   Типы компонентов:")
    for comp_type, count in data['Component_Type'].value_counts().items():
        print(f"      {comp_type}: {count} ({count/len(data)*100:.1f}%)")
    
    erythro_data = data[data['Component_Type'] == 'Эритроциты']
    if len(erythro_data) > 0:
        print(f"\n   Группы крови (эритроциты):")
        blood_stats = erythro_data['Blood_Group_Patient_Full'].value_counts().sort_index()
        for blood, count in blood_stats.items():
            print(f"      {blood}: {count} ({count/len(erythro_data)*100:.1f}%)")
    
    print(f"\n   Источники групп крови:")
    for source, count in data['Blood_Group_Source'].value_counts().items():
        print(f"      {source}: {count} ({count/len(data)*100:.1f}%)")
    
    print("\n✅ ПРЕПРОЦЕССИНГ ЗАВЕРШЕН")
    print(f"📁 Результат: {output_path}")

if __name__ == "__main__":
    main()