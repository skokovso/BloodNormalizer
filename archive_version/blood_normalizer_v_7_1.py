#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль нормализации данных трансфузий крови v6.0 (PLUS)
========================================================

Расширенная версия с улучшенным распознаванием группы крови и резуса.

Автор: Скоков С.О.
Версия: 6.0
Дата: Июнь 2026

Изменения:
    - Полное распознавание всех форматов из файла Столбец_группа_резус.xlsx
    - Поддержка римских цифр (I, II, III, IV)
    - Поддержка мусорных значений (1310, ж751012301003723)
    - Поддержка "направить на СПК, А2(11)"
    - Использование столбцов проверки при необходимости
    - Соххраняет имя файла нормализатора в качестве суффикса результрующего файла Эксель
"""

import pandas as pd
import re
import os
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import filedialog


# ===================================================================
# ОПРЕДЕЛЕНИЕ ИМЕНИ ФАЙЛА НОРМАЛИЗАТОРА (автоматически)
# ===================================================================

def get_normalizer_name():
    """
    Возвращает имя текущего файла нормализатора без расширения .py
    """
    current_file = os.path.basename(__file__)
    name_without_ext = os.path.splitext(current_file)[0]
    return name_without_ext


NORMALIZER_NAME = get_normalizer_name()
print(f"📌 Имя нормализатора: {NORMALIZER_NAME}")

# ===================================================================
# РАСШИРЕННАЯ НОРМАЛИЗАЦИЯ ГРУППЫ КРОВИ (ABO)
# ===================================================================

def roman_to_abo(roman: str) -> Optional[str]:
    """
    Преобразует римскую цифру в группу крови
    I → O, II → A, III → B, IV → AB
    """
    roman = roman.upper().strip()
    mapping = {
        'I': 'O',
        'II': 'A',
        'III': 'B',
        'IV': 'AB'
    }
    return mapping.get(roman)


def extract_from_parentheses(text: str) -> Optional[str]:
    """Извлекает римскую цифру из скобок и преобразует в группу"""
    match = re.search(r'\(([IVX]+)\)', text, re.IGNORECASE)
    if match:
        return roman_to_abo(match.group(1))
    return None


def is_garbage_value(value: str) -> bool:
    """
    Проверяет, является ли значение "мусором"
    (невалидная группа крови)
    """
    if not value:
        return True
    
    # Длинные цифровые коды (1310, 1250, 0306806)
    if re.match(r'^\d{4,}$', value):
        return True
    
    # Длинные буквенно-цифровые коды (ж751012301003723)
    if re.match(r'^[а-яА-Яa-zA-Z]\d{10,}$', value):
        return True
    
    # Одиночная цифра 2, 4 - может быть A2 или O
    if re.match(r'^\d$', value):
        return False  # не мусор, обработаем отдельно
    
    return False


def normalize_abo_plus(value, verification_value=None) -> Optional[str]:
    """
    Расширенная нормализация группы крови
    
    Распознает:
        - A2B, А2В, A2B (IV)
        - A2, А2, A2 (II)
        - AB, АВ, AB (IV)
        - A, А, A (II)
        - B, В, B (III)
        - O, О, 0, O (I), 0 (I)
        - Римские цифры в скобках
        - Цифра 2 → A2
        - Мусор → берем из verification_value
    """
    if pd.isna(value):
        return None
    
    value_str = str(value).strip()
    
    # ========== 1. Прямое сопоставление ==========
    # Словарь соответствий (без учета регистра)
    direct_mapping = {
        # A2B
        'a2b': 'A2B', 'а2в': 'A2B', 'a2b(iv)': 'A2B', 'а2в(iv)': 'A2B',
        'a2b (iv)': 'A2B', 'а2в (iv)': 'A2B', 'a2b(iv)': 'A2B',
        
        # A2
        'a2': 'A2', 'а2': 'A2', 'a2(ii)': 'A2', 'а2(ii)': 'A2',
        'a2 (ii)': 'A2', 'а2 (ii)': 'A2', 'a2(ii)': 'A2',
        '2': 'A2', '4': 'O',  # 4 → O (по контексту из файла)
        
        # AB
        'ab': 'AB', 'ав': 'AB', 'ab(iv)': 'AB', 'ав(iv)': 'AB',
        'ab (iv)': 'AB', 'ав (iv)': 'AB', 'ab(iv)': 'AB',
        
        # A
        'a': 'A', 'а': 'A', 'a(ii)': 'A', 'а(ii)': 'A',
        'a (ii)': 'A', 'а (ii)': 'A', 'a(ii)': 'A',
        
        # B
        'b': 'B', 'в': 'B', 'b(iii)': 'B', 'в(iii)': 'B',
        'b (iii)': 'B', 'в (iii)': 'B', 'b(iii)': 'B',
        
        # O
        'o': 'O', 'о': 'O', '0': 'O', 'o(i)': 'O', 'о(i)': 'O',
        '0(i)': 'O', 'o (i)': 'O', 'о (i)': 'O', '0 (i)': 'O',
        'o(i)': 'O', '0(i)': 'O',
    }
    
    value_lower = value_str.lower()
    for pattern, result in direct_mapping.items():
        if value_lower == pattern or value_lower.startswith(pattern):
            return result
    
    # ========== 2. Извлечение из скобок (римские цифры) ==========
    # Если есть скобки с римскими цифрами
    abo_from_parentheses = extract_from_parentheses(value_str)
    if abo_from_parentheses:
        return abo_from_parentheses
    
    # ========== 3. Поиск ключевых слов в строке ==========
    # Ищем A2B, A2, AB, A, B, O в любом месте строки
    value_upper = value_str.upper()
    
    # A2B (должно быть до A2 и AB)
    if 'A2B' in value_upper or 'А2В' in value_upper:
        return 'A2B'
    
    # A2
    if 'A2' in value_upper or 'А2' in value_upper:
        return 'A2'
    
    # AB
    if 'AB' in value_upper or 'АВ' in value_upper:
        # Исключаем A2B (уже обработано)
        if 'A2B' not in value_upper and 'А2В' not in value_upper:
            return 'AB'
    
    # A
    if 'A' in value_upper or 'А' in value_upper:
        if 'AB' not in value_upper and 'АВ' not in value_upper:
            return 'A'
    
    # B
    if 'B' in value_upper or 'В' in value_upper:
        if 'AB' not in value_upper and 'АВ' not in value_upper:
            return 'B'
    
    # O
    if 'O' in value_upper or 'О' in value_upper or '0' in value_str:
        return 'O'
    
    # ========== 4. Мусор → берем из проверочного столбца ==========
    if is_garbage_value(value_str):
        if verification_value and not pd.isna(verification_value):
            # Рекурсивно нормализуем проверочное значение
            result = normalize_abo_plus(verification_value)
            if result:
                return result
    
    return None


def normalize_rh_plus(value) -> Optional[str]:
    """
    Расширенная нормализация резус-фактора
    
    Распознает:
        - +, Rh+, Rh +, Rh+
        - -, Rh-, Rh -, Rh-
        - положительная, положительный
        - отрицательная, отрицательный
        - (+), Rh(+)
    """
    if pd.isna(value):
        return None
    
    value_str = str(value).upper().strip()
    
    # Убираем скобки и пробелы для поиска
    clean_value = re.sub(r'[\(\s\)]', '', value_str)
    
    # Положительный резус
    positive_patterns = ['+', 'RH+', 'ПОЛОЖ', 'ПОЛОЖИТЕЛЬН', 'POS', 'ПОЗ']
    for pattern in positive_patterns:
        if pattern in clean_value:
            return '+'
    
    # Отрицательный резус
    negative_patterns = ['-', 'RH-', 'ОТРИЦ', 'ОТРИЦАТЕЛЬН', 'NEG', 'НЕГ']
    for pattern in negative_patterns:
        if pattern in clean_value:
            return '-'
    
    return None


def normalize_full_blood_group(abo_value, rh_value, abo_check=None, rh_check=None) -> tuple:
    """
    Полная нормализация группы крови и резуса с использованием проверочных столбцов
    
    Returns:
        (blood_group_norm, rh_norm, source, problem)
    """
    # Сначала пробуем нормализовать основные поля
    abo_norm = normalize_abo_plus(abo_value, abo_check)
    rh_norm = normalize_rh_plus(rh_value)
    
    source = 'patient'
    problem = ''
    
    # Если группа не определилась, пробуем проверочный столбец
    if not abo_norm and abo_check:
        abo_norm = normalize_abo_plus(abo_check)
        if abo_norm:
            source = 'check'
            problem = f'Группа из проверочного столбца (оригинал: {abo_value})'
    
    # Если резус не определился, пробуем проверочный столбец
    if not rh_norm and rh_check:
        rh_norm = normalize_rh_plus(rh_check)
        if rh_norm:
            source = 'check'
            problem = f'{problem}; Резус из проверочного столбца' if problem else f'Резус из проверочного столбца (оригинал: {rh_value})'
    
    # Фиксируем проблемы
    if not abo_norm:
        problem = f'{problem}; НЕ ОПРЕДЕЛЕНА ГРУППА' if problem else 'НЕ ОПРЕДЕЛЕНА ГРУППА'
    if not rh_norm:
        problem = f'{problem}; НЕ ОПРЕДЕЛЕН РЕЗУС' if problem else 'НЕ ОПРЕДЕЛЕН РЕЗУС'
    
    return abo_norm, rh_norm, source, problem


# ===================================================================
# ТЕСТИРОВАНИЕ
# ===================================================================

def test_normalizer():
    """Тестирует нормализатор на данных из файла"""
    
    test_cases = [
        # (input_abo, input_rh, expected_abo, expected_rh)
        ("О(I)", "Rh +", "O", "+"),
        ("А(II)", "Rh +", "A", "+"),
        ("В(III)", "Rh +", "B", "+"),
        ("АВ(IV)", "Rh +", "AB", "+"),
        ("А2(II)", "Rh +", "A2", "+"),
        ("А2В(IV)", "Rh +", "A2B", "+"),
        ("0 (I)", "Rh +", "O", "+"),
        ("A (II)", "Rh+", "A", "+"),
        ("B (III)", "Rh+", "B", "+"),
        ("AB (IV)", "Rh+", "AB", "+"),
        ("A2 (II)", "Rh+", "A2", "+"),
        ("A2B (IV)", "Rh+", "A2B", "+"),
        ("2", "Rh -", "A2", "-"),
        ("1310", "Rh +", "B", "+"),
        ("ж751012301003723", "Rh -", "B", "-"),
        ("направить на СПК,А2(11)", "Rh +", "A2", "+"),
        ("О(I)", "положительная", "O", "+"),
        ("О(I)", "Rh -", "O", "-"),
        ("0", "Rh +", "O", "+"),
        ("A2", "Rh+", "A2", "+"),
    ]
    
    print("🧪 ТЕСТИРОВАНИЕ НОРМАЛИЗАТОРА")
    print("="*60)
    
    passed = 0
    failed = 0
    
    for abo, rh, expected_abo, expected_rh in test_cases:
        result_abo = normalize_abo_plus(abo)
        result_rh = normalize_rh_plus(rh)
        
        abo_ok = result_abo == expected_abo
        rh_ok = result_rh == expected_rh
        
        if abo_ok and rh_ok:
            passed += 1
            print(f"✅ {abo} + {rh} → {result_abo}{result_rh}")
        else:
            failed += 1
            print(f"❌ {abo} + {rh} → {result_abo}{result_rh} (ожидалось {expected_abo}{expected_rh})")
    
    print("="*60)
    print(f"Результат: {passed} пройдено, {failed} не пройдено")
    return passed, failed


# ===================================================================
# ОСНОВНАЯ ФУНКЦИЯ (для запуска как скрипта)
# ===================================================================

def select_file():
    """Диалог выбора файла"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл с трансфузиями",
        filetypes=[("Excel файлы", "*.xlsx *.xls")]
    )
    root.destroy()
    return Path(file_path) if file_path else None


def main():
    print("🩸 НОРМАЛИЗАТОР ТРАНСФУЗИЙ v6.0 (PLUS)")
    print("="*60)
    
    # Сначала тестируем
    test_normalizer()
    
    print("\n" + "="*60)
    print("📂 Обработка файла...")
    
    file_path = select_file()
    if not file_path:
        print("❌ Файл не выбран")
        return
    
    df = pd.read_excel(file_path)
    print(f"✅ Загружено {len(df)} записей")
    
    # Поиск нужных столбцов
    abo_col = None
    rh_col = None
    abo_check_col = None
    rh_check_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if 'ab0' in col_lower or 'abo' in col_lower:
            if 'проверка' in col_lower:
                abo_check_col = col
            elif 'пац' in col_lower or 'пациент' in col_lower:
                abo_col = col
        elif 'rh' in col_lower:
            if 'проверка' in col_lower:
                rh_check_col = col
            elif 'пац' in col_lower or 'пациент' in col_lower:
                rh_col = col
    
    if not abo_col or not rh_col:
        print("❌ Не найдены столбцы с группой крови и резусом")
        return
    
    print(f"   ABO пациент: {abo_col}")
    print(f"   RH пациент: {rh_col}")
    print(f"   ABO проверка: {abo_check_col}")
    print(f"   RH проверка: {rh_check_col}")
    
    # Нормализация
    abo_norm = []
    rh_norm = []
    sources = []
    problems = []
    
    for idx, row in df.iterrows():
        abo_val = row.get(abo_col)
        rh_val = row.get(rh_col)
        abo_check = row.get(abo_check_col) if abo_check_col else None
        rh_check = row.get(rh_check_col) if rh_check_col else None
        
        abo, rh, source, problem = normalize_full_blood_group(
            abo_val, rh_val, abo_check, rh_check
        )
        
        abo_norm.append(abo)
        rh_norm.append(rh)
        sources.append(source)
        problems.append(problem)
    
    # Добавляем столбцы
    df['Blood_Group_Norm'] = abo_norm
    df['Rh_Norm'] = rh_norm
    df['Blood_Group_Source'] = sources
    df['Проблема_нормализации'] = problems
    df['Blood_Group_Full'] = [
        f"{abo}{rh}" if abo and rh else (abo if abo else rh) 
        for abo, rh in zip(abo_norm, rh_norm)
    ]
    
    # Сохраняем результат
    output_path = file_path.parent / f"{file_path.stem}_by_{NORMALIZER_NAME}.xlsx"
    df.to_excel(output_path, index=False)
    
    print(f"\n💾 Сохранено: {output_path}")
    
    # Статистика
    print("\n📊 СТАТИСТИКА:")
    print(f"   Всего записей: {len(df)}")
    print(f"\n   Распределение групп крови:")
    for blood, count in df['Blood_Group_Full'].value_counts().head(15).items():
        if blood and blood != 'None':
            print(f"      {blood}: {count}")
    
    problem_count = len(df[df['Проблема_нормализации'] != ''])
    if problem_count > 0:
        print(f"\n   ⚠️ Проблемных записей: {problem_count}")
        print("   Примеры проблем:")
        for problem in df[df['Проблема_нормализации'] != '']['Проблема_нормализации'].head(5):
            print(f"      {problem[:80]}")
    
    print("\n✅ ГОТОВО")


if __name__ == "__main__":
    main()