import pandas as pd
import streamlit as st
import io
from datetime import datetime
from openpyxl.styles import Alignment, Font

def normalize_column_name(col):
    """Нормализует имя столбца или вкладки для защиты от ошибок раскладки (РУС/ENG) и пробелов"""
    if not isinstance(col, str):
        return col
    col = col.strip().upper().replace(" ", "").replace("_", "")
    replacements = {
        'K': 'К', 'M': 'М', 'P': 'Р', 'A': 'А', 'C': 'С', 
        'O': 'О', 'T': 'Т', 'B': 'В', 'E': 'Е', 'H': 'Н', 'X': 'Х'
    }
    for eng, rus in replacements.items():
        col = col.replace(eng, rus)
    return col

def find_column(df, target_name):
    """Ищет столбец в листе Excel по его нормализованному имени"""
    norm_target = normalize_column_name(target_name)
    for col in df.columns:
        if normalize_column_name(col) == norm_target:
            return col
    return None

def find_sheet_smart(excel_file, target_name):
    """Умный поиск имени вкладки в файле Excel без привязки к регистру, пробелам и раскладке"""
    norm_target = normalize_column_name(target_name)
    for sheet in excel_file.sheet_names:
        if normalize_column_name(sheet) == norm_target:
            return sheet
    return None

def safe_int(value, default=0):
    """Безопасно преобразует значение в int, защищая от прочерков '-', nan и текста"""
    if pd.isna(value):
        return default
    val_str = str(value).strip()
    if val_str in ['-', '', 'nan', 'None', '.0']:
        return default
    try:
        return int(float(val_str))
    except ValueError:
        return default

def process_track_data(uploaded_file, exclude_curves=False):
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        sheet_assessment_name = find_sheet_smart(xl, 'ОЦЕНКАКМ')
        sheet_defects_name = find_sheet_smart(xl, 'ОТСТУПЛЕНИЯ')
        
        if not sheet_assessment_name:
            raise ValueError("В исходном файле не найдена вкладка 'Оценка КМ'.")
        if not sheet_defects_name:
            raise ValueError("В исходном файле не найдена вкладка 'Отступления'.")
            
        df_assessment = xl.parse(sheet_assessment_name)
        df_defects = xl.parse(sheet_defects_name)
    except Exception as e:
        raise ValueError(f"Ошибка при чтении файла Excel: {e}")

    # --- Сбор и нормализация столбцов 'Оценка КМ' ---
    col_kodnapr = find_column(df_assessment, 'КОДНАПР')
    col_path_assess = find_column(df_assessment, 'ПУТЬ')
    col_km_assess = find_column(df_assessment, 'КМ')
    col_pd_assess = find_column(df_assessment, 'ПД')
    col_rating = find_column(df_assessment, 'ОЦЕНКА')
    col_score_assess = find_column(df_assessment, 'БАЛЛ')
    col_lim_pass = find_column(df_assessment, 'СК_ОГР_ПАСС')
    col_lim_gruz = find_column(df_assessment, 'СК_ОГР_ГРУЗ')

    if not col_kodnapr or not col_km_assess:
        raise KeyError("В листе 'Оценка КМ' не найдены обязательные столбцы КОДНАПР или КМ.")

    df_assess_filtered = df_assessment[df_assessment[col_kodnapr] == 24701].copy()
    if df_assess_filtered.empty:
        raise ValueError("В листе 'Оценка КМ' не найдено записей с КОДНАПР == 24701")

    df_assess_filtered['MATCH_ПУТЬ'] = df_assess_
