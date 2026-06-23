import pandas as pd
import streamlit as st
import io
from datetime import datetime
from openpyxl.styles import Alignment, Font

# Настройка страницы Streamlit
st.set_page_config(page_title="Формирование плана устранения отступлений", layout="wide")

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

    df_assess_filtered['MATCH_ПУТЬ'] = df_assess_filtered[col_path_assess].apply(safe_int)
    df_assess_filtered['MATCH_КМ'] = df_assess_filtered[col_km_assess].apply(safe_int)

    def make_speed_limit(row):
        p = str(row[col_lim_pass]).split('.')[0].strip() if col_lim_pass and pd.notna(row[col_lim_pass]) else "-"
        g = str(row[col_lim_gruz]).split('.')[0].strip() if col_lim_gruz and pd.notna(row[col_lim_gruz]) else "-"
        if p in ['0', '0.0', '', 'nan', '-']: p = "-"
        if g in ['0', '0.0', '', 'nan', '-']: g = "-"
        if p == "-" and g == "-": return ""
        return f"{p}/{g}"

    df_assess_filtered['ОГРАНИЧЕНИЕ_СКОРОСТИ'] = df_assess_filtered.apply(make_speed_limit, axis=1)

    # --- Сбор и нормализация столбцов 'Отступления' ---
    col_path_def = find_column(df_defects, 'ПУТЬ')
    col_km_def = find_column(df_defects, 'КМ М') or find_column(df_defects, 'КМ')
    col_meter = find_column(df_defects, 'МЕТР') or find_column(df_defects, 'М')
    col_defect = find_column(df_defects, 'ОТСТУПЛЕНИЕ')
    col_degree = find_column(df_defects, 'СТЕПЕНЬ')
    col_ampl = find_column(df_defects, 'АМПЛИТУДА')
    col_len = find_column(df_defects, 'ДЛИНА')
    col_is = find_column(df_defects, 'ИС')
    col_str = find_column(df_defects, 'СТРЕЛКА')
    col_obk = find_column(df_defects, 'ОБК')
    col_most = find_column(df_defects, 'МОСТ')
    col_pr = find_column(df_defects, 'PR_PREDUPR')
    col_score_def = find_column(df_defects, 'БАЛЛ')

    if not col_km_def or not col_degree:
        raise KeyError("В листе 'Отступления' не найдены столбцы КМ (КМ М) или СТЕПЕНЬ.")

    df_defects['MATCH_ПУТЬ'] = df_defects[col_path_def].apply(safe_int)
    df_defects['MATCH_КМ'] = df_defects[col_km_def].apply(safe_int)
    df_defects['INT_СТЕПЕНЬ'] = df_defects[col_degree].apply(safe_int)

    if col_str:
        df_defects['FILTER_СТРЕЛКА'] = df_defects[col_str].apply(safe_int)
        df_defects = df_defects[df_defects['FILTER_СТРЕЛКА'] == 0]

    if exclude_curves and col_defect:
        df_defects['TMP_DEF_NAME'] = df_defects[col_defect].astype(str).str.strip().str.upper()
        forbidden_defects = ['ПРУ', 'ДНПРОФ', 'КРИВАЯ', 'АНП']
        df_defects = df_defects[~df_defects['TMP_DEF_NAME'].isin(forbidden_defects)]

    df_defects['IS_2_PR'] = (df_defects['INT_СТЕПЕНЬ'] == 2) & (df_defects[col_pr].astype(str).str.contains('1', na=False) if col_pr else False)
    df_defects['IS_2_REGULAR'] = (df_defects['INT_СТЕПЕНЬ'] == 2) & (~df_defects['IS_2_PR'])
    df_defects['IS_3'] = df_defects['INT_СТЕПЕНЬ'] == 3
    df_defects['IS_4'] = df_defects['INT_СТЕПЕНЬ'] == 4

    df_valid_defects = df_defects[df_defects['INT_СТЕПЕНЬ'].isin([2, 3, 4])].copy()

    def aggregate_km_defects(group):
        count_2 = group['IS_2_REGULAR'].sum()
        count_2_3 = group['IS_2_PR'].sum()
        count_3 = group['IS_3'].sum()
        count_4 = group['IS_4'].sum()
        
        list_desc = []
        text_rows = group[group['IS_3'] | group['IS_4'] | group['IS_2_PR']]
        
        for _, row in text_rows.iterrows():
            m_val = str(safe_int(row[col_meter])) if col_meter else "0"
            def_type = str(row[col_defect]).strip() if col_defect else "ОТСТ"
            deg_str = "2к3ст" if row['IS_2_PR'] else f"{safe_int(row['INT_СТЕПЕНЬ'])}ст"
            ampl = str(safe_int(row[col_ampl])) if col_ampl else "0"
            length = str(safe_int(row[col_len])) if col_len else "0"
            def_score = str(safe_int(row[col_score_def])) if col_score_def else "0"
            
            tags = []
            if col_obk and pd.notna(row[col_obk]) and str(row[col_obk]).strip() not in ['0', '0.0', '', '-']: tags.append("обк")
            if col_most and pd.notna(row[col_most]) and str(row[col_most]).strip() not in ['0', '0.0', '', '-']: tags.append("м")
            if col_is and pd.notna(row[col_is]) and str(row[col_is]).strip() not in ['0', '0.0', '', '-']: tags.append("ис")
            
            tag_str = " " + " ".join(tags) if tags else ""
            desc = f"{m_val}-{def_type} {deg_str} {ampl}/{length}{tag_str.lower()} {def_score}б."
            list_desc.append(desc)
            
        final_text = ", ".join(list_desc) if list_desc else ""
        
        return pd.Series({
            'КОЛ_ВО_2': count_2,
            'КОЛ_ВО_2_3': count_2_3,
            'КОЛ_ВО_3': count_3,
            'КОЛ_ВО_4': count_4,
            'ПЕРЕЧЕНЬ_ОТСТУПЛЕНИЙ': final_text
        })

    if not df_valid_defects.empty:
        df_aggregated = df_valid_defects.groupby(['MATCH_ПУТЬ', 'MATCH_КМ']).apply(aggregate_km_defects, include_groups=False).reset_index()
        result_df = pd.merge(df_assess_filtered, df_aggregated, on=['MATCH_ПУТЬ', 'MATCH_КМ'], how='left')
    else:
        result_df = df_assess_filtered.copy()
        result_df['КОЛ_ВО_2'] = 0
        result_df['КОЛ_ВО_2_3'] = 0
        result_df['КОЛ_ВО_3'] = 0
        result_df['КОЛ_ВО_4'] = 0
        result_df['ПЕРЕЧЕНЬ_ОТСТУПЛЕНИЙ'] = ""

    for col in ['КОЛ_ВО_2', 'КОЛ_ВО_2_3', 'КОЛ_ВО_3', 'КОЛ_ВО_4']:
        result_df[col] = result_df[col].fillna(0).astype(int)
    result_df['ПЕРЕЧЕНЬ_ОТСТУПЛЕНИЙ'] = result_df['ПЕРЕЧЕНЬ_ОТСТУПЛЕНИЙ'].fillna("")

    output_data = {
        'КОДНАПР': result_df[col_kodnapr],
        'ПД': result_df[col_pd_assess] if col_pd_assess else "",
        'ПУТЬ': result_df['MATCH_ПУТЬ'],
        'КМ': result_df['MATCH_КМ'],
        'ОЦЕНКА': result_df[col_rating],
        'БАЛЛ': result_df[col_score_assess],
        'КОЛ-ВО 2 СТЕПЕНЕЙ': result_df['КОЛ_ВО_2'],
        'КОЛ-ВО 2 К 3 СТЕПЕНЕЙ': result_df['КОЛ_ВО_2_3'],
        'КОЛ-ВО 3 СТЕПЕНЕЙ': result_df['КОЛ_ВО_3'],
        'КОЛ-ВО 4 СТЕПЕНЕЙ': result_df['КОЛ_ВО_4'],
        'ОГРАНИЧЕНИЕ СКОРОСТИ': result_df['ОГРАНИЧЕНИЕ_СКОРОСТИ'],
        'ПЕРЕЧЕНЬ ОТСТУПЛЕНИЙ': result_df['ПЕРЕЧЕНЬ_ОТСТУПЛЕНИЙ']
    }
    final_df = pd.DataFrame(output_data)

    def pd_sheet_sort_key(name):
        digits = ''.join(c for c in str(name) if c.isdigit())
        return int(digits) if digits else 999

    final_df['ПД'] = final_df['ПД'].astype(str).str.strip()
    unique_pds = [x for x in final_df['ПД'].unique() if x not in ['', 'nan', 'None']]
    unique_pds_sorted = sorted(unique_pds, key=pd_sheet_sort_key)

    # Запись в буфер памяти вместо жесткого диска
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for pd_name in unique_pds_sorted:
            df_pd = final_df[final_df['ПД'] == pd_name].copy()
            df_pd.sort_values(by=['ПУТЬ', 'КМ'], ascending=[True, True], inplace=True)
            
            sheet_title = f"ПД-{pd_name}" if not str(pd_name).startswith('ПД') else str(pd_name)
            df_pd.to_excel(writer, sheet_name=sheet_title, index=False)
            
            workbook = writer.book
            worksheet = workbook[sheet_title]
            
            worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
            worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
            worksheet.sheet_properties.pageSetUpPr.fitToPage = True
            worksheet.page_setup.fitToWidth = 1
            worksheet.page_setup.fitToHeight = 0
            
            align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)
            align_center = Alignment(horizontal='center', vertical='center')
            align_left = Alignment(horizontal='left', vertical='center')
            
            font_bold = Font(name='Arial', size=10, bold=True)
            font_normal = Font(name='Arial', size=10, bold=False)
            
            col_idx_rating = 5 
            col_idx_desc = 12
            
            worksheet.row_dimensions[1].height = 45
            for col_idx in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=1, column=col_idx)
                cell.alignment = align_header
                cell.font = font_bold
            
            for row_idx in range(2, worksheet.max_row + 1):
                rating_value = str(worksheet.cell(row=row_idx, column=col_idx_rating).value).strip()
                is_rating_2 = (rating_value == '2' or rating_value == '2.0')
                
                for col_idx in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    
                    if col_idx == col_idx_desc:
                        cell.alignment = align_left
                    else:
                        cell.alignment = align_center
                        
                    if is_rating_2:
                        cell.font = font_bold
                    else:
                        cell.font = font_normal

            for col in worksheet.columns:
                col_letter = col[0].column_letter
                max_data_len = 0
                for cell in col[1:]: 
                    if cell.value is not None:
                        max_data_len = max(max_data_len, len(str(cell.value)))
                
                if col_letter in ['A', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                    worksheet.column_dimensions[col_letter].width = 11
                elif col_letter == 'B':
                    worksheet.column_dimensions[col_letter].width = 8
                else:
                    worksheet.column_dimensions[col_letter].width = max(max_data_len + 3, 25)
                    
    processed_data = output.getvalue()
    return processed_data

# --- Веб-Интерфейс Streamlit ---
st.title("📋 План устранения отступлений")
st.subheader("Формирование списка отступлений по проходу вагона-путеизмерителя")

uploaded_file = st.file_uploader("Выберите исходный Excel-файл с вагона", type=["xlsx", "xls"])
exclude_curves = st.checkbox("Исключить оценку кривых (ПрУ, ДНпроф, Кривая, Анп)")

if uploaded_file is not None:
    if st.button("Сформировать и рассчитать план", type="primary"):
        with st.spinner("⏳ Выполняются вычисления... Пожалуйста, подождите."):
            try:
                # Обработка данных
                excel_data = process_track_data(uploaded_file, exclude_curves=exclude_curves)
                
                # Генерация имени файла
                current_time = datetime.now().strftime("%d-%m-%Y_%H-%M")
                file_title = f"Plan_ustr_otst_24701_{current_time}.xlsx"
                
                st.success("✅ Расчет успешно завершен!")
                
                # Кнопка для скачивания готового результата
                st.download_button(
                    label="📥 Скачать готовый план (Excel)",
                    data=excel_data,
                    file_name=file_title,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Ошибка при распознавании данных: {str(e)}")
