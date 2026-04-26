import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime

# --- КОНФИГУРАЦИЯ ---
URL_STRUCT = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/adm_struktur.xlsx"
URL_STATIONS = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/stations_base.xlsx"

st.set_page_config(page_title="Динамика Nуч по перегонам", layout="wide")

# --- 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def fix_headers(df):
    """Исправляет заголовки: регистр, пробелы, латиница -> кириллица"""
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

def find_sheet(xl, target_name):
    """Ищет лист в Excel, игнорируя регистр и пробелы"""
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

@st.cache_data
def load_base_data(url):
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.read_excel(io.BytesIO(response.content), engine='openpyxl')
        return fix_headers(df)
    except Exception as e:
        st.error(f"Ошибка загрузки справочника ({url}): {e}")
        return None

# --- 2. ЛОГИКА РАСЧЕТА ---

def process_file(file):
    """Загрузка и расчет Nуч по каждому перегону"""
    xl = pd.ExcelFile(file)
    sheet = find_sheet(xl, "Оценка КМ")
    if not sheet:
        st.error(f"Лист 'Оценка КМ' не найден в файле {file.name}")
        return None
    
    df = fix_headers(xl.parse(sheet))
    
    # Ищем колонку перегона (алиасы)
    seg_col = next((c for c in df.columns if c in ["ПЕРЕГОН", "ПЕРЕГОН.", "УЧАСТОК"]), None)
    if not seg_col:
        st.error(f"Колонка ПЕРЕГОН не найдена в {file.name}")
        return None

    def calc_group_nuch(group):
        check = pd.to_numeric(group["ПРОВЕРЕНО"], errors='coerce').fillna(0)
        score = pd.to_numeric(group["ОЦЕНКА"], errors='coerce').fillna(0)
        fact_km = check.sum()
        if fact_km == 0: return 0
        
        # Стандартная формула RZD: (5*км5 + 4*км4 + 3*км3 - 5*км2) / всего_км
        val_sum = (check[score == 5].sum() * 5 + 
                   check[score == 4].sum() * 4 + 
                   check[score == 3].sum() * 3 - 
                   check[score == 2].sum() * 5)
        return round(val_sum / fact_km, 2)

    res = df.groupby(seg_col).apply(lambda x: pd.Series({
        'Nуч': calc_group_nuch(x),
        'КМ_ПРОВ': round(x['ПРОВЕРЕНО'].sum(), 3)
    })).reset_index()
    
    res.columns = ['ПЕРЕГОН', 'Nуч', 'КМ_ПРОВ']
    return res

# --- 3. ОСНОВНОЙ ИНТЕРФЕЙС ---

st.title("📈 Модуль 2: Динамика балловой оценки")

# Загрузка справочника станций для кодов ЕСР
df_stations = load_base_data(URL_STATIONS)

col1, col2 = st.columns(2)
with col1:
    f_curr = st.file_uploader("📂 Текущий период (Excel)", type=["xlsx"], key="curr")
with col2:
    f_prev = st.file_uploader("📂 Прошлый период (Excel)", type=["xlsx"], key="prev")

if f_curr and f_prev:
    curr_data = process_file(f_curr)
    prev_data = process_file(f_prev)

    if curr_data is not None and prev_data is not None:
        # Мержим данные
        df_final = pd.merge(curr_data, prev_data, on='ПЕРЕГОН', how='left', suffixes=('', '_ПРОШЛ'))
        df_final['Динамика'] = (df_final['Nуч'] - df_final['Nуч_ПРОШЛ']).fillna(0)
        
        # Интеграция кодов станций (если справочник загружен)
        if df_stations is not None:
            # Пытаемся подтянуть код станции для перегона (по названию)
            # Примечание: логика Match-инга зависит от структуры вашего stations_base
            df_final = pd.merge(df_final, df_stations[['СТАНЦИЯ', 'КОД СТАНЦИИ']], 
                                left_on='ПЕРЕГОН', right_on='СТАНЦИЯ', how='left').drop(columns=['СТАНЦИЯ'])

        # Стилизация
        def color_dyn(val):
            if val > 0: return 'color: #008000; font-weight: bold'
            if val < 0: return 'color: #FF0000; font-weight: bold'
            return ''

        st.subheader("📋 Сравнительный анализ по перегонам")
        st.dataframe(
            df_final.style.format({
                "Nуч": "{:.2f}",
                "Nуч_ПРОШЛ": "{:.2f}",
                "Динамика": "{:+.2f}"
            })
            .background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3.5, vmax=5)
            .map(color_dyn, subset=['Динамика']),
            use_container_width=True
        )

        # Экспорт
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Динамика_Перегоны')
            
        st.download_button(
            label="📥 Скачать отчет в Excel",
            data=output.getvalue(),
            file_name=f"Dinamika_Per_{datetime.now().strftime('%d_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
