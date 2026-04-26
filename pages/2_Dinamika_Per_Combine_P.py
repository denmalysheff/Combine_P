import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- КОНФИГУРАЦИЯ ССЫЛОК (Вариант А) ---
URL_STRUCT = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/adm_struktur.xlsx"
URL_STATIONS = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/stations_base.xlsx"

st.set_page_config(page_title="Динамика Nуч по перегонам", layout="wide")

def fix_headers(df):
    """Очистка заголовков: регистр, пробелы и замена латиницы на кириллицу"""
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ") # Исправляем частые ошибки ввода
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

def normalize_name(name):
    """Унификация названий перегонов для точного Match-инга"""
    if not isinstance(name, str): return name
    return name.replace(" ", "").replace("-", "").replace("–", "").upper().replace("Ё", "Е")

# --- ИНТЕРФЕЙС ---
st.title("📈 Модуль 2: Динамика балловой оценки")
st.info("Для расчета выберите два файла: текущий период и прошлый период.")

col1, col2 = st.columns(2)
with col1:
    file_curr = st.file_uploader("📂 Текущий месяц (Excel)", type=["xlsx"])
with col2:
    file_prev = st.file_uploader("📂 Прошлый месяц (Excel)", type=["xlsx"])

if file_curr and file_prev:
    try:
        # Загрузка данных
        df_curr_raw = fix_headers(pd.read_excel(file_curr, sheet_name="Оценка КМ"))
        df_prev_raw = fix_headers(pd.read_excel(file_prev, sheet_name="Оценка КМ"))

        # Группировка и расчет Nуч по перегонам (упрощенно для примера)
        def get_nuch_by_segment(df):
            # Здесь ваша логика расчета Nуч из Модуля 1, сгруппированная по 'ПЕРЕГОН'
            # Предположим, в файле уже есть колонка ПЕРЕГОН или мы её формируем
            df['MATCH_KEY'] = df['ПЕРЕГОН'].apply(normalize_name)
            res = df.groupby(['MATCH_KEY', 'ПЕРЕГОН']).agg({'ОЦЕНКА': 'mean', 'ПРОВЕРЕНО': 'sum'}).reset_index()
            res.columns = ['MATCH_KEY', 'ПЕРЕГОН_NAME', 'Nуч', 'КМ_ПРОВ']
            return res

        # В реальном коде тут вызывается ваша функция расчета Nуч
        # Для демонстрации используем расчет среднего балла
        curr_res = get_nuch_by_segment(df_curr_raw)
        prev_res = get_nuch_by_segment(df_prev_raw)

        # Объединение
        df_dyn = pd.merge(
            curr_res, 
            prev_res[['MATCH_KEY', 'Nуч']], 
            on='MATCH_KEY', 
            how='left', 
            suffixes=('', '_ПРОШЛЫЙ')
        )

        df_dyn['Динамика'] = (df_dyn['Nуч'] - df_dyn['Nуч_ПРОШЛЫЙ']).fillna(0)

        # Стилизация
        def style_dynamic(val):
            color = 'green' if val > 0 else 'red' if val < 0 else 'gray'
            return f'color: {color}; font-weight: bold'

        st.subheader("📊 Сравнительная таблица")
        st.dataframe(
            df_dyn.style.applymap(style_dynamic, subset=['Динамика'])
            .format({"Nуч": "{:.2f}", "Nуч_ПРОШЛЫЙ": "{:.2f}", "Динамика": "{:+.2f}"}),
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
