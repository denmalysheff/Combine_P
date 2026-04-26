import streamlit as st
import pandas as pd
import io
import requests
import urllib.parse
from datetime import datetime

# --- 1. НАСТРОЙКИ ПУТЕЙ ---
# Файлы подгружаются напрямую из вашего репозитория
GITHUB_BASE = "https://raw.githubusercontent.com/denmalysheff/Combine_P/main/"
URL_STRUCT = GITHUB_BASE + "adm_struktur.xlsx"
URL_STATIONS = GITHUB_BASE + "stations_base.xlsx"

st.set_page_config(page_title="Динамика Nуч по перегонам", layout="wide")

# --- 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

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
    """Универсальный загрузчик справочников с GitHub"""
    try:
        urls_to_try = [url, url.replace("/main/", "/refs/heads/main/")]
        for target_url in urls_to_try:
            parsed = list(urllib.parse.urlparse(target_url))
            parsed[2] = urllib.parse.quote(parsed[2])
            encoded_url = urllib.parse.urlunparse(parsed)
            res = requests.get(encoded_url, timeout=10)
            if res.status_code == 200:
                return fix_headers(pd.read_excel(io.BytesIO(res.content), engine='openpyxl'))
        return None
    except:
        return None

# --- 3. ЯДРО РАСЧЕТА ---

def process_file_data(file):
    """Загрузка файла и расчет Nуч по перегонам"""
    try:
        xl = pd.ExcelFile(file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet:
            st.error(f"Лист 'Оценка КМ' не найден в {file.name}")
            return None
        
        df = xl.parse(sheet)
        df = fix_headers(df)

        # Гибкий поиск колонок (учитываем КМ/KM и Перегон/Участок)
        seg_aliases = ["ПЕРЕГОН", "ПЕРЕГОН.", "УЧАСТОК", "SEGMENT"]
        km_aliases = ["КМ", "KM", "№КМ", "№KM", "№ КМ", "№ KM", "КМ.", "KM."]
        
        seg_col = next((c for c in df.columns if c in seg_aliases), None)
        km_col = next((c for c in df.columns if c in km_aliases), None)
        score_col = next((c for c in df.columns if c in ["ОЦЕНКА", "SCORE"]), None)
        check_col = next((c for c in df.columns if c in ["ПРОВЕРЕНО", "CHECKED"]), None)

        if not all([seg_col, km_col, score_col, check_col]):
            st.error(f"В файле {file.name} не найдены нужные колонки. Проверьте заголовки.")
            st.info(f"Доступные колонки: {list(df.columns)}")
            return None

        # Очистка типов данных
        df[score_col] = pd.to_numeric(df[score_col], errors='coerce')
        df[check_col] = pd.to_numeric(df[check_col], errors='coerce').fillna(0)

        def nuch_formula(group):
            fact_km = group[check_col].sum()
            if fact_km == 0: return 0
            # Баллы: 5, 4, 3 и штраф -5 за неуд (2)
            s5 = group[group[score_col] == 5][check_col].sum()
            s4 = group[group[score_col] == 4][check_col].sum()
            s3 = group[group[score_col] == 3][check_col].sum()
            s2 = group[group[score_col] == 2][check_col].sum()
            return round((s5*5 + s4*4 + s3*3 - s2*5) / fact_km, 2)

        res = df.groupby(seg_col).apply(lambda x: pd.Series({
            'Nуч': nuch_formula(x),
            'КМ_ПРОВ': round(x[check_col].sum(), 3)
        })).reset_index()
        
        res.columns = ['ПЕРЕГОН', 'Nуч', 'КМ_ПРОВ']
        return res
    except Exception as e:
        st.error(f"Ошибка при обработке {file.name}: {e}")
        return None

# --- 4. ИНТЕРФЕЙС ---

st.title("📈 Модуль 2: Динамика балловой оценки")

# Загрузка справочников
df_stations = load_base_data(URL_STATIONS)

col1, col2 = st.columns(2)
with col1:
    f_curr = st.file_uploader("📂 Текущий период", type=["xlsx"], key="c")
with col2:
    f_prev = st.file_uploader("📂 Прошлый период", type=["xlsx"], key="p")

if f_curr and f_prev:
    curr_data = process_file_data(f_curr)
    prev_data = process_file_data(f_prev)

    if curr_data is not None and prev_data is not None:
        # Объединение
        df_dyn = pd.merge(curr_data, prev_data, on='ПЕРЕГОН', how='left', suffixes=('', '_ПРОШЛ'))
        df_dyn['Динамика'] = (df_dyn['Nуч'] - df_dyn['Nуч_ПРОШЛ']).fillna(0)
        
        # Подтягиваем код станции
        if df_stations is not None:
            df_dyn['M'] = df_dyn['ПЕРЕГОН'].str.upper().str.strip()
            df_stations['M'] = df_stations['СТАНЦИЯ'].str.upper().str.strip()
            df_dyn = pd.merge(df_dyn, df_stations[['M', 'КОД СТАНЦИИ']], on='M', how='left').drop(columns=['M'])

        # Стилизация и вывод
        st.subheader("📊 Результаты сравнения")
        
        def color_dyn(val):
            if val > 0: return 'color: green; font-weight: bold'
            if val < 0: return 'color: red; font-weight: bold'
            return ''

        st.dataframe(
            df_dyn.sort_values('Динамика').style.format({
                "Nуч": "{:.2f}", "Nуч_ПРОШЛ": "{:.2f}", "Динамика": "{:+.2f}"
            })
            .background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3.5, vmax=5)
            .map(color_dyn, subset=['Динамика']),
            use_container_width=True
        )

        # Экспорт
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_dyn.to_excel(writer, index=False, sheet_name='Динамика')
        st.download_button("📥 Скачать отчет", output.getvalue(), "Dynamic_Report.xlsx")
