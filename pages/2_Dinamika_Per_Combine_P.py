import streamlit as st
import pandas as pd
import io
import requests
import urllib.parse
from datetime import datetime

# --- 1. КОНФИГУРАЦИЯ ПУТЕЙ ---
GITHUB_BASE = "https://raw.githubusercontent.com/denmalysheff/Combine_P/main/"
URL_STRUCT = GITHUB_BASE + "adm_struktur.xlsx"
URL_STATIONS = GITHUB_BASE + "stations_base.xlsx"

st.set_page_config(page_title="Динамика Nуч", layout="wide")

# --- 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def fix_headers(df):
    """Исправляет заголовки: регистр, пробелы, латиница -> кириллица"""
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

@st.cache_data
def load_base_data(url):
    """Загрузчик справочников с GitHub с защитой от 404"""
    try:
        # Пробуем разные варианты путей GitHub Raw
        urls_to_try = [
            url,
            url.replace("/main/", "/refs/heads/main/")
        ]
        
        response = None
        for target_url in urls_to_try:
            parsed = list(urllib.parse.urlparse(target_url))
            parsed[2] = urllib.parse.quote(parsed[2])
            encoded_url = urllib.parse.urlunparse(parsed)
            try:
                res = requests.get(encoded_url, timeout=10)
                if res.status_code == 200:
                    response = res
                    break
            except:
                continue

        if response and response.status_code == 200:
            return fix_headers(pd.read_excel(io.BytesIO(response.content), engine='openpyxl'))
        else:
            st.warning(f"Не удалось найти файл: {url.split('/')[-1]}")
            return None
    except Exception as e:
        st.error(f"Ошибка загрузки справочника: {e}")
        return None

def find_sheet(xl, target_name):
    """Поиск листа без учета регистра и пробелов"""
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

# --- 3. ЛОГИКА РАСЧЕТА ---

def calculate_segment_metrics(file):
    """Загрузка файла и расчет Nуч по каждому перегону"""
    try:
        xl = pd.ExcelFile(file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet:
            st.error(f"В файле {file.name} не найден лист 'Оценка КМ'")
            return None
        
        df = fix_headers(xl.parse(sheet))
        
        # Поиск ключевых колонок
        seg_col = next((c for c in df.columns if c in ["ПЕРЕГОН", "ПЕРЕГОН.", "УЧАСТОК"]), None)
        km_col = next((c for c in df.columns if c in ["КМ", "KM", "№ КМ"]), None)
        
        if not seg_col or not km_col:
            st.error(f"В файле {file.name} отсутствуют обязательные колонки (Перегон или КМ)")
            return None

        def nuch_formula(group):
            # Приведение к числам
            check = pd.to_numeric(group["ПРОВЕРЕНО"], errors='coerce').fillna(0)
            score = pd.to_numeric(group["ОЦЕНКА"], errors='coerce').fillna(0)
            fact_km = check.sum()
            if fact_km == 0: return 0
            
            # Веса согласно Модулю 1
            weighted_sum = (check[score == 5].sum() * 5 + 
                           check[score == 4].sum() * 4 + 
                           check[score == 3].sum() * 3 - 
                           check[score == 2].sum() * 5)
            return round(weighted_sum / fact_km, 2)

        # Агрегация по перегонам
        res = df.groupby(seg_col).apply(lambda x: pd.Series({
            'Nуч': nuch_formula(x),
            'КМ_ПРОВ': round(x['ПРОВЕРЕНО'].sum(), 3)
        })).reset_index()
        
        res.columns = ['ПЕРЕГОН', 'Nуч', 'КМ_ПРОВ']
        return res
    except Exception as e:
        st.error(f"Ошибка при обработке {file.name}: {e}")
        return None

# --- 4. ИНТЕРФЕЙС ---

st.title("📈 Модуль 2: Сравнительная динамика Nуч")
st.markdown("Сравнение балловой оценки между двумя периодами в разрезе перегонов.")

# Предварительная загрузка справочников
df_stations = load_base_data(URL_STATIONS)

col1, col2 = st.columns(2)
with col1:
    f_curr = st.file_uploader("📂 Текущий период (Excel)", type=["xlsx"], key="c")
with col2:
    f_prev = st.file_uploader("📂 Прошлый период (Excel)", type=["xlsx"], key="p")

if f_curr and f_prev:
    with st.spinner("Рассчитываю динамику..."):
        curr_res = calculate_segment_metrics(f_curr)
        prev_res = calculate_segment_metrics(f_prev)

    if curr_res is not None and prev_res is not None:
        # Объединение данных
        df_dyn = pd.merge(curr_res, prev_res[['ПЕРЕГОН', 'Nуч']], on='ПЕРЕГОН', 
                         how='left', suffixes=('', '_ПРОШЛ'))
        
        df_dyn['Динамика'] = (df_dyn['Nуч'] - df_dyn['Nуч_ПРОШЛ']).fillna(0)
        
        # Добавляем данные о станциях, если справочник доступен
        if df_stations is not None:
            # Нормализация для поиска
            df_dyn['MATCH'] = df_dyn['ПЕРЕГОН'].str.upper().str.strip()
            df_stations['MATCH'] = df_stations['СТАНЦИЯ'].str.upper().str.strip()
            
            df_dyn = pd.merge(df_dyn, df_stations[['MATCH', 'КОД СТАНЦИИ']], 
                             on='MATCH', how='left').drop(columns=['MATCH'])

        # Сортировка по ухудшению (самые проблемные сверху)
        df_dyn = df_dyn.sort_values(by='Динамика', ascending=True)

        # Отображение
        st.subheader("📊 Анализ изменений")
        
        def style_dynamic(val):
            if val > 0: return 'color: #008000; font-weight: bold' # Зеленый (рост)
            if val < 0: return 'color: #d32f2f; font-weight: bold' # Красный (падение)
            return 'color: #757575' # Серый (без изменений)

        st.dataframe(
            df_dyn.style.format({
                "Nуч": "{:.2f}",
                "Nуч_ПРОШЛ": "{:.2f}",
                "Динамика": "{:+.2f}",
                "КМ_ПРОВ": "{:.3f}"
            })
            .background_gradient(subset=['Nуч'], cmap='RdYlGn', vประกอบ=3.5, vmax=5)
            .map(style_dynamic, subset=['Динамика']),
            use_container_width=True,
            height=600
        )

        # Кнопка скачивания
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_dyn.to_excel(writer, index=False, sheet_name='Динамика')
        
        st.download_button(
            label="💾 Сохранить отчет в Excel",
            data=output.getvalue(),
            file_name=f"Dynamic_Nuch_{datetime.now().strftime('%d_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
