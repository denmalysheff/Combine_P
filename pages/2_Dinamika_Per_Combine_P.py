import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Анализ Nуч по КМ", layout="wide")

# --- 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def fix_headers(df):
    """Приведение заголовков к единому стандарту (регистр, очистка, кириллица)"""
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

def find_sheet(xl, target_name):
    """Поиск листа в Excel (регистронезависимый)"""
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

# --- 2. ЗАГРУЗКА СПРАВОЧНИКА СТРУКТУРЫ ---

@st.cache_data
def load_local_structure():
    """Загрузка adm_struktur.xlsx из корня проекта"""
    file_path = "adm_struktur.xlsx"
    if not os.path.exists(file_path):
        st.error(f"Критическая ошибка: Файл {file_path} не найден в корне скрипта!")
        return None
    try:
        df = pd.read_excel(file_path)
        df = fix_headers(df)
        # Убеждаемся, что КМ — это числа
        df['КМНАЧ'] = pd.to_numeric(df['КМНАЧ'], errors='coerce')
        df['КМКОН'] = pd.to_numeric(df['КМКОН'], errors='coerce')
        return df.dropna(subset=['КМНАЧ', 'КМКОН'])
    except Exception as e:
        st.error(f"Ошибка чтения справочника: {e}")
        return None

# --- 3. ЛОГИКА СОПОСТАВЛЕНИЯ ---

def get_nuch_by_structure(data_file, structure_df):
    """Сопоставление КМ из файла данных с диапазонами из структуры"""
    try:
        xl = pd.ExcelFile(data_file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet: return None
        
        df = xl.parse(sheet)
        df = fix_headers(df)

        # Ищем обязательные колонки в загруженном файле
        km_col = next((c for c in df.columns if c in ["КМ", "KM", "№КМ"]), None)
        score_col = next((c for c in df.columns if c in ["ОЦЕНКА", "SCORE"]), None)
        check_col = next((c for c in df.columns if c in ["ПРОВЕРЕНО", "CHECKED"]), None)

        if not all([km_col, score_col, check_col]):
            st.warning(f"В файле {data_file.name} не найдены нужные колонки (КМ, Оценка или Проверено)")
            return None

        # Подготовка данных
        df[km_col] = pd.to_numeric(df[km_col], errors='coerce')
        df[score_col] = pd.to_numeric(df[score_col], errors='coerce')
        df[check_col] = pd.to_numeric(df[check_col], errors='coerce').fillna(0)
        df = df.dropna(subset=[km_col, score_col])

        # ФУНКЦИЯ ПРИВЯЗКИ: Для каждой строки данных ищем перегон в структуре
        def find_segment(km):
            # Ищем строку в структуре, где КМ попадает в границы
            match = structure_df[(km >= structure_df['КМНАЧ']) & (km <= structure_df['КМКОН'])]
            if not match.empty:
                # Берем 'НАПРАВЛЕНИЕ' или 'ПЕРЕГОН' из справочника
                return match.iloc[0].get('НАПРАВЛЕНИЕ', 'Неизвестно')
            return "Вне структуры"

        df['ПЕРЕГОН_ИЗ_СПРАВОЧНИКА'] = df[km_col].apply(find_segment)

        # Расчет Nуч (формула РЖД)
        def nuch_calc(group):
            total_km = group[check_col].sum()
            if total_km == 0: return 0
            s5 = group[group[score_col] == 5][check_col].sum()
            s4 = group[group[score_col] == 4][check_col].sum()
            s3 = group[group[score_col] == 3][check_col].sum()
            s2 = group[group[score_col] == 2][check_col].sum()
            return round((s5*5 + s4*4 + s3*3 - s2*5) / total_km, 2)

        result = df.groupby('ПЕРЕГОН_ИЗ_СПРАВОЧНИКА').apply(lambda x: pd.Series({
            'Nуч': nuch_calc(x),
            'КМ_ПРОВ': round(x[check_col].sum(), 3)
        })).reset_index()

        return result

    except Exception as e:
        st.error(f"Ошибка: {e}")
        return None

# --- 4. ИНТЕРФЕЙС ---

st.title("📈 Расчет Nуч по локальному справочнику (КМ)")

struct = load_local_structure()

if struct is not None:
    st.success("✅ Справочник adm_struktur.xlsx загружен")
    
    col1, col2 = st.columns(2)
    with col1:
        f_curr = st.file_uploader("📂 Текущий период", type=["xlsx"], key="c")
    with col2:
        f_prev = st.file_uploader("📂 Прошлый период", type=["xlsx"], key="p")

    if f_curr and f_prev:
        curr_res = get_nuch_by_structure(f_curr, struct)
        prev_res = get_nuch_by_structure(f_prev, struct)

        if curr_res is not None and prev_res is not None:
            # Слияние по перегону из справочника
            final = pd.merge(curr_res, prev_res, on='ПЕРЕГОН_ИЗ_СПРАВОЧНИКА', how='left', suffixes=('', '_ПРОШЛ'))
            final['Динамика'] = final['Nуч'] - final['Nуч_ПРОШЛ'].fillna(0)
            
            st.subheader("📊 Анализ перегонов (на основе километража)")
            st.dataframe(
                final.sort_values('Динамика').style.format({
                    "Nуч": "{:.2f}", "Nуч_ПРОШЛ": "{:.2f}", "Динамика": "{:+.2f}"
                })
                .background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3, vmax=5),
                use_container_width=True
            )
