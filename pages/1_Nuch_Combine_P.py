import streamlit as st
import pandas as pd
import io
import requests

# --- КОНФИГУРАЦИЯ ---
URL_STRUCT = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/adm_struktur.xlsx"

st.set_page_config(page_title="Расчет Nуч", layout="wide")

@st.cache_data
def load_admin_structure(url):
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status() 
        df = pd.read_excel(io.BytesIO(response.content), engine='openpyxl')
        df.columns = [str(col).strip().upper() for col in df.columns]
        if 'КМКОН' in df.columns and 'КМНАЧ' in df.columns:
            df['ПЛАН_ДЛИНА'] = (df['КМКОН'] - df['КМНАЧ']).abs()
        return df
    except Exception as e:
        st.error(f"Ошибка справочника: {e}")
        return None

def calculate_metrics(group_name, group_data, level, plan_km=0):
    checked = pd.to_numeric(group_data["ПРОВЕРЕНО"], errors='coerce').fillna(0)
    scores = pd.to_numeric(group_data["ОЦЕНКА"], errors='coerce').fillna(0)
    fact_km = checked.sum()
    
    km_5 = checked[scores == 5].sum()
    km_4 = checked[scores == 4].sum()
    km_3 = checked[scores == 3].sum()
    km_2 = checked[scores == 2].sum()

    n_uch = 0
    if fact_km > 0:
        n_uch = (km_5*5 + km_4*4 + km_3*3 - km_2*5) / fact_km

    return {
        "Уровень": level,
        "Группа": group_name,
        "Nуч": round(n_uch, 2),
        "Проверено (км)": round(fact_km, 3),
        "План (км)": round(plan_km, 3),
        "Полнота %": round((fact_km / plan_km * 100), 1) if plan_km > 0 else 0,
        "Отл": round(km_5, 3),
        "Хор": round(km_4, 3),
        "Удов": round(km_3, 3),
        "Неуд": round(km_2, 3)
    }

def style_results(df):
    """Кастомная стилизация для разделения уровней управления"""
    def apply_row_style(row):
        style = [''] * len(row)
        if row['Уровень'] == 'Дистанция':
            return ['background-color: #1f4e78; color: white; font-weight: bold; font-size: 16px'] * len(row)
        elif row['Уровень'] == 'Зам. ПЧ':
            return ['background-color: #2e75b6; color: white; font-weight: bold'] * len(row)
        elif row['Уровень'] == 'ПЧУ':
            return ['background-color: #deeaf6; color: black; font-weight: bold'] * len(row)
        return style

    return df.style.apply(apply_row_style, axis=1)\
        .background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3, vmax=5)\
        .background_gradient(subset=['Полнота %'], cmap='YlOrRd', vmin=80, vmax=100)

# --- ИНТЕРФЕЙС ---
st.title("📊 Аналитика балловой оценки (Nуч)")

df_struct = load_admin_structure(URL_STRUCT)

# Загрузка файла на главном экране
st.info("📌 Загрузите файл 'Оценка КМ' для начала анализа")
uploaded_file = st.file_uploader("Выбор файла Excel", type=["xlsx"], label_visibility="collapsed")

if uploaded_file and df_struct is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name="Оценка КМ")
        df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]
        
        main_codes = ['24701', '24602', '24603']
        df_eval = df_raw[df_raw["КОДНАПР"].astype(str).isin(main_codes)].copy()
        pd_plan_map = df_struct.groupby('ПД')['ПЛАН_ДЛИНА'].sum().to_dict()
        
        final_stats = []

        # 1. Линейный (ПД)
        for pd_id in sorted(df_eval["ПД"].unique()):
            group = df_eval[df_eval["ПД"] == pd_id]
            p_km = pd_plan_map.get(pd_id, 0)
            final_stats.append(calculate_metrics(f"ПД-{pd_id}", group, "Линейный", p_km))

        # 2. Групповой (ПЧУ)
        pchu_config = {
            "ПЧУ-1": [1, 2, 3], "ПЧУ-2": [4, 5, 12], "ПЧУ-3": [7, 8, 13],
            "ПЧУ-4": [9, 10, 11], "ПЧУ-5": [6, 14, 15]
        }
        for name, pds in pchu_config.items():
            g_data = df_eval[df_eval["ПД"].isin(pds)]
            g_plan = sum([pd_plan_map.get(p, 0) for p in pds])
            final_stats.append(calculate_metrics(name, g_data, "ПЧУ", g_plan))

        # 3. Руководство (ПЧЗ / ПЧ)
        adm_config = {
            "ПЧЗ Юг": [1, 2, 3, 4, 5, 12],
            "ПЧЗ Запад": [6, 7, 8, 9, 10, 11, 13, 14, 15],
            "ПЧ (Дистанция)": list(range(1, 16))
        }
        for name, pds in adm_config.items():
            g_data = df_eval[df_eval["ПД"].isin(pds)]
            g_plan = sum([pd_plan_map.get(p, 0) for p in pds])
            lvl = "Дистанция" if "ПЧ (" in name else "Зам. ПЧ"
            final_stats.append(calculate_metrics(name, g_data, lvl, g_plan))

        results_df = pd.DataFrame(final_stats)

        # --- ВКЛАДКИ С РЕЗУЛЬТАТАМИ ---
        tabs = st.tabs(["📊 Итоги Nуч", "✅ Отличные", "⭐ Хорошие", "⚠️ Удовл.", "🚨 Неуд (Причины)"])
        
        with tabs[0]:
            st.dataframe(style_results(results_df), use_container_width=True, height=600)

        # Функционал для фильтрации по оценкам
        for i, score in enumerate([5, 4, 3, 2]):
            with tabs[i+1]:
                subset = df_eval[df_eval["ОЦЕНКА"] == score].copy()
                if subset.empty:
                    st.write(f"Километры с оценкой {score} не найдены.")
                else:
                    if score == 2:
                        st.warning("⚠️ Требуется немедленное устранение причин неудовлетворительной оценки!")
                        # Здесь можно добавить логику вывода причин, если в файле есть колонки с типами отступлений
                        cols_to_show = ["ПД", "КМ", "ПК", "ПУТЬ", "ПРОВЕРЕНО", "БАЛЛ"]
                        st.dataframe(subset[[c for c in cols_to_show if c in subset.columns]], use_container_width=True)
                    else:
                        st.dataframe(subset, use_container_width=True)

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
