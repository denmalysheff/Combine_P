import streamlit as st
import pandas as pd
import io
import requests

# --- КОНФИГУРАЦИЯ ---
URL_STRUCT = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/adm_struktur.xlsx"

st.set_page_config(page_title="Расчет Nуч — Аналитика ПЧ-22", layout="wide")

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
        st.error(f"❌ Ошибка загрузки справочника: {e}")
        return None

def calculate_metrics(group_name, group_data, level, plan_km=0):
    # Принудительно переводим в числа для корректного расчета
    checked = pd.to_numeric(group_data["ПРОВЕРЕНО"], errors='coerce').fillna(0)
    scores = pd.to_numeric(group_data["ОЦЕНКА"], errors='coerce').fillna(0)
    
    fact_km = checked.sum()
    
    # Распределение по оценкам
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

# --- ИНТЕРФЕЙС ---
st.title("📊 Расчет балловой оценки (Nуч)")
st.markdown("---")

df_struct = load_admin_structure(URL_STRUCT)

if df_struct is not None:
    st.sidebar.success("✅ Справочник структуры загружен")
    uploaded_file = st.sidebar.file_uploader("Загрузите файл 'Оценка КМ'", type=["xlsx"])
    
    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, sheet_name="Оценка КМ")
            df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]
            
            # Фильтр направлений (основные ходы)
            main_codes = ['24701', '24602', '24603']
            df_eval = df_raw[df_raw["КОДНАПР"].astype(str).isin(main_codes)].copy()

            # План из справочника
            pd_plan_map = df_struct.groupby('ПД')['ПЛАН_ДЛИНА'].sum().to_dict()

            final_stats = []
            # Расчет по Линейным участкам (ПД)
            for pd_id, group in df_eval.groupby("ПД"):
                p_km = pd_plan_map.get(pd_id, 0)
                final_stats.append(calculate_metrics(f"ПД-{pd_id}", group, "Линейный", p_km))

            # Расчет по Групповым 
            groups_config = {
                "ПЧ": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
                "ПЧЗ Юг": [1, 2, 3, 4, 5, 12],
                "ПЧЗ Запад": [6, 7, 8, 9, 10, 11, 13, 14, 15],
                "ПЧУ-1": [1, 2, 3]
                "ПЧУ-2": [4, 5, 12]
                "ПЧУ-3": [7, 8, 13]
                "ПЧУ-4": [9, 10, 11]
                "ПЧУ-5": [6, 14, 15]
            
            }
            
            for g_name, pds in groups_config.items():
                g_data = df_eval[df_eval["ПД"].isin(pds)]
                g_plan = sum([pd_plan_map.get(p, 0) for p in pds])
                if not g_data.empty:
                    final_stats.append(calculate_metrics(g_name, g_data, "Групповой", g_plan))

            results_df = pd.DataFrame(final_stats)

            # --- ВЫВОД ---
            tab1, tab2 = st.tabs(["📋 Итоги расчета", "🔍 Детальная проверка"])

            with tab1:
                try:
                    # Попытка стилизации (требует matplotlib)
                    st.dataframe(
                        results_df.style.background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3, vmax=5)
                        .background_gradient(subset=['Полнота %'], cmap='YlOrRd', vmin=80, vmax=100),
                        use_container_width=True
                    )
                except Exception:
                    # Если matplotlib еще не установился, выводим без цвета
                    st.dataframe(results_df, use_container_width=True)

            with tab2:
                st.subheader("Сверка фактического прохода с планом")
                path_fact = df_eval.groupby(['КОДНАПР', 'ПУТЬ', 'ПД'])['ПРОВЕРЕНО'].sum().reset_index()
                path_plan = df_struct.groupby(['НАПРАВЛЕНИЕ', 'ПУТЬ', 'ПД'])['ПЛАН_ДЛИНА'].sum().reset_index()
                
                detail_check = path_plan.merge(
                    path_fact, left_on=['НАПРАВЛЕНИЕ','ПУТЬ','ПД'], 
                    right_on=['КОДНАПР','ПУТЬ','ПД'], how='left'
                ).fillna(0)
                detail_check['ДЕФИЦИТ'] = (detail_check['ПЛАН_ДЛИНА'] - detail_check['ПРОВЕРЕНО']).round(3)
                st.dataframe(detail_check, use_container_width=True)

        except Exception as e:
            st.error(f"❌ Ошибка обработки данных: {e}")
else:
    st.info("Ожидание подключения справочника структур...")
