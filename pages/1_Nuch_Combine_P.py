import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime

# --- КОНФИГУРАЦИЯ ---
URL_STRUCT = "https://raw.githubusercontent.com/denmalysheff/Nuch/main/adm_struktur.xlsx"

st.set_page_config(page_title="Аналитика Nуч", layout="wide")

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
        st.error(f"Ошибка справочника структур: {e}")
        return None

def calculate_metrics(group_name, group_data, level, plan_km=0):
    checked = pd.to_numeric(group_data["ПРОВЕРЕНО"], errors='coerce').fillna(0)
    scores = pd.to_numeric(group_data["ОЦЕНКА"], errors='coerce').fillna(0)
    fact_km = checked.sum()
    
    km_5 = checked[scores == 5].sum()
    km_4 = checked[scores == 4].sum()
    km_3 = checked[scores == 3].sum()
    km_2 = checked[scores == 2].sum()

    n_uch = (km_5*5 + km_4*4 + km_3*3 - km_2*5) / fact_km if fact_km > 0 else 0

    return {
        "Уровень": level,
        "Группа": group_name,
        "Nуч": round(n_uch, 2),
        "Проверено (км)": round(fact_km, 3),
        "План (км)": round(plan_km, 3),
        "Полнота %": round((fact_km / plan_km * 100), 1) if plan_km > 0 else 0,
        "Отл": round(km_5, 3), "Хор": round(km_4, 3), "Удов": round(km_3, 3), "Неуд": round(km_2, 3)
    }

def style_results(df):
    def apply_row_style(row):
        if row['Уровень'] == 'Дистанция':
            return ['background-color: #1f4e78; color: white; font-weight: bold; font-size: 15px'] * len(row)
        elif row['Уровень'] == 'Зам. ПЧ':
            return ['background-color: #2e75b6; color: white; font-weight: bold'] * len(row)
        elif row['Уровень'] == 'Эксплуатационный':
            return ['background-color: #deeaf6; color: black; font-weight: bold'] * len(row)
        return [''] * len(row)

    return df.style.apply(apply_row_style, axis=1)\
        .background_gradient(subset=['Nуч'], cmap='RdYlGn', vmin=3, vmax=5)\
        .background_gradient(subset=['Полнота %'], cmap='YlOrRd', vmin=80, vmax=100)

# --- ИНТЕРФЕЙС ---
st.title("📊 Модуль 1: Расчет балловой оценки (Nуч)")

df_struct = load_admin_structure(URL_STRUCT)

st.markdown("### 📥 Шаг 1: Загрузка первичных данных")
uploaded_file = st.file_uploader("Загрузите Excel-файл (лист 'Оценка КМ')", type=["xlsx"])

if uploaded_file and df_struct is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name="Оценка КМ")
        df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]
        
        # Фильтрация и подготовка данных
        main_codes = ['24701', '24602', '24603']
        df_eval = df_raw[df_raw["КОДНАПР"].astype(str).isin(main_codes)].copy()
        pd_plan_map = df_struct.groupby('ПД')['ПЛАН_ДЛИНА'].sum().to_dict()
        
        final_stats = []

        # 1. ПД (Линейный уровень)
        for pd_id in sorted(df_eval["ПД"].unique()):
            group = df_eval[df_eval["ПД"] == pd_id]
            final_stats.append(calculate_metrics(f"ПД-{pd_id}", group, "Линейный", pd_plan_map.get(pd_id, 0)))

        # 2. ПЧУ (Эксплуатационный уровень)
        pchu_cfg = {"ПЧУ-1": [1,2,3], "ПЧУ-2": [4,5,12], "ПЧУ-3": [7,8,13], "ПЧУ-4": [9,10,11], "ПЧУ-5": [6,14,15]}
        for name, pds in pchu_cfg.items():
            g_data = df_eval[df_eval["ПД"].isin(pds)]
            g_plan = sum([pd_plan_map.get(p, 0) for p in pds])
            final_stats.append(calculate_metrics(name, g_data, "Эксплуатационный", g_plan))

        # 3. Руководство
        adm_cfg = {"ПЧЗ Юг": [1,2,3,4,5,12], "ПЧЗ Запад": [6,7,8,9,10,11,13,14,15], "ПЧ (Дистанция)": list(range(1,16))}
        for name, pds in adm_cfg.items():
            g_data = df_eval[df_eval["ПД"].isin(pds)]
            g_plan = sum([pd_plan_map.get(p, 0) for p in pds])
            lvl = "Дистанция" if "ПЧ (" in name else "Зам. ПЧ"
            final_stats.append(calculate_metrics(name, g_data, lvl, g_plan))

        results_df = pd.DataFrame(final_stats)

        # --- ПОДГОТОВКА EXCEL С ВЫРАВНИВАНИЕМ ---
        st.markdown("### 📝 Шаг 2: Анализ и экспорт")
        
        current_date = datetime.now().strftime("%d_%m_%Y")
        file_name = f"Nuch_Report_{current_date}.xlsx"
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='Итоги_Nуч', index=False)
            
            # Колонки для вкладок КМ
            target_cols = ["КОДНАПР", "ПУТЬ", "ПД", "КМ", "ОЦЕНКА", "ПРИЧИНА"]
            
            for score, s_name in {5: "Отличные", 4: "Хорошие", 3: "Удовл", 2: "Неуд"}.items():
                subset = df_eval[df_eval["ОЦЕНКА"] == score]
                # Оставляем только существующие из списка целевых колонок
                available_cols = [c for c in target_cols if c in subset.columns]
                subset_to_save = subset[available_cols]
                subset_to_save.to_excel(writer, sheet_name=s_name, index=False)
            
            # Форматирование (Выравнивание по центру)
            from openpyxl.styles import Alignment
            for sheetname in writer.sheets:
                ws = writer.sheets[sheetname]
                for row in ws.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(horizontal='center', vertical='center')

        st.download_button(
            label="💾 Скачать итоговый отчет Excel",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- ТАБЛИЦЫ В ИНТЕРФЕЙСЕ ---
        tabs = st.tabs(["📊 Итоги Nуч", "✅ Отличные", "⭐ Хорошие", "⚠️ Удовл.", "🚨 Неуд"])
        
        with tabs[0]:
            st.dataframe(style_results(results_df), use_container_width=True, height=550)

        for i, score in enumerate([5, 4, 3, 2]):
            with tabs[i+1]:
                subset = df_eval[df_eval["ОЦЕНКА"] == score]
                available_cols = [c for c in target_cols if c in subset.columns]
                if not subset.empty:
                    st.dataframe(subset[available_cols], use_container_width=True)
                else:
                    st.info(f"Километров с оценкой {score} не обнаружено.")

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
