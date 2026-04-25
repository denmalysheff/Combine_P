import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Рост отступлений", layout="wide")

# ==============================
# БЕЗОПАСНАЯ НОРМАЛИЗАЦИЯ
# ==============================
def normalize_columns(df):
    """
    Исправляет заголовки и типы данных. 
    Критически важно: .str вызывается только у Series после .astype(str).
    """
    # 1. Чистим заголовки: убираем пробелы и в верхний регистр
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # 2. Карта переименования для унификации
    rename_map = {
        "КМ": "KM", "KM": "KM",
        "М": "M", "M": "M",
        "КОДНАПРВ": "KOD", "КОД": "KOD", "KOD": "KOD",
        "ПУТЬ": "PATH", "PATH": "PATH",
        "АМПЛИТУДА": "AMP", "AMP": "AMP",
        "СТЕПЕНЬ": "STEP", "STEP": "STEP",
        "ОТСТУПЛЕНИЕ": "OTST", "OTST": "OTST"
    }
    df = df.rename(columns=rename_map)

    # 3. БЕЗОПАСНАЯ ОБРАБОТКА СТОЛБЦОВ (Решение вашей ошибки)
    # Обрабатываем 'KOD', если он есть
    if "KOD" in df.columns:
        # Сначала в строку, потом замена. Это предотвращает ошибку 'AttributeError'
        df["KOD"] = df["KOD"].astype(str).str.replace(".0", "", regex=False)
    
    # Обработка числовых колонок (замена запятой на точку для корректного float)
    for col in ["KM", "M", "AMP"]:
        if col in df.columns:
            # Применяем .str только к конкретной колонке
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# ==============================
# ИНТЕРФЕЙС И ЛОГИКА
# ==============================
st.title("📈 Анализ роста амплитуд (Модуль 4)")

col1, col2 = st.columns(2)
with col1:
    f_old = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="u_old")
with col2:
    f_new = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="u_new")

if f_old and f_new:
    try:
        # Читаем данные (по умолчанию первый лист)
        df_old = normalize_dataframe(pd.read_excel(f_old))
        df_new = normalize_dataframe(pd.read_excel(f_new))

        # Проверка обязательных полей
        required = {"KM", "M", "AMP", "OTST"}
        if not required.issubset(df_new.columns):
            st.error(f"В файле не найдены колонки: {required - set(df_new.columns)}")
            st.stop()

        # Округление метров для синхронизации (допуск 2 метра)
        df_old['M_S'] = (df_old['M'] / 2).round() * 2
        df_new['M_S'] = (df_new['M'] / 2).round() * 2

        # Склеиваем данные (Merge)
        merged = pd.merge(
            df_new, 
            df_old[['KOD', 'PATH', 'KM', 'M_S', 'OTST', 'AMP']], 
            on=['KOD', 'PATH', 'KM', 'M_S', 'OTST'], 
            how='inner', 
            suffixes=('', '_OLD')
        )

        # Вычисляем рост
        merged['РОСТ'] = (merged['AMP'] - merged['AMP_OLD']).round(1)
        df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

        if not df_result.empty:
            st.subheader("📊 Результаты сравнения")
            
            # График
            fig = px.scatter(df_result, x="KM", y="РОСТ", size="AMP", color="OTST", hover_data=['M'])
            st.plotly_chart(fig, use_container_width=True)

            # Таблица
            st.dataframe(df_result, use_container_width=True)
            
            # Экспорт
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_result.to_excel(writer, index=False, sheet_name="Результат")
            st.download_button("📥 Скачать Excel", output.getvalue(), "growth_report.xlsx")
        else:
            st.success("✅ Роста амплитуд не выявлено.")

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
