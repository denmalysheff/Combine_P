import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. ПУЛЕНЕПРОБИВАЕМАЯ НОРМАЛИЗАЦИЯ ---
def normalize_dataframe(df):
    """Приводит заголовки к стандарту и защищает от отсутствующих колонок."""
    def clean_header(text):
        if not isinstance(text, str): return text
        # Исправляем раскладку букв (латиница -> кириллица)
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_header(col) for col in df.columns]

    # Карта синонимов
    column_map = {
        "КОДНАПРВ": "КОД", "КОДНАПР": "КОД", "KOD": "КОД",
        "ПУТЬ": "ПУТЬ", "PATH": "ПУТЬ",
        "КМ": "КМ", "KM": "КМ",
        "М": "М", "M": "М",
        "АМПЛИТУДА": "АМП", "AMP": "АМП",
        "СТЕПЕНЬ": "СТЕПЕНЬ", "STEP": "СТЕПЕНЬ",
        "ОТСТУПЛЕНИЕ": "ТИП", "OTST": "ТИП"
    }
    df = df.rename(columns={k: v for k, v in column_map.items() if k in df.columns})

    # --- ЗАЩИТА: проверяем наличие колонки перед обработкой ---
    if "КОД" in df.columns:
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        # Если кода нет, создаем заглушку, чтобы мердж не упал
        df["КОД"] = "0"
    
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ---
st.title("📈 Динамика-О: Анализ роста")

col_f1, col_f2 = st.columns(2)
with col_f1:
    file1 = st.file_uploader("📂 Прошлый проход", type=["xlsx"], key="f1")
with col_f2:
    file2 = st.file_uploader("📂 Текущий проход", type=["xlsx"], key="f2")

if file1 and file2:
    with st.spinner("Анализирую данные..."):
        # Читаем данные с защитой от пустых листов
        df_old = normalize_dataframe(pd.read_excel(file1))
        df_new = normalize_dataframe(pd.read_excel(file2))

    # Проверка: есть ли минимальный набор данных для сравнения?
    required = {"КМ", "М", "АМП", "ТИП"}
    if not required.issubset(df_new.columns) or not required.issubset(df_old.columns):
        st.error(f"В файлах не найдены обязательные колонки: {required - set(df_new.columns)}")
        st.stop()

    # Синхронизация метров (допуск 2 метра)
    df_old['М_SYNC'] = (df_old['М'] / 2).round() * 2
    df_new['М_SYNC'] = (df_new['М'] / 2).round() * 2

    # Сопоставление
    merged = pd.merge(
        df_new, 
        df_old[['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП', 'АМП']], 
        on=['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП'], 
        how='inner', 
        suffixes=('', '_OLD')
    )

    merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
    df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

    if not df_result.empty:
        # График
        fig = px.scatter(df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП", title="Карта роста амплитуд")
        st.plotly_chart(fig, use_container_width=True)

        # Таблица
        st.subheader("📋 Выявленный рост")
        st.dataframe(df_result[['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']], use_container_width=True)
        
        # Кнопка скачивания
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False)
        st.download_button("📥 Скачать Excel", output.getvalue(), "rost_ots.xlsx")
    else:
        st.success("Роста не найдено.")
