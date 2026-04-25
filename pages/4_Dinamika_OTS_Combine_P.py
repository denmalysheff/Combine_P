import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. ФУНКЦИЯ ЗАЩИЩЕННОЙ НОРМАЛИЗАЦИИ ---
def normalize_dataframe(df):
    """Приводит заголовки к стандарту и исправляет типы данных."""
    def clean_header(text):
        if not isinstance(text, str): return text
        # Исправляем раскладку (латиница -> кириллица для РЖД стандартов)
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    # Очищаем заголовки
    df.columns = [clean_header(col) for col in df.columns]

    # Карта синонимов для колонок
    column_map = {
        "КОДНАПРВ": "КОД", "КОДНАПР": "КОД", "KOD": "КОД", "НАПРАВЛЕНИЕ": "КОД",
        "ПУТЬ": "ПУТЬ", "PATH": "ПУТЬ",
        "КМ": "КМ", "KM": "КМ",
        "М": "М", "M": "М",
        "АМПЛИТУДА": "АМП", "AMP": "АМП",
        "СТЕПЕНЬ": "СТЕПЕНЬ", "STEP": "СТЕПЕНЬ",
        "ОТСТУПЛЕНИЕ": "ТИП", "OTST": "ТИП", "НЕИСПРАВНОСТЬ": "ТИП"
    }
    df = df.rename(columns={k: v for k, v in column_map.items() if k in df.columns})

    # --- ИСПРАВЛЕНИЕ ОШИБКИ 'DataFrame object has no attribute str' ---
    # Обрабатываем колонку "КОД" только если она есть
    if "КОД" in df.columns:
        # Важно: вызываем .astype(str).str у КОЛОНКИ, а не у df
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["КОД"] = "0"
    
    if "ПУТЬ" not in df.columns:
        df["ПУТЬ"] = 1

    # Безопасное приведение чисел для колонок КМ, М, АМП
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            # Сначала в строку, потом замена запятой, потом в число
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ---
st.title("📈 Динамика-О: Анализ роста")
st.write("Сравнение амплитуд отступлений между двумя проходами.")

col_f1, col_f2 = st.columns(2)
with col_f1:
    file1 = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_file_ots")
with col_f2:
    file2 = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_file_ots")

if file1 and file2:
    with st.spinner("Загрузка и обработка..."):
        try:
            # Читаем данные
            df_old = normalize_dataframe(pd.read_excel(file1))
            df_new = normalize_dataframe(pd.read_excel(file2))

            # Проверка наличия обязательных данных
            required = {"КМ", "М", "АМП", "ТИП"}
            if not required.issubset(df_new.columns):
                st.error(f"В файле отсутствуют колонки: {required - set(df_new.columns)}")
                st.stop()

            # Синхронизация по метрам (допуск 2 метра)
            df_old['М_S'] = (df_old['М'] / 2).round() * 2
            df_new['М_S'] = (df_new['М'] / 2).round() * 2

            # Склеиваем данные
            merged = pd.merge(
                df_new, 
                df_old[['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП', 'АМП']], 
                on=['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                # Визуализация
                st.subheader("📊 Анализ динамики")
                fig = px.scatter(
                    df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП",
                    hover_data=['М', 'АМП_OLD', 'АМП']
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица
                st.subheader("📋 Список выросших отступлений")
                st.dataframe(df_result[['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']], use_container_width=True)
                
                # Экспорт
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Рост")
                st.download_button("📥 Скачать результат", output.getvalue(), "dynamic_report.xlsx")
            else:
                st.success("Рост амплитуд не выявлен.")

        except Exception as e:
            st.error(f"Произошла ошибка при обработке: {e}")
