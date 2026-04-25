import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. ПУЛЕНЕПРОБИВАЕМАЯ НОРМАЛИЗАЦИЯ ---
def normalize_dataframe(df):
    """
    Безопасно приводит заголовки к стандарту и исправляет типы данных.
    """
    def clean_header(text):
        if not isinstance(text, str): return text
        # Замена латиницы на кириллицу (KM -> КМ)
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    # 1. Очищаем заголовки
    df.columns = [clean_header(col) for col in df.columns]

    # 2. Словарь синонимов
    column_map = {
        "КОДНАПРВ": "КОД", "КОДНАПР": "КОД", "KOD": "КОД",
        "ПУТЬ": "ПУТЬ", "PATH": "ПУТЬ",
        "КМ": "КМ", "KM": "КМ",
        "М": "М", "M": "М",
        "АМПЛИТУДА": "АМП", "AMP": "АМП",
        "ОТСТУПЛЕНИЕ": "ТИП", "OTST": "ТИП"
    }
    df = df.rename(columns={k: v for k, v in column_map.items() if k in df.columns})

    # 3. Безопасная обработка столбцов (исправляет ошибку 'DataFrame' object has no attribute 'str')
    if "КОД" in df.columns:
        # Применяем .str только к конкретной колонке!
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["КОД"] = "0"  # Заглушка, если колонки нет

    if "ПУТЬ" not in df.columns:
        df["ПУТЬ"] = 1

    # Обработка числовых данных
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            # Преобразуем серию в строку -> меняем запятую -> в число
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.title("📈 Динамика-О: Анализ роста")
st.write("Сравнение амплитуд отступлений между двумя проходами.")

col_f1, col_f2 = st.columns(2)
with col_f1:
    file1 = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_file_safe")
with col_f2:
    file2 = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_file_safe")

if file1 and file2:
    with st.spinner("Синхронизация данных..."):
        try:
            # Загрузка
            df_old = normalize_dataframe(pd.read_excel(file1))
            df_new = normalize_dataframe(pd.read_excel(file2))

            # Проверка обязательных полей
            required = {"КМ", "М", "АМП", "ТИП"}
            if not required.issubset(df_new.columns):
                st.error(f"В файлах не найдены колонки: {required - set(df_new.columns)}")
                st.stop()

            # Округление метров (синхронизация с допуском 2м)
            df_old['М_SYNC'] = (df_old['М'] / 2).round() * 2
            df_new['М_SYNC'] = (df_new['М'] / 2).round() * 2

            # Объединение (Merge)
            merged = pd.merge(
                df_new, 
                df_old[['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП', 'АМП']], 
                on=['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            # Вычисление роста
            merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                st.subheader("📊 Результаты анализа")
                
                # График
                fig = px.scatter(
                    df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП",
                    hover_data=['М', 'АМП_OLD', 'АМП'],
                    title="Точки роста амплитуд"
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица
                display_cols = ['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']
                st.dataframe(df_result[display_cols], use_container_width=True)
                
                # Кнопка скачивания
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result[display_cols].to_excel(writer, index=False)
                st.download_button("📥 Скачать отчет (Excel)", output.getvalue(), "dynamic_report.xlsx")
            else:
                st.success("✅ Роста амплитуд не обнаружено.")

        except Exception as e:
            st.error(f"❌ Критическая ошибка: {e}")
