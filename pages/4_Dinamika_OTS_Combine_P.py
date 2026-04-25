import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. ПУЛЕНЕПРОБИВАЕМАЯ НОРМАЛИЗАЦИЯ ---
def normalize_dataframe(df):
    """
    Безопасно приводит заголовки к стандарту и исправляет типы данных,
    исключая ошибку 'DataFrame object has no attribute str'.
    """
    # Очистка заголовков: убираем пробелы, в верхний регистр, латиница -> кириллица
    def clean_header(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_header(col) for col in df.columns]

    # Карта синонимов (адаптация под разные выгрузки)
    mapping = {
        "КОДНАПРВ": "КОД", "КОДНАПР": "КОД", "KOD": "КОД",
        "ПУТЬ": "ПУТЬ", "PATH": "ПУТЬ",
        "КМ": "КМ", "KM": "КМ",
        "М": "М", "M": "М",
        "АМПЛИТУДА": "АМП", "AMP": "АМП",
        "ОТСТУПЛЕНИЕ": "ТИП", "OTST": "ТИП"
    }
    df = df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})

    # --- РЕШЕНИЕ ОШИБКИ .str ---
    # Мы вызываем .str ТОЛЬКО у конкретных колонок (Series), а не у всего df
    if "КОД" in df.columns:
        # astype(str) гарантирует, что даже числа станут текстом перед заменой
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["КОД"] = "0"

    # Безопасное преобразование числовых данных
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            # Сначала в строку -> замена запятой -> в число
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.title("📈 Динамика-О: Анализ роста")
st.write("Сравнение амплитуд отступлений между двумя проходами.")

c1, c2 = st.columns(2)
with c1:
    f1 = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_ots_final")
with c2:
    f2 = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_ots_final")

if f1 and f2:
    with st.spinner("Синхронизация данных..."):
        try:
            # Загрузка и нормализация
            df_old = normalize_dataframe(pd.read_excel(f1))
            df_new = normalize_dataframe(pd.read_excel(f2))

            # Проверка обязательных полей
            req = {"КМ", "М", "АМП", "ТИП"}
            if not req.issubset(df_new.columns):
                st.error(f"В файлах не найдены колонки: {req - set(df_new.columns)}")
                st.stop()

            # Синхронизация по метрам (округление до 2м для поиска совпадений)
            df_old['М_S'] = (df_old['М'] / 2).round() * 2
            df_new['М_S'] = (df_new['М'] / 2).round() * 2

            # Объединение (Merge)
            merged = pd.merge(
                df_new, 
                df_old[['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП', 'АМП']], 
                on=['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            # Вычисление роста
            merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                st.subheader("📊 Выявленные точки роста")
                
                # График
                fig = px.scatter(
                    df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП",
                    hover_data=['М', 'АМП_OLD', 'АМП']
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица
                display_cols = ['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']
                st.dataframe(df_result[display_cols], use_container_width=True)
                
                # Экспорт в Excel с выделением цвета
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result[display_cols].to_excel(writer, index=False, sheet_name="Рост")
                
                st.download_button("📥 Скачать результат (Excel)", output.getvalue(), "growth_report.xlsx")
            else:
                st.success("✅ Совпадающих отступлений с ростом амплитуды не найдено.")

        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
