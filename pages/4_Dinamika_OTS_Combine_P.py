import streamlit as st
import pandas as pd
import io
import plotly.express as px

# --- 1. ФУНКЦИЯ БЕЗОПАСНОЙ ОБРАБОТКИ ДАННЫХ ---
def safe_normalize(df):
    """Исправляет заголовки и типы данных без риска падения по .str"""
    
    # Очистка заголовков: убираем пробелы, в верхний регистр, латиница -> кириллица
    def clean_header(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_header(col) for col in df.columns]

    # Карта синонимов (адаптация под разные форматы выгрузок РЖД)
    mapping = {
        "КОДНАПРВ": "КОД", "КОДНАПР": "КОД", "KOD": "КОД",
        "ПУТЬ": "ПУТЬ", "PATH": "ПУТЬ",
        "КМ": "КМ", "KM": "КМ",
        "М": "М", "M": "М",
        "АМПЛИТУДА": "АМП", "AMP": "АМП",
        "ОТСТУПЛЕНИЕ": "ТИП", "OTST": "ТИП"
    }
    df = df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})

    # --- ИСПРАВЛЕНИЕ ОШИБКИ .str ---
    # Мы обрабатываем только конкретные колонки (Series), а не весь DataFrame
    
    if "КОД" in df.columns:
        # astype(str) гарантирует, что мы работаем с текстом, а не с числами
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["КОД"] = "0"

    # Безопасная конвертация чисел (замена запятой на точку)
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            # Сначала в строку -> замена -> в число (ошибки станут NaN)
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ОСНОВНОЙ ИНТЕРФЕЙС STREAMLIT ---
st.title("📈 Динамика-О: Анализ роста")
st.write("Сравнение амплитуд отступлений между двумя проходами.")

c1, c2 = st.columns(2)
with c1:
    f_old = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_ots")
with c2:
    f_new = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_ots")

if f_old and f_new:
    with st.spinner("Анализируем данные..."):
        try:
            # Загрузка и нормализация
            df_old = safe_normalize(pd.read_excel(f_old))
            df_new = safe_normalize(pd.read_excel(f_new))

            # Проверка критических колонок
            req = {"КМ", "М", "АМП", "ТИП"}
            if not req.issubset(df_new.columns):
                st.error(f"В файлах не найдены нужные колонки: {req - set(df_new.columns)}")
                st.stop()

            # Синхронизация по метрам (допуск 2 метра)
            df_old['М_S'] = (df_old['М'] / 2).round() * 2
            df_new['М_S'] = (df_new['М'] / 2).round() * 2

            # Объединение данных
            merged = pd.merge(
                df_new, 
                df_old[['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП', 'АМП']], 
                on=['КОД', 'ПУТЬ', 'КМ', 'М_S', 'ТИП'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            # Расчет разницы
            merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
            # Оставляем только те, где амплитуда реально выросла
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                st.subheader("📊 Выявленные точки роста")
                
                # Визуализация
                fig = px.scatter(
                    df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП",
                    hover_data=['М', 'АМП_OLD', 'АМП']
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица
                cols_to_show = ['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']
                st.dataframe(df_result[cols_to_show], use_container_width=True)
                
                # Кнопка экспорта
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result[cols_to_show].to_excel(writer, index=False)
                st.download_button("📥 Скачать Excel", output.getvalue(), "growth_report.xlsx")
            else:
                st.success("✅ Совпадающих отступлений с ростом амплитуды не найдено.")

        except Exception as e:
            st.error(f"Критическая ошибка при обработке: {e}")
