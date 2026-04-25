import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. ФУНКЦИЯ ЗАЩИЩЕННОЙ НОРМАЛИЗАЦИИ ---
def normalize_dataframe(df):
    """Приводит заголовки к стандарту и защищает от отсутствующих данных."""
    def clean_header(text):
        if not isinstance(text, str): return text
        # Исправляем раскладку (латиница -> кириллица)
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_header(col) for col in df.columns]

    # Карта синонимов (добавил больше вариантов из практики РЖД)
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

    # --- ПРОВЕРКА И ФОРМАТИРОВАНИЕ ---
    # Если КОД не найден, создаем его как "0" (чтобы не было ошибки)
    if "КОД" in df.columns:
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["КОД"] = "0"
    
    # Если ПУТЬ не найден, ставим "1" по умолчанию
    if "ПУТЬ" not in df.columns:
        df["ПУТЬ"] = 1

    # Численные колонки
    for col in ["КМ", "М", "АМП"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ---
st.title("📈 Динамика-О: Анализ роста")
st.info("Загрузите файлы для сравнения. Если в файлах нет 'Кода направления', программа сопоставит данные по КМ и Пути.")

col_f1, col_f2 = st.columns(2)
with col_f1:
    file1 = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="up_old")
with col_f2:
    file2 = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="up_new")

if file1 and file2:
    with st.spinner("Синхронизация данных..."):
        # Читаем файлы
        try:
            df_old = normalize_dataframe(pd.read_excel(file1))
            df_new = normalize_dataframe(pd.read_excel(file2))
        except Exception as e:
            st.error(f"Ошибка при чтении Excel: {e}")
            st.stop()

    # Проверка критически важных колонок (без которых сравнение невозможно)
    required = {"КМ", "М", "АМП", "ТИП"}
    missing = required - set(df_new.columns)
    if missing:
        st.error(f"В текущем файле не найдены колонки: {', '.join(missing)}")
        st.stop()

    # Округление метров для сопоставления (допуск 2 метра)
    df_old['М_SYNC'] = (df_old['М'] / 2).round() * 2
    df_new['М_SYNC'] = (df_new['М'] / 2).round() * 2

    # Сопоставление (merge)
    merged = pd.merge(
        df_new, 
        df_old[['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП', 'АМП']], 
        on=['КОД', 'ПУТЬ', 'КМ', 'М_SYNC', 'ТИП'], 
        how='inner', 
        suffixes=('', '_OLD')
    )

    # Расчет роста
    merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
    df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

    if not df_result.empty:
        # Графики
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.histogram(df_result, x="РОСТ", title="Гистограмма роста (мм)", color_discrete_sequence=['#E03C31']), use_container_width=True)
        with c2:
            st.plotly_chart(px.scatter(df_result, x="КМ", y="РОСТ", size="АМП", color="ТИП", title="Локализация роста по КМ"), use_container_width=True)

        # Таблица результатов
        st.subheader("📋 Реестр выявленного роста")
        
        def highlight_val(val):
            if val >= 5: return 'background-color: #ff9999'
            if val >= 2: return 'background-color: #ffff99'
            return ''

        display_cols = ['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП_OLD', 'АМП', 'РОСТ']
        st.dataframe(df_result[display_cols].style.applymap(highlight_val, subset=['РОСТ']), use_container_width=True)
        
        # Экспорт
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result[display_cols].to_excel(writer, index=False)
        st.download_button("📥 Скачать Excel-отчет", output.getvalue(), "Rost_Ots.xlsx")
    else:
        st.success("✅ Рост амплитуд не выявлен.")
