import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# --- 1. УНИВЕРСАЛЬНАЯ НОРМАЛИЗАЦИЯ ---
def normalize_dataframe(df):
    """Приводит заголовки и типы данных к единому стандарту."""
    def clean_header(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_header(col) for col in df.columns]

    # Маппинг синонимов для РЖД-выгрузок
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

    # Приведение типов
    if "КОД" in df.columns:
        df["КОД"] = df["КОД"].astype(str).str.replace(".0", "", regex=False)
    
    for col in ["КМ", "М", "АМП", "БАЛЛ"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС МОДУЛЯ ---
st.title("📈 Динамика-О: Рост отступлений")
st.info("Сравнение двух проходов для выявления быстрорастущих неисправностей.")

col_f1, col_f2 = st.columns(2)
with col_f1:
    file1 = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"])
with col_f2:
    file2 = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"])

if file1 and file2:
    # Загрузка
    with st.spinner("Нормализация данных..."):
        df_old = normalize_dataframe(pd.read_excel(file1))
        df_new = normalize_dataframe(pd.read_excel(file2))

    # Логика сопоставления
    # Склеиваем по Коду, Пути, КМ и Типу отступления. Допуск по метрам (М) реализуем через округление или merge_asof
    # Для простоты здесь - точное совпадение КМ, далее можно расширить
    merged = pd.merge(
        df_new, 
        df_old[['КОД', 'ПУТЬ', 'КМ', 'М', 'ТИП', 'АМП']], 
        on=['КОД', 'ПУТЬ', 'КМ', 'ТИП'], 
        how='inner', 
        suffixes=('', '_OLD')
    )

    # Расчет роста
    merged['РОСТ'] = (merged['АМП'] - merged['АМП_OLD']).round(1)
    df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

    if not df_result.empty:
        # --- 3. ВИЗУАЛИЗАЦИЯ ---
        st.subheader("📊 Анализ интенсивности роста")
        fig = px.histogram(
            df_result, x="РОСТ", 
            title="Распределение отступлений по величине роста",
            labels={'РОСТ': 'Рост амплитуды (мм)', 'count': 'Кол-во'},
            color_discrete_sequence=['#ff4b4b']
        )
        st.plotly_chart(fig, use_container_width=True)

        # --- 4. ТАБЛИЦА РЕЗУЛЬТАТОВ ---
        st.subheader("📋 Выявленный рост (ТОП-20)")
        
        def color_growth(val):
            color = 'white'
            if val >= 5: color = '#ff9999' # Красный
            elif val >= 2: color = '#ffff99' # Желтый
            return f'background-color: {color}'

        st.dataframe(
            df_result.head(20).style.applymap(color_growth, subset=['РОСТ']),
            use_container_width=True
        )

        # --- 5. ЭКСПОРТ ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="Рост_отступлений")
            
            # Авто-ширина и раскраска в Excel
            ws = writer.book["Рост_отступлений"]
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=df_result.columns.get_loc("РОСТ") + 1)
                if cell.value >= 5:
                    cell.fill = PatternFill(start_color="FF9999", fill_type="solid")
                elif cell.value >= 2:
                    cell.fill = PatternFill(start_color="FFFF99", fill_type="solid")

        st.download_button(
            label="📥 Скачать полный отчет в Excel",
            data=output.getvalue(),
            file_name="Dinamika_OTS_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.balloons()
        st.success("Рост отступлений не выявлен! Все амплитуды стабильны или уменьшились.")
