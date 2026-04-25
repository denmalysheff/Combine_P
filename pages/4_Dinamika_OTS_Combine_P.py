import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Рост отступлений", layout="wide")

# ==============================
# НОРМАЛИЗАЦИЯ
# ==============================
def normalize_columns(df):
    new_cols = {}
    for col in df.columns:
        c = str(col).strip().upper()
        if c in ["КМ", "KM"]: new_cols[col] = "KM"
        elif c in ["М", "M"]: new_cols[col] = "M"
        elif c in ["КОДНАПРВ"]: new_cols[col] = "KOD"
        elif c in ["ПУТЬ"]: new_cols[col] = "PATH"
        elif c in ["АМПЛИТУДА"]: new_cols[col] = "AMP"
        elif c in ["ДЛИНА"]: new_cols[col] = "LEN"
        elif c in ["СТЕПЕНЬ"]: new_cols[col] = "STEP"
        elif c in ["ОТСТУПЛЕНИЕ"]: new_cols[col] = "OTST"
        elif c in ["БАЛЛ"]: new_cols[col] = "BALL"
        elif c in ["ИС"]: new_cols[col] = "IS"
        elif c in ["СТРЕЛКА"]: new_cols[col] = "STR"
        elif c in ["МОСТ"]: new_cols[col] = "MOST"
        elif c in ["ГОД"]: new_cols[col] = "YEAR"
        elif c in ["МЕСЯЦ"]: new_cols[col] = "MONTH"
        elif c in ["ДЕНЬ"]: new_cols[col] = "DAY"
        else: new_cols[col] = col
    return df.rename(columns=new_cols)

def to_numeric(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "."), errors="coerce")

def make_date(df):
    return pd.to_datetime(
        df["YEAR"].astype(str) + "-" + df["MONTH"].astype(str) + "-" + df["DAY"].astype(str),
        errors="coerce"
    )

# ==============================
# UI
# ==============================
st.title("📊 Мониторинг роста амплитуд")

c1, c2 = st.columns(2)
with c1:
    file1 = st.file_uploader("Загрузите файл предыдущей проверки", type=["xlsx"])
with c2:
    file2 = st.file_uploader("Загрузите файл текущей проверки", type=["xlsx"])

tolerance = st.number_input("Допуск по метрам (поиск совпадений)", value=3)

# ==============================
# ОБРАБОТКА
# ==============================
if st.button("🚀 Начать анализ", use_container_width=True):

    if not file1 or not file2:
        st.error("Загрузите оба файла для сравнения")
        st.stop()

    try:
        # Загрузка
        df1 = normalize_columns(pd.read_excel(file1, sheet_name="Отступления"))
        df2 = normalize_columns(pd.read_excel(file2, sheet_name="Отступления"))

        # Даты
        df1["DATE"] = make_date(df1)
        df2["DATE"] = make_date(df2)

        # Хронология
        if df1["DATE"].min() < df2["DATE"].min():
            old, new = df1, df2
        else:
            old, new = df2, df1

        # Очистка и типы
        for df in [old, new]:
            for col in ["KOD", "PATH", "OTST"]:
                if col in df.columns: df[col] = df[col].astype(str).str.strip()
            for col in ["KM", "M", "AMP"]:
                if col in df.columns: df[col] = to_numeric(df[col])

        old = old.dropna(subset=["KM", "M", "AMP"])
        new = new.dropna(subset=["KM", "M", "AMP"])

        # Merge
        merged = pd.merge(old, new, on=["KOD", "PATH", "KM", "OTST"], suffixes=("_old", "_new"))
        merged["delta_m"] = abs(merged["M_new"] - merged["M_old"])
        merged = merged[merged["delta_m"] <= tolerance]
        
        # Удаление дублей
        merged = merged.sort_values("delta_m").drop_duplicates(
            subset=["KOD", "PATH", "KM", "M_old", "OTST"], keep="first"
        )

        # Результат (только РОСТ)
        result = merged[merged["AMP_new"].abs() > merged["AMP_old"].abs()].copy()
        result["Рост"] = (result["AMP_new"].abs() - result["AMP_old"].abs()).round(1)

        if result.empty:
            st.warning("Совпадений с ростом амплитуды не найдено.")
            st.stop()

        # Формирование таблицы для отображения
        df_display = pd.DataFrame({
            "КМ": result["KM"],
            "М": result["M_new"],
            "Отступление": result["OTST"],
            "Было (мм)": result["AMP_old"],
            "Стало (мм)": result["AMP_new"],
            "Рост (мм)": result["Рост"],
            "Путь": result["PATH"],
            "Код": result["KOD"]
        })

        st.success(f"Выявлено точек роста: {len(df_display)}")
        st.dataframe(df_display.sort_values("Рост (мм)", ascending=False), use_container_width=True)

        # ==============================
        # НОВАЯ ЧИТАЕМАЯ ВИЗУАЛИЗАЦИЯ
        # ==============================
        st.subheader("📈 График динамики (по километрам)")
        
        fig = px.scatter(
            df_display, 
            x="КМ", 
            y="Рост (мм)", 
            color="Отступление",
            hover_data=["М", "Было (мм)", "Стало (мм)", "Путь"],
            title="Распределение прироста амплитуд",
            labels={"КМ": "Километр пути", "Рост (мм)": "Величина роста (мм)"},
            template="plotly_dark", # Темная тема часто лучше для инженерных графиков
            opacity=0.7             # Прозрачность для наслоений
        )

        # Настройка точек и сетки
        fig.update_traces(marker=dict(size=12, line=dict(width=1, color='White')))
        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='Gray')
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='Gray')
        
        # Добавляем горизонтальную линию "нулевого роста"
        fig.add_hline(y=0, line_dash="dash", line_color="white", opacity=0.5)

        st.plotly_chart(fig, use_container_width=True)

        # ==============================
        # ЭКСПОРТ EXCEL
        # ==============================
        def to_excel(df):
            from openpyxl.styles import PatternFill
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Рост")
                ws = writer.book["Рост"]
                red = PatternFill(start_color="FF9999", fill_type="solid")
                for row in range(2, ws.max_row + 1):
                    # Если рост в 6-й колонке > 5мм
                    if ws.cell(row=row, column=6).value > 5:
                        ws.cell(row=row, column=6).fill = red
            output.seek(0)
            return output

        st.download_button(
            label="📥 Скачать отчет (Excel)",
            data=to_excel(df_display),
            file_name="growth_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
