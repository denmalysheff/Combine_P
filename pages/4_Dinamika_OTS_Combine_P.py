import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Мониторинг роста амплитуд", layout="wide")

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
        elif c in ["СТРЕЛКА", "СТР"]: new_cols[col] = "STR"
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
    file1 = st.file_uploader("Файл предыдущей проверки", type=["xlsx"])
with c2:
    file2 = st.file_uploader("Файл текущей проверки", type=["xlsx"])

tolerance = st.number_input("Допуск (м) для поиска совпадений", value=3)

if st.button("🚀 Начать анализ", use_container_width=True):
    if not file1 or not file2:
        st.error("Загрузите оба файла")
        st.stop()

    try:
        df1 = normalize_columns(pd.read_excel(file1, sheet_name="Отступления"))
        df2 = normalize_columns(pd.read_excel(file2, sheet_name="Отступления"))

        df1["DATE"] = make_date(df1)
        df2["DATE"] = make_date(df2)

        # Определение хронологии
        if df1["DATE"].min() < df2["DATE"].min():
            old, new = df1, df2
        else:
            old, new = df2, df1

        # Очистка
        for df in [old, new]:
            for col in ["KOD", "PATH", "OTST", "IS", "STR", "MOST"]:
                if col in df.columns: df[col] = df[col].astype(str).replace("nan", "").str.strip()
            for col in ["KM", "M", "AMP", "LEN", "BALL"]:
                if col in df.columns: df[col] = to_numeric(df[col])

        old = old.dropna(subset=["KM", "M", "AMP"])
        new = new.dropna(subset=["KM", "M", "AMP"])

        # Сопоставление
        merged = pd.merge(old, new, on=["KOD", "PATH", "KM", "OTST"], suffixes=("_old", "_new"))
        merged["delta_m"] = abs(merged["M_new"] - merged["M_old"])
        merged = merged[merged["delta_m"] <= tolerance]
        
        merged = merged.sort_values("delta_m").drop_duplicates(
            subset=["KOD", "PATH", "KM", "M_old", "OTST"], keep="first"
        )

        # Только рост по абсолютному значению
        result = merged[merged["AMP_new"].abs() > merged["AMP_old"].abs()].copy()
        result["Рост"] = (result["AMP_new"].abs() - result["AMP_old"].abs()).round(1)

        def get_c(col_name):
            return result[col_name] if col_name in result.columns else ""

        # Формирование таблицы в заданном порядке колонок
        df_display = pd.DataFrame({
            "Код": get_c("KOD"),
            "КМ": get_c("KM"),
            "М": get_c("M_new"),
            "Путь": get_c("PATH"),
            "Отступление": get_c("OTST"),
            "Дата (пред)": result["DATE_old"].dt.strftime('%d.%m.%Y'),
            "Амп (пред)": get_c("AMP_old"),
            "Длина (пред)": get_c("LEN_old"),
            "Балл (пред)": get_c("BALL_old"),
            "Степень (пред)": get_c("STEP_old"),
            "Дата (тек)": result["DATE_new"].dt.strftime('%d.%m.%Y'),
            "Амп (тек)": get_c("AMP_new"),
            "Длина (тек)": get_c("LEN_new"),
            "Балл (тек)": get_c("BALL_new"),
            "Степень (тек)": get_c("STEP_new"),
            "ИС": get_c("IS_old"),
            "СТР": get_c("STR_old"),
            "МОСТ": get_c("MOST_old"),
            "Рост (мм)": result["Рост"] # Рост - последняя колонка
        }).sort_values("Рост (мм)", ascending=False) # Фильтрация/сортировка по росту

        # Вывод в интерфейс
        critical = df_display[df_display["Рост (мм)"] >= 10]
        if not critical.empty:
            st.error("⚠️ КРИТИЧЕСКИЙ РОСТ (ОБРАТИ ВНИМАНИЕ)")
            st.table(critical)
        
        st.subheader("📋 Список выявленного роста")
        st.dataframe(df_display, use_container_width=True)

        # Функция экспорта с форматированием
        def to_excel(df):
            from openpyxl.styles import PatternFill, Alignment, Border, Side
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Анализ_Роста")
                ws = writer.book["Анализ_Роста"]
                
                red_fill = PatternFill(start_color="FF9999", fill_type="solid")
                center_align = Alignment(horizontal="center", vertical="center")
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                
                growth_col_idx = df.columns.get_loc("Рост (мм)") + 1
                
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                    for cell in row:
                        # 1. Выравнивание по центру всей таблицы
                        cell.alignment = center_align
                        # 2. Сетка
                        cell.border = thin_border
                        
                        # 3. Подсветка роста в последней колонке
                        if cell.row > 1 and cell.column == growth_col_idx:
                            if cell.value and cell.value >= 5:
                                cell.fill = red_fill
                
                # Автоподбор ширины
                for col in ws.columns:
                    max_len = max(len(str(cell.value)) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 3
                    
            output.seek(0)
            return output

        filename = f"Мониторинг_роста_амплитуд_{datetime.now().strftime('%d.%m.%Y')}.xlsx"
        st.download_button("📥 Скачать Excel с форматированием", 
                           data=to_excel(df_display), 
                           file_name=filename,
                           use_container_width=True)

    except Exception as e:
        st.error(f"Ошибка: {e}")
