import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Анализ роста отступлений", layout="wide")

# --- 1. ФУНКЦИЯ БЕЗОПАСНОЙ НОРМАЛИЗАЦИИ ---
def normalize_dataframe(df):
    """
    Приводит заголовки к стандарту и исправляет типы данных.
    Исключает ошибку 'DataFrame object has no attribute str'.
    """
    # Очистка заголовков: убираем пробелы и переводим в верхний регистр
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Карта синонимов для унификации (КМ/KM, ПУТЬ/PATH и т.д.)
    rename_map = {
        "КМ": "KM", "KM": "KM",
        "М": "M", "M": "M",
        "КОДНАПРВ": "KOD", "КОД": "KOD", "KOD": "KOD",
        "ПУТЬ": "PATH", "PATH": "PATH",
        "АМПЛИТУДА": "AMP", "AMP": "AMP",
        "ОТСТУПЛЕНИЕ": "OTST", "OTST": "OTST"
    }
    df = df.rename(columns=rename_map)

    # Исправляем KOD (Код направления)
    if "KOD" in df.columns:
        # Критично: сначала .astype(str), потом .str.replace
        df["KOD"] = df["KOD"].astype(str).str.replace(".0", "", regex=False)
    
    # Безопасное преобразование числовых колонок
    for col in ["KM", "M", "AMP"]:
        if col in df.columns:
            # Превращаем в строку -> меняем запятую на точку -> в число
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.title("📈 Динамика-О: Анализ роста")
st.write("Сравнение амплитуд отступлений между двумя проходами.")

c1, c2 = st.columns(2)
with c1:
    f_old = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_ots")
with c2:
    f_new = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_ots")

if f_old and f_new:
    try:
        with st.spinner("Синхронизация данных..."):
            # Загрузка и нормализация
            df_old = normalize_dataframe(pd.read_excel(f_old))
            df_new = normalize_dataframe(pd.read_excel(f_new))

            # Проверка обязательных полей
            req = {"KM", "M", "AMP", "OTST"}
            if not req.issubset(df_new.columns):
                st.error(f"В файлах не найдены колонки: {req - set(df_new.columns)}")
                st.stop()

            # Синхронизация по метрам (округление до 2м для поиска совпадений)
            df_old['M_S'] = (df_old['M'] / 2).round() * 2
            df_new['M_S'] = (df_new['M'] / 2).round() * 2

            # Объединение (Merge)
            merged = pd.merge(
                df_new, 
                df_old[['KOD', 'PATH', 'KM', 'M_S', 'OTST', 'AMP']], 
                on=['KOD', 'PATH', 'KM', 'M_S', 'OTST'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            # Вычисление роста
            merged['РОСТ'] = (merged['AMP'] - merged['AMP_OLD']).round(1)
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                st.subheader(f"📊 Выявлено точек роста: {len(df_result)}")
                
                # Визуализация
                fig = px.scatter(
                    df_result, x="KM", y="РОСТ", size="AMP", color="OTST",
                    hover_data=['M', 'AMP_OLD', 'AMP'],
                    title="Распределение роста амплитуд по километрам"
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица
                disp_cols = ['KOD', 'PATH', 'KM', 'M', 'OTST', 'AMP_OLD', 'AMP', 'РОСТ']
                st.dataframe(df_result[disp_cols], use_container_width=True)
                
                # Экспорт в Excel с подсветкой
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result[disp_cols].to_excel(writer, index=False, sheet_name="Рост")
                    ws = writer.book["Рост"]
                    red = PatternFill(start_color="FF9999", fill_type="solid")
                    for row in range(2, ws.max_row + 1):
                        # Подсвечиваем ячейку "РОСТ" (8-я колонка), если рост > 5мм
                        if ws.cell(row=row, column=8).value > 5:
                            ws.cell(row=row, column=8).fill = red
                
                st.download_button("📥 Скачать результат (Excel)", output.getvalue(), "growth_report.xlsx")
            else:
                st.success("✅ Совпадающих отступлений с ростом амплитуды не найдено.")

    except Exception as e:
        st.error(f"❌ Ошибка при обработке: {e}")
