import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill

# Настройка страницы
st.set_page_config(page_title="Анализ роста отступлений", layout="wide")

# --- 1. ФУНКЦИЯ НОРМАЛИЗАЦИИ (БЕЗОПАСНАЯ) ---
def normalize_dataframe(df):
    """
    Приводит заголовки к стандарту и исправляет типы данных.
    Исправляет ошибку 'DataFrame object has no attribute str'.
    """
    # Очистка заголовков (убираем пробелы, в верхний регистр)
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Карта переименования (синонимы для разных версий выгрузок)
    rename_map = {
        "КМ": "KM", "KM": "KM",
        "М": "M", "M": "M",
        "КОДНАПРВ": "KOD", "КОД": "KOD", "KOD": "KOD",
        "ПУТЬ": "PATH", "PATH": "PATH",
        "АМПЛИТУДА": "AMP", "AMP": "AMP",
        "ОТСТУПЛЕНИЕ": "OTST", "OTST": "OTST",
        "СТЕПЕНЬ": "STEP"
    }
    df = df.rename(columns=rename_map)

    # Исправляем колонку KOD (Код направления)
    # Сначала в строку (.astype(str)), потом убираем хвосты .0
    if "KOD" in df.columns:
        df["KOD"] = df["KOD"].astype(str).str.replace(".0", "", regex=False)
    else:
        df["KOD"] = "0"

    # Исправляем числовые колонки (замена запятой на точку)
    for col in ["KM", "M", "AMP"]:
        if col in df.columns:
            # .str вызывается только у конкретной колонки!
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
            
    return df

# --- 2. ОСНОВНОЙ ИНТЕРФЕЙС ---
st.title("📈 Динамика-О: Мониторинг роста амплитуд")
st.write("Загрузите результаты двух проходов вагона-дефектоскопа для сравнения.")

col1, col2 = st.columns(2)

with col1:
    file_old = st.file_uploader("📂 Прошлый проход (Excel)", type=["xlsx"], key="old_file")
with col2:
    file_new = st.file_uploader("📂 Текущий проход (Excel)", type=["xlsx"], key="new_file")

if file_old and file_new:
    try:
        with st.spinner("Синхронизация данных..."):
            # Загрузка
            df_old = normalize_dataframe(pd.read_excel(file_old))
            df_new = normalize_dataframe(pd.read_excel(file_new))

            # Проверка на наличие критических колонок
            required = {"KM", "M", "AMP", "OTST"}
            if not required.issubset(df_new.columns):
                st.error(f"В файлах отсутствуют необходимые колонки: {required - set(df_new.columns)}")
                st.stop()

            # Синхронизация по метрам (округление до 2м, чтобы поймать одно и то же отступление)
            df_old['M_SYNC'] = (df_old['M'] / 2).round() * 2
            df_new['M_SYNC'] = (df_new['M'] / 2).round() * 2

            # Объединение таблиц (Merge)
            # Ищем совпадения по коду, пути, км, сглаженным метрам и типу отступления
            merged = pd.merge(
                df_new, 
                df_old[['KOD', 'PATH', 'KM', 'M_SYNC', 'OTST', 'AMP']], 
                on=['KOD', 'PATH', 'KM', 'M_SYNC', 'OTST'], 
                how='inner', 
                suffixes=('', '_OLD')
            )

            # Расчет разницы (роста)
            merged['РОСТ'] = (merged['AMP'] - merged['AMP_OLD']).round(1)
            
            # Фильтруем только те, где амплитуда реально выросла
            df_result = merged[merged['РОСТ'] > 0].sort_values(by='РОСТ', ascending=False)

            if not df_result.empty:
                st.subheader(f"✅ Найдено точек роста: {len(df_result)}")
                
                # График распределения роста по километрам
                fig = px.scatter(
                    df_result, 
                    x="KM", 
                    y="РОСТ", 
                    size="AMP", 
                    color="OTST",
                    title="Карта роста амплитуд",
                    labels={"KM": "Километр", "РОСТ": "Прирост (мм)", "OTST": "Тип"},
                    hover_data=['M', 'AMP_OLD', 'AMP']
                )
                st.plotly_chart(fig, use_container_width=True)

                # Таблица результатов
                display_cols = ['KOD', 'PATH', 'KM', 'M', 'OTST', 'AMP_OLD', 'AMP', 'РОСТ']
                st.dataframe(df_result[display_cols], use_container_width=True)
                
                # Подготовка Excel для скачивания
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result[display_cols].to_excel(writer, index=False, sheet_name="Результат")
                    
                    # Цветовая индикация в Excel (опционально)
                    ws = writer.book["Результат"]
                    red_fill = PatternFill(start_color="FF9999", fill_type="solid")
                    for row in range(2, ws.max_row + 1):
                        # Если рост больше 5мм — подсвечиваем красным
                        if ws.cell(row=row, column=8).value > 5:
                            ws.cell(row=row, column=8).fill = red_fill

                st.download_button(
                    label="📥 Скачать отчет в Excel",
                    data=output.getvalue(),
                    file_name="Otchet_Rosta_OTS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("Совпадающих отступлений с ростом амплитуды не обнаружено.")

    except Exception as e:
        st.error(f"Критическая ошибка: {e}")
