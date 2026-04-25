import streamlit as st

# Настройка страницы
st.set_page_config(page_title="Аналитический Хаб", layout="wide", initial_sidebar_state="collapsed")

# Кастомный CSS для красоты карточек (опционально)
st.markdown("""
    <style>
    .module-card {
        border: 1px solid #ddd;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        background-color: #f9f9f9;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎛 Панель управления модулями")
st.write("Выберите необходимый инструмент для начала работы")

# Создаем сетку 2x3 (для 5 модулей)
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)
col5, _ = st.columns(2)

# --- МОДУЛЬ 1 ---
with col1:
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.subheader("📊 Калькулятор Nuch")
    st.write("Расчет балловой оценки состояния пути.")
    if st.button("Открыть Модуль Nuch", key="btn1", use_container_width=True):
        st.switch_page("pages/1_Km.py")
    st.markdown('</div>', unsafe_allow_html=True)

# --- МОДУЛЬ 2 ---
with col2:
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.subheader("📐 Калькулятор Nuch")
    st.write("Расчет показателей качества устройства пути по формулам.")
    if st.button("Открыть Расчеты", key="btn2", use_container_width=True):
        st.switch_page("pages/2_Nuch.py")
    st.markdown('</div>', unsafe_allow_html=True)

# --- МОДУЛЬ 3 ---
with col3:
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.subheader("📦 Складской учет")
    st.write("Контроль наличия материалов и инвентаря на дистанции.")
    if st.button("Открыть Склад", key="btn3", use_container_width=True):
        # Замени на реальное имя файла в папке pages
        st.switch_page("pages/3_Sklad.py") 
    st.markdown('</div>', unsafe_allow_html=True)

# --- МОДУЛЬ 4 ---
with col4:
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.subheader("📅 Dinamika_OTS_Combine_P")
    st.write("Анализ роста амплитуд отступлений от предыдущей проверки к текущей.")
    if st.button("Открыть модуль Dinamika_OTS", key="btn4", use_container_width=True):
        st.switch_page("pages/4_Dinamika_OTS_Combine_P.py")
    st.markdown('</div>', unsafe_allow_html=True)

# --- МОДУЛЬ 5 ---
with col5:
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.subheader("👥 Кадры")
    st.write("Отчеты по обучению и квалификации сотрудников.")
    if st.button("Открыть Кадры", key="btn5", use_container_width=True):
        st.switch_page("pages/5_Staff.py")
    st.markdown('</div>', unsafe_allow_html=True)
