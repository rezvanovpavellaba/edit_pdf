import streamlit as st
import Page1, Page2

#st.set_page_config(
    #page_title="Моё приложение",
   # page_icon="📚",
   # layout="wide",
#)

# Боковая панель для переключения страниц
page = st.sidebar.selectbox(
    "Выберите страницу:",
    ["Получение исходных данных", "Редактирование PDF"]
)

# Логика переключения
if page == "Получение исходных данных":
    Page1.Page1()
else:
    Page2.Page2()
