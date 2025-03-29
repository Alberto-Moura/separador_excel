import streamlit as st

home_page = st.Page("home.py", title="Home", icon=":material/home:", default=True)
config_page = st.Page("config.py", title="Configure seu Excel", icon=":material/edit:")
relatorio_page = st.Page("separador.py", title="Gerar Relatórios", icon=":material/check_circle:")

pg = st.navigation([home_page, config_page, relatorio_page])
st.set_page_config(page_title="Separador de Planilha", page_icon=":material/edit:")
pg.run()