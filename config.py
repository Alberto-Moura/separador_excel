import streamlit as st
import json
import pandas as pd

CONFIG_FILE = "config_excel.json"

def salvar_configuracoes(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

def carregar_configuracoes():
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

st.title("Configuração de Estilo para Excel")

config_atual = carregar_configuracoes()

config, exemplo = st.columns(2)

with config:
    # Fonte
    st.subheader("Fontes", divider='grey')
    tamanho_fonte_cabecalho = st.slider("Tamanho da Fonte do Cabeçalho", 8, 40, config_atual.get("tamanho_fonte_cabecalho", 16))
    tamanho_fonte_tabela = st.slider("Tamanho da Fonte da Tabela", 8, 40, config_atual.get("tamanho_fonte_tabela", 10))

    # Altura das Linhas
    st.subheader("Altura das Linhas", divider='gray')
    altura_linhas_cabecalho = st.slider("Altura da Linha do Cabeçalho", 10, 50, config_atual.get("altura_linhas_cabecalho", 30))
    altura_linhas_tabela = st.slider("Altura da Linha da Tabela", 10, 50, config_atual.get("altura_linhas_tabela", 16))

    # Alinhamento Vertical
    st.subheader("Alinhamento Cabeçalho", divider='grey')
    col1, col2 = st.columns(2)
    with col1:
        alinhamento_horizontal_cabecalho = st.selectbox("Alinhamento Horizontal", ["left", "center", "right"], index=["left", "center", "right"].index(config_atual.get("alinhamento_horizontal_cabecalho", "center")))
    with col2:
        alinhamento_vertical_cabecalho = st.selectbox("Alinhamento Vertical", ["top", "middle", "bottom"], index=["top", "middle", "bottom"].index(config_atual.get("alinhamento_vertical_cabecalho", "middle")))

    # Alinhamento Horizontal
    st.subheader("Alinhamento Texto", divider='grey')
    col1, col2 = st.columns(2)
    with col1:
        alinhamento_horizontal_texto = st.selectbox("Alinhamento Horizontal ", ["left", "center", "right"], index=["left", "center", "right"].index(config_atual.get("alinhamento_horizontal_texto", "center")))
    with col2:
        alinhamento_vertical_texto = st.selectbox("Alinhamento Vertical ", ["top", "middle", "bottom"], index=["top", "middle", "bottom"].index(config_atual.get("alinhamento_vertical_texto", "middle")))

with exemplo:
    # Cores
    st.subheader("Cores", divider='grey')
    col1, col2 = st.columns(2)
    with col1:
        cor_cabecalho = st.color_picker("Cor Cabeçalho", config_atual.get("cor_cabecalho", "#FFDD57"))
    with col2:
        cor_fonte_cabecalho = st.color_picker("Cor Fonte Cabeçalho", config_atual.get("cor_fonte_cabecalho", "#FFFFFF"))

    col3, col4 = st.columns(2)
    with col3:
        cor_base_tabela = st.color_picker("Cor Base Tabela", config_atual.get("cor_fundo_tabela", "#000000"))
    with col4:
        cor_texto_tabela = st.color_picker("Cor do Texto Tabela", config_atual.get("cor_texto_tabela", "#000000"))

    st.subheader("Exemplo de Tabela", divider='grey')
    # Exemplo de dados
    dados = [
        ["Nome"],
        ["Alice"],
        ["Bob"],
        ["Carol"]
    ]

    # Estilização via HTML e CSS
    tabela_html = f"""
    <style>
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: {tamanho_fonte_tabela}px;
            text-align: {alinhamento_horizontal_cabecalho.lower()};
        }}
        th {{
            background-color: {cor_cabecalho};
            color: {cor_fonte_cabecalho};
            font-size: {tamanho_fonte_cabecalho}px;
            height: {altura_linhas_cabecalho}px;
            text-align: {alinhamento_horizontal_cabecalho.lower()};
            vertical-align: {alinhamento_vertical_cabecalho.lower()};
            padding: 5px;
        }}
        td {{
            background-color: {cor_base_tabela};
            color: {cor_texto_tabela};
            height: {altura_linhas_tabela}px;
            text-align: {alinhamento_horizontal_texto.lower()};
            vertical-align: {alinhamento_vertical_texto.lower()};
            padding: 5px;
        }}
        tr:nth-child(even) td {{
            background-color: {cor_base_tabela};
        }}
    </style>
    <table border="1">
    """

    # Criar linhas da tabela
    for i, linha in enumerate(dados):
        tabela_html += "<tr>"
        for celula in linha:
            if i == 0:
                tabela_html += f"<th>{celula}</th>"  # Cabeçalho
            else:
                tabela_html += f"<td>{celula}</td>"  # Corpo
        tabela_html += "</tr>"

    tabela_html += "</table>"

    # Exibir a tabela no Streamlit
    st.markdown(tabela_html, unsafe_allow_html=True)

st.divider()

bot1, bot2 = st.columns(2)

config_salva = False
with bot1:
    # Salvar configurações
    if st.button("Salvar Configurações", type='primary'):
        config = {
            "cor_cabecalho": cor_cabecalho,
            "cor_fonte_cabecalho": cor_fonte_cabecalho,
            "tamanho_fonte_cabecalho": tamanho_fonte_cabecalho,
            "altura_linhas_cabecalho": altura_linhas_cabecalho,
            "alinhamento_horizontal_cabecalho": alinhamento_horizontal_cabecalho.lower(),
            "alinhamento_vertical_cabecalho": alinhamento_vertical_cabecalho.lower(),
            "cor_fundo_tabela": cor_base_tabela,
            "cor_texto_tabela": cor_texto_tabela,
            "tamanho_fonte_tabela": tamanho_fonte_tabela,
            "altura_linhas_tabela": altura_linhas_tabela,
            "alinhamento_horizontal_texto": alinhamento_horizontal_texto.lower(),
            "alinhamento_vertical_texto": alinhamento_vertical_texto.lower()
        }
        salvar_configuracoes(config)
        config_salva = True
        st.success("Configurações salvas com sucesso!")

    with bot2:
        if config_salva:
            st.download_button("Baixar Configurações", json.dumps(config, indent=4), "config_excel.json", "json", key="download_config")