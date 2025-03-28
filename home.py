import streamlit as st

home_page = st.Page("home.py", title="Home", icon=":material/home:")
config_page = st.Page("config.py", title="Configure seu Excel", icon=":material/edit:")
relatorio_page = st.Page("app.py", title="Gerar Relatórios", icon=":material/check_circle:")

pg = st.navigation([home_page, config_page, relatorio_page])
st.set_page_config(page_title="Home", page_icon=":material/edit:")

st.title("Seu Assistente de Planilhas!")

st.write(
    "Cansado de perder tempo separando e organizando planilhas manualmente? "
    "Seu problema acabou! Nosso site faz isso de forma rápida, ágil e personalizável. "
    "Diga adeus ao trabalho repetitivo e foque no que realmente importa, seu tempo!!"
)

st.subheader("Objetivo do site?")
st.write(
    "🔹 Separar automaticamente planilhas por fornecedor ou cliente.\n\n"
    "🔹 Personalizar arquivos com estilo pré-definido pelos usuários.\n\n"
    "🔹 Economizar seu tempo e reduzir erros!"
)

st.subheader("Por que usar?")
st.write(
    "✔️ Simples e intuitivo!\n\n"
    "✔️ Sem complicação, sem fórmulas mágicas, só eficiência!\n\n"
    "✔️ Feito para quem precisa lidar com muitas planilhas sem perder a paciência!"
)

st.subheader("Proximos passos:")
st.write(
    "⚠️ Incluir opção de escolher por qual coluna será separado o arquivo excel!\n\n"
    "⚠️ Possibilitar que o usuário faça a inclusão de uma imagem!\n\n"
    "⚠️ Download em Excel ou PDF!\n\n"
    "⚠️ Opção de download de arquivo com a pré-configuração das colunas padrões!"
)

st.success("Pronto para agilizar sua rotina? Vamos nessa!")