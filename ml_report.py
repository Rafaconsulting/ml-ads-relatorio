import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Relatório ML Ads", layout="wide")
st.title("Relatório Automático – Mercado Livre Ads")

st.write("1) Faça upload dos 3 relatórios do Mercado Livre (XLSX).")
st.write("2) Clique em **Gerar relatório**.")
st.write("3) Baixe o Excel final com as abas: PAUSAR / ENTRAR / ESCALAR / ACOS.")

col1, col2, col3 = st.columns(3)

with col1:
    organico = st.file_uploader("Relatório orgânico (publicações)", type=["xlsx"])
with col2:
    campanhas = st.file_uploader("Relatório campanhas Ads (diário)", type=["xlsx"])
with col3:
    patrocinados = st.file_uploader("Relatório anúncios patrocinados", type=["xlsx"])

if organico and campanhas and patrocinados:
    if st.button("Gerar relatório"):
        with st.spinner("Processando..."):
            bytes_xlsx = gerar_relatorio(organico, campanhas, patrocinados)

        nome = f"Relatorio_ML_ADs_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        st.success("Pronto! Clique abaixo para baixar.")
        st.download_button(
            "Baixar Excel",
            data=bytes_xlsx,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Envie os 3 arquivos para liberar o botão.")
