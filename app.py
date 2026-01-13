import streamlit as st
from datetime import datetime
import pandas as pd

from ml_report import (
    load_organico, load_campanhas, load_patrocinados,
    build_tables, build_daily, gerar_excel
)

st.set_page_config(page_title="ML Ads – Dashboard & Relatório", layout="wide")
st.title("Mercado Livre Ads – Dashboard e Relatório Automático")

st.write("1) Envie os 3 relatórios (XLSX) | 2) Veja o Dashboard | 3) Gere o Excel final")

# ----- Regras ajustáveis (na tela) -----
with st.expander("⚙️ Regras (ajustáveis)"):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        enter_visitas_min = st.number_input("ENTRAR: Visitas mín.", min_value=0, value=50, step=10)
    with c2:
        enter_conv_pct = st.number_input("ENTRAR: Conversão orgânica mín. (%)", min_value=0.0, value=5.0, step=0.5)
    with c3:
        pause_invest_min = st.number_input("PAUSAR: Investimento mín. (R$)", min_value=0.0, value=100.0, step=50.0)
    with c4:
        pause_cvr_pct = st.number_input("PAUSAR: CVR máx. (%)", min_value=0.0, value=1.0, step=0.2)

# ----- Upload -----
u1, u2, u3 = st.columns(3)
with u1:
    organico_file = st.file_uploader("Relatório orgânico (publicações)", type=["xlsx"])
with u2:
    campanhas_file = st.file_uploader("Relatório campanhas Ads (diário)", type=["xlsx"])
with u3:
    patrocinados_file = st.file_uploader("Relatório anúncios patrocinados", type=["xlsx"])

if not (organico_file and campanhas_file and patrocinados_file):
    st.info("Envie os 3 arquivos para liberar o Dashboard e o botão de gerar Excel.")
    st.stop()

# ----- Carregar dados -----
with st.spinner("Lendo arquivos..."):
    org = load_organico(organico_file)
    camp = load_campanhas(campanhas_file)
    pat = load_patrocinados(patrocinados_file)

# ----- Construir tabelas -----
kpis, camp_agg, pause, enter, scale, acos = build_tables(
    org, camp, pat,
    enter_visitas_min=int(enter_visitas_min),
    enter_conv_min=float(enter_conv_pct)/100.0,
    pause_invest_min=float(pause_invest_min),
    pause_cvr_max=float(pause_cvr_pct)/100.0,
)
daily = build_daily(camp)

tab1, tab2 = st.tabs(["📊 Dashboard", "📄 Gerar Relatório (Excel)"])

# =========================
# DASHBOARD
# =========================
with tab1:
    st.subheader("KPIs (15 dias)")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Investimento", f"R$ {kpis['Investimento Ads (R$)']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k2.metric("Receita Ads", f"R$ {kpis['Receita Ads (R$)']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k3.metric("Vendas Ads", f"{kpis['Vendas Ads']}")
    k4.metric("ROAS", f"{kpis['ROAS']:.2f}")
    k5.metric("Campanhas únicas", f"{kpis['Campanhas únicas']}")
    k6.metric("IDs patrocinados", f"{kpis['IDs patrocinados únicos']}")

    st.divider()

    st.subheader("Evolução diária (Investimento x Receita x Vendas)")
    chart_df = daily.copy()
    chart_df = chart_df.set_index("Desde")
    st.line_chart(chart_df[["Investimento", "Receita", "Vendas"]])

    st.divider()

    st.subheader("Campanhas – visão rápida")
    st.caption("Dispersão: ROAS x CVR | tamanho ~ investimento (quanto maior, mais gastou)")
    scatter = camp_agg.copy()
    scatter["CVR_%"] = scatter["CVR"] * 100
    st.scatter_chart(scatter, x="ROAS", y="CVR_%", size="Investimento")

    st.divider()

    cA, cB = st.columns(2)
    with cA:
        st.subheader("Campanhas para ESCALAR orçamento")
        st.dataframe(scale, use_container_width=True)
    with cB:
        st.subheader("Campanhas para AJUSTAR ACOS")
        st.dataframe(acos, use_container_width=True)

    cC, cD = st.columns(2)
    with cC:
        st.subheader("Campanhas para PAUSAR")
        st.dataframe(pause, use_container_width=True)
    with cD:
        st.subheader("Anúncios orgânicos para ENTRAR em Ads")
        st.dataframe(enter, use_container_width=True)

# =========================
# GERAR EXCEL
# =========================
with tab2:
    st.subheader("Gerar Excel final com decisões")
    st.write("Esse Excel vai sair com as abas: RESUMO, PAUSAR CAMPANHAS, ENTRAR EM ADS, ESCALAR ORÇAMENTO, AJUSTAR ACOS, BASE CAMPANHAS (AGG).")

    if st.button("Gerar e baixar Excel"):
        with st.spinner("Gerando Excel..."):
            bytes_xlsx = gerar_excel(kpis, camp_agg, pause, enter, scale, acos)

        nome = f"Relatorio_ML_ADs_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        st.download_button(
            "Baixar Excel",
            data=bytes_xlsx,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Pronto ✅")
