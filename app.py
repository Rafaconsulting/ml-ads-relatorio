import streamlit as st
from datetime import datetime
import ml_report as ml

st.set_page_config(page_title="ML Ads - Dashboard & Relatorio", layout="wide")
st.title("Mercado Livre Ads - Dashboard e Relatorio Automatico (Estrategico)")

modo = st.radio(
    "Selecione o tipo de relatorio de campanhas que voce exportou:",
    ["CONSOLIDADO (decisao)", "DIARIO (monitoramento)"],
    horizontal=True
)
modo_key = "consolidado" if "CONSOLIDADO" in modo else "diario"

with st.expander("Regras (ajustaveis)"):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        enter_visitas_min = st.number_input("ENTRAR: Visitas min.", min_value=0, value=50, step=10)
    with c2:
        enter_conv_pct = st.number_input("ENTRAR: Conversao organica min. (%)", min_value=0.0, value=5.0, step=0.5)
    with c3:
        pause_invest_min = st.number_input("PAUSAR: Investimento min. (R$)", min_value=0.0, value=100.0, step=50.0)
    with c4:
        pause_cvr_pct = st.number_input("PAUSAR: CVR max. (%)", min_value=0.0, value=1.0, step=0.2)

u1, u2, u3 = st.columns(3)
with u1:
    organico_file = st.file_uploader("Relatorio organico (publicacoes)", type=["xlsx"])
with u2:
    campanhas_file = st.file_uploader("Relatorio campanhas Ads", type=["xlsx"])
with u3:
    patrocinados_file = st.file_uploader("Relatorio anuncios patrocinados", type=["xlsx"])

if not (organico_file and campanhas_file and patrocinados_file):
    st.info("Envie os 3 arquivos para liberar o dashboard.")
    st.stop()

with st.spinner("Lendo arquivos..."):
    org = ml.load_organico(organico_file)
    pat = ml.load_patrocinados(patrocinados_file)
    camp = ml.load_campanhas_diario(campanhas_file) if modo_key == "diario" else ml.load_campanhas_consolidado(campanhas_file)

camp_agg = ml.build_campaign_agg(camp, modo_key)

kpis, pause, enter, scale, acos, camp_strat = ml.build_tables(
    org, camp_agg, pat,
    enter_visitas_min=int(enter_visitas_min),
    enter_conv_min=float(enter_conv_pct) / 100.0,
    pause_invest_min=float(pause_invest_min),
    pause_cvr_max=float(pause_cvr_pct) / 100.0,
)

daily = None
if modo_key == "diario":
    daily = ml.build_daily_from_diario(camp)

diagnosis = ml.build_executive_diagnosis(camp_strat, daily=daily)
panel = ml.build_control_panel(camp_strat)
high = ml.build_opportunity_highlights(camp_strat)
plan7 = ml.build_7_day_plan(camp_strat)

tab1, tab2 = st.tabs(["Dashboard", "Gerar Excel"])

with tab1:
    st.subheader("Diagnostico Executivo")
    st.write(f"ROAS da conta: **{diagnosis['ROAS']:.2f}** | ACOS real: **{diagnosis['ACOS_real']:.2f}**")
    st.write(f"**Veredito:** {diagnosis['Veredito']}")

    if diagnosis["Tendencias"]["cpc_proxy_up"] is not None:
        st.caption(
            f"Tendencias (ultimos 7d vs 7d anteriores) | "
            f"CPC proxy: {diagnosis['Tendencias']['cpc_proxy_up']:+.1%} | "
            f"Ticket: {diagnosis['Tendencias']['ticket_down']:+.1%} | "
            f"ROAS: {diagnosis['Tendencias']['roas_down']:+.1%}"
        )

    st.divider()
    st.subheader("Painel de Controle Geral")
    st.dataframe(panel, use_container_width=True)

    st.divider()
    st.subheader("Plano de Acao - 7 dias")
    st.dataframe(plan7, use_container_width=True)

with tab2:
    if st.button("Gerar e baixar Excel"):
        bytes_xlsx = ml.gerar_excel(
            kpis, camp_agg, pause, enter, scale, acos, camp_strat, daily=daily
        )
        nome = f"Relatorio_ML_ADs_Estrategico_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
        st.download_button("Baixar Excel", bytes_xlsx, nome)
