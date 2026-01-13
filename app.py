import streamlit as st
from datetime import datetime
import ml_report as ml

st.set_page_config(
    page_title="ML Ads – Dashboard & Relatório",
    layout="wide"
)

st.title("Mercado Livre Ads – Dashboard e Relatório Automático")

# =========================
# TOGGLE DE MODO
# =========================
modo = st.radio(
    "Selecione o tipo de relatório de campanhas que você exportou:",
    ["CONSOLIDADO (decisão)", "DIÁRIO (monitoramento)"],
    horizontal=True
)
modo_key = "consolidado" if "CONSOLIDADO" in modo else "diario"

# =========================
# REGRAS AJUSTÁVEIS
# =========================
with st.expander("⚙️ Regras (ajustáveis)"):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        enter_visitas_min = st.number_input(
            "ENTRAR: Visitas mín.",
            min_value=0,
            value=50,
            step=10
        )

    with c2:
        enter_conv_pct = st.number_input(
            "ENTRAR: Conversão orgânica mín. (%)",
            min_value=0.0,
            value=5.0,
            step=0.5
        )

    with c3:
        pause_invest_min = st.number_input(
            "PAUSAR: Investimento mín. (R$)",
            min_value=0.0,
            value=100.0,
            step=50.0
        )

    with c4:
        pause_cvr_pct = st.number_input(
            "PAUSAR: CVR máx. (%)",
            min_value=0.0,
            value=1.0,
            step=0.2
        )

# =========================
# UPLOAD DOS ARQUIVOS
# =========================
u1, u2, u3 = st.columns(3)

with u1:
    organico_file = st.file_uploader(
        "Relatório orgânico (publicações)",
        type=["xlsx"]
    )

with u2:
    campanhas_file = st.file_uploader(
        "Relatório campanhas Ads",
        type=["xlsx"]
    )

with u3:
    patrocinados_file = st.file_uploader(
        "Relatório anúncios patrocinados",
        type=["xlsx"]
    )

if not (organico_file and campanhas_file and patrocinados_file):
    st.info("Envie os 3 arquivos para liberar o Dashboard.")
    st.stop()

# =========================
# CARREGAMENTO DOS DADOS
# =========================
with st.spinner("Lendo arquivos..."):
    org = ml.load_organico(organico_file)
    pat = ml.load_patrocinados(patrocinados_file)

    if modo_key == "diario":
        camp = ml.load_campanhas_diario(campanhas_file)
    else:
        camp = ml.load_campanhas_consolidado(campanhas_file)

camp_agg = ml.build_campaign_agg(camp, modo_key)

kpis, pause, enter, scale, acos = ml.build_tables(
    org,
    camp_agg,
    pat,
    enter_visitas_min=int(enter_visitas_min),
    enter_conv_min=float(enter_conv_pct) / 100.0,
    pause_invest_min=float(pause_invest_min),
    pause_cvr_max=float(pause_cvr_pct) / 100.0
)

# =========================
# TABS
# =========================
tab1, tab2 = st.tabs(["📊 Dashboard", "📄 Gerar Relatório (Excel)"])

# =========================
# DASHBOARD
# =========================
with tab1:
    st.subheader("KPIs")

    k1, k2, k3, k4, k5, k6 = st.columns(6)

    k1.metric("Investimento", f"R$ {kpis['Investimento Ads (R$)']:.2f}")
    k2.metric("Receita Ads", f"R$ {kpis['Receita Ads (R$)']:.2f}")
    k3.metric("Vendas Ads", kpis["Vendas Ads"])
    k4.metric("ROAS", f"{kpis['ROAS']:.2f}")
    k5.metric("Campanhas únicas", kpis["Campanhas únicas"])
    k6.metric("IDs patrocinados", kpis["IDs patrocinados únicos"])

    st.divider()

    if modo_key == "diario":
        st.subheader("Evolução diária")
        daily = ml.build_daily_from_diario(camp).set_index("Desde")
        st.line_chart(daily[["Investimento", "Receita", "Vendas"]])
        st.divider()

    # =========================
    # SELETORES TOP 10
    # =========================
    col1, col2 = st.columns([2, 1])

    with col1:
        metrica_top10 = st.selectbox(
            "Top 10 campanhas por:",
            ["Receita", "Investimento", "ROAS", "Vendas"]
        )

    with col2:
        ordem = st.selectbox(
            "Ordem:",
            ["Maior → Menor", "Menor → Maior"]
        )

    asc = ordem == "Menor → Maior"

    st.subheader(f"Top 10 campanhas por {metrica_top10}")

    bar = camp_agg.copy()

    for col in ["Receita", "Investimento", "Vendas", "ROAS", "CVR"]:
        if col in bar.columns:
            bar[col] = bar[col].astype(float)

    bar = (
        bar.sort_values(metrica_top10, ascending=asc)
           .head(10)
           .set_index("Nome")
    )

    st.bar_chart(bar[[metrica_top10]])

    st.divider()

    # =========================
    # TABELAS DE ACAO
    # =========================
    cA, cB = st.columns(2)

    with cA:
        st.subheader("Campanhas para ESCALAR")
        st.dataframe(scale, use_container_width=True)

    with cB:
        st.subheader("Campanhas para AJUSTAR ACOS")
        st.dataframe(acos, use_container_width=True)

    cC, cD = st.columns(2)

    with cC:
        st.subheader("Campanhas para PAUSAR")
        st.dataframe(pause, use_container_width=True)

    with cD:
        st.subheader("Anúncios para ENTRAR em Ads")
        st.dataframe(enter, use_container_width=True)

# =========================
# EXPORT EXCEL
# =========================
with tab2:
    st.subheader("Gerar relatório final (Excel)")

    if st.button("Gerar e baixar Excel"):
        with st.spinner("Gerando Excel..."):
            bytes_xlsx = ml.gerar_excel(
                kpis,
                camp_agg,
                pause,
                enter,
                scale,
                acos
            )

        nome = f"Relatorio_ML_ADs_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

        st.download_button(
            "Baixar Excel",
            data=bytes_xlsx,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Relatório gerado com sucesso.")
