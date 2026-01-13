import pandas as pd
from io import BytesIO

# =========================
# LOADERS
# =========================
def load_organico(organico_file) -> pd.DataFrame:
    org = pd.read_excel(organico_file, header=4)
    org.columns = [
        "ID","Titulo","Status","Variacao","SKU",
        "Visitas","Qtd_Vendas","Compradores",
        "Unidades","Vendas_Brutas","Participacao",
        "Conv_Visitas_Vendas","Conv_Visitas_Compradores"
    ]
    org = org[org["ID"] != "ID do anúncio"].copy()

    for c in ["Visitas","Qtd_Vendas","Compradores","Unidades","Vendas_Brutas",
              "Participacao","Conv_Visitas_Vendas","Conv_Visitas_Compradores"]:
        org[c] = pd.to_numeric(org[c], errors="coerce")

    org["ID"] = org["ID"].astype(str)
    return org


def load_patrocinados(patrocinados_file) -> pd.DataFrame:
    pat = pd.read_excel(patrocinados_file, sheet_name="Relatório Anúncios patrocinados", header=1)
    pat["ID"] = pat["Código do anúncio"].astype(str).str.replace("MLB", "", regex=False)

    for c in ["Impressões","Cliques","Receita
(Moeda local)","Investimento
(Moeda local)",
              "Vendas por publicidade
(Diretas + Indiretas)"]:
        if c in pat.columns:
            pat[c] = pd.to_numeric(pat[c], errors="coerce")

    return pat


# =========================
# CAMPANHAS: DIÁRIO vs CONSOLIDADO
# =========================
def _coerce_campaign_numeric(df: pd.DataFrame) -> pd.DataFrame:
    cols_num = [
        "Impressões","Cliques","Receita
(Moeda local)","Investimento
(Moeda local)",
        "Vendas por publicidade
(Diretas + Indiretas)","ROAS
(Receitas / Investimento)",
        "CVR
(Conversion rate)","% de impressões perdidas por orçamento",
        "% de impressões perdidas por classificação","Orçamento","ACOS Objetivo"
    ]
    for c in cols_num:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def load_campanhas_diario(campanhas_file) -> pd.DataFrame:
    camp = pd.read_excel(campanhas_file, sheet_name="Relatório de campanha", header=1)
    camp["Desde"] = pd.to_datetime(camp["Desde"], errors="coerce")
    camp = _coerce_campaign_numeric(camp)
    return camp


def load_campanhas_consolidado(campanhas_file) -> pd.DataFrame:
    camp = pd.read_excel(campanhas_file, sheet_name="Relatório de campanha", header=1)
    # No consolidado normalmente não tem a coluna "Desde"
    camp = _coerce_campaign_numeric(camp)
    return camp


def build_daily_from_diario(camp_diario: pd.DataFrame) -> pd.DataFrame:
    daily = camp_diario.groupby("Desde", as_index=False).agg(
        Investimento=("Investimento\n(Moeda local)", "sum"),
        Receita=("Receita\n(Moeda local)", "sum"),
        Vendas=("Vendas por publicidade\n(Diretas + Indiretas)", "sum"),
        Cliques=("Cliques", "sum"),
        Impressoes=("Impressões", "sum"),
    )
    return daily.sort_values("Desde")


def build_campaign_agg(camp: pd.DataFrame, modo: str) -> pd.DataFrame:
    if modo == "diario":
        # agrega várias linhas (1 por dia) em 1 linha por campanha
        camp_agg = camp.groupby("Nome", as_index=False).agg(
            Status=("Status", "last"),
            Orçamento=("Orçamento", "last"),
            **{
                "ACOS Objetivo": ("ACOS Objetivo", "last"),
                "Impressões": ("Impressões", "sum"),
                "Cliques": ("Cliques", "sum"),
                "Receita": ("Receita\n(Moeda local)", "sum"),
                "Investimento": ("Investimento\n(Moeda local)", "sum"),
                "Vendas": ("Vendas por publicidade\n(Diretas + Indiretas)", "sum"),
                "ROAS": ("ROAS\n(Receitas / Investimento)", "mean"),
                "CVR": ("CVR\n(Conversion rate)", "mean"),
                "Perdidas_Orc": ("% de impressões perdidas por orçamento", "mean"),
                "Perdidas_Class": ("% de impressões perdidas por classificação", "mean"),
            }
        )
        return camp_agg

    # CONSOLIDADO: já vem 1 linha por campanha
    camp_agg = camp.rename(columns={
        "Receita\n(Moeda local)": "Receita",
        "Investimento\n(Moeda local)": "Investimento",
        "Vendas por publicidade\n(Diretas + Indiretas)": "Vendas",
        "ROAS\n(Receitas / Investimento)": "ROAS",
        "CVR\n(Conversion rate)": "CVR",
        "% de impressões perdidas por orçamento": "Perdidas_Orc",
        "% de impressões perdidas por classificação": "Perdidas_Class",
    }).copy()

    needed = [
        "Nome","Status","Orçamento","ACOS Objetivo",
        "Impressões","Cliques","Receita","Investimento","Vendas",
        "ROAS","CVR","Perdidas_Orc","Perdidas_Class"
    ]
    for col in needed:
        if col not in camp_agg.columns:
            camp_agg[col] = pd.NA

    camp_agg = camp_agg[needed].copy()
    return camp_agg


# =========================
# TABELAS DE DECISÃO + KPIs
# =========================
def build_tables(
    org: pd.DataFrame,
    camp_agg: pd.DataFrame,
    pat: pd.DataFrame,
    enter_visitas_min: int = 50,
    enter_conv_min: float = 0.05,   # regra: > 5%
    pause_invest_min: float = 100.0,
    pause_cvr_max: float = 0.01,    # 1%
    scale_lost_budget_min: float = 20.0,
    scale_cvr_min: float = 0.02,    # 2%
    scale_roas_min: float = 6.0,
    acos_lost_rank_min: float = 30.0,
    acos_roas_min: float = 7.0
):
    # Pausar
    pause = camp_agg[
        (camp_agg["Investimento"] > pause_invest_min) &
        ((camp_agg["Vendas"] <= 0) | (camp_agg["CVR"] < pause_cvr_max))
    ].copy()
    pause["Ação"] = "PAUSAR"
    pause = pause.sort_values("Investimento", ascending=False)

    # Entrar em Ads
    ads_ids = set(pat["ID"].dropna().astype(str).unique())
    enter = org[
        (org["Visitas"] >= enter_visitas_min) &
        (org["Conv_Visitas_Vendas"] > enter_conv_min) &
        (~org["ID"].isin(ads_ids))
    ].copy()
    enter["Codigo_MLB"] = "MLB" + enter["ID"]
    enter["Ação"] = "INSERIR EM ADS"
    enter = enter.sort_values(["Conv_Visitas_Vendas","Visitas"], ascending=[False, False])
    enter = enter[["ID","Codigo_MLB","Titulo","Conv_Visitas_Vendas","Visitas","Qtd_Vendas","Vendas_Brutas","Ação"]]

    # Escalar orçamento
    scale = camp_agg[
        (camp_agg["Perdidas_Orc"] > scale_lost_budget_min) &
        (camp_agg["CVR"] >= scale_cvr_min) &
        (camp_agg["ROAS"] >= scale_roas_min)
    ].copy()
    scale["Ação"] = "AUMENTAR ORÇAMENTO (+10% a +20%)"
    scale = scale.sort_values("Perdidas_Orc", ascending=False)

    # Ajustar ACOS
    acos = camp_agg[
        (camp_agg["Perdidas_Class"] > acos_lost_rank_min) &
        (camp_agg["ROAS"] >= acos_roas_min)
    ].copy()
    acos["Ação"] = "AUMENTAR ACOS OBJETIVO"
    acos = acos.sort_values("Perdidas_Class", ascending=False)

    invest_total = float(pd.to_numeric(camp_agg["Investimento"], errors="coerce").fillna(0).sum())
    receita_total = float(pd.to_numeric(camp_agg["Receita"], errors="coerce").fillna(0).sum())
    vendas_total = int(pd.to_numeric(camp_agg["Vendas"], errors="coerce").fillna(0).sum())
    roas_total = (receita_total / invest_total) if invest_total else 0.0

    kpis = {
        "Campanhas únicas": int(camp_agg["Nome"].nunique()),
        "IDs patrocinados únicos": int(pat["ID"].nunique()),
        "Investimento Ads (R$)": invest_total,
        "Receita Ads (R$)": receita_total,
        "Vendas Ads": vendas_total,
        "ROAS": roas_total,
    }

    return kpis, pause, enter, scale, acos


def gerar_excel(kpis, camp_agg, pause, enter, scale, acos) -> bytes:
    resumo = pd.DataFrame([kpis])

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo.to_excel(writer, index=False, sheet_name="RESUMO")
        pause.to_excel(writer, index=False, sheet_name="PAUSAR CAMPANHAS")
        enter.to_excel(writer, index=False, sheet_name="ENTRAR EM ADS")
        scale.to_excel(writer, index=False, sheet_name="ESCALAR ORÇAMENTO")
        acos.to_excel(writer, index=False, sheet_name="AJUSTAR ACOS")
        camp_agg.to_excel(writer, index=False, sheet_name="BASE CAMPANHAS (AGG)")
    out.seek(0)
    return out.read()
