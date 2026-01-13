import pandas as pd
from io import BytesIO

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


def load_campanhas(campanhas_file) -> pd.DataFrame:
    camp = pd.read_excel(campanhas_file, sheet_name="Relatório de campanha", header=1)
    camp["Desde"] = pd.to_datetime(camp["Desde"], errors="coerce")

    cols_num = [
        "Impressões","Cliques","Receita\n(Moeda local)","Investimento\n(Moeda local)",
        "Vendas por publicidade\n(Diretas + Indiretas)","ROAS\n(Receitas / Investimento)",
        "CVR\n(Conversion rate)","% de impressões perdidas por orçamento",
        "% de impressões perdidas por classificação"
    ]
    for c in cols_num:
        camp[c] = pd.to_numeric(camp[c], errors="coerce")

    return camp


def load_patrocinados(patrocinados_file) -> pd.DataFrame:
    pat = pd.read_excel(patrocinados_file, sheet_name="Relatório Anúncios patrocinados", header=1)
    pat["ID"] = pat["Código do anúncio"].astype(str).str.replace("MLB", "", regex=False)

    # Colunas numéricas (nem sempre usadas no dashboard, mas já deixa pronto)
    for c in ["Impressões","Cliques","Receita\n(Moeda local)","Investimento\n(Moeda local)",
              "Vendas por publicidade\n(Diretas + Indiretas)"]:
        if c in pat.columns:
            pat[c] = pd.to_numeric(pat[c], errors="coerce")

    return pat


def build_tables(
    org: pd.DataFrame,
    camp: pd.DataFrame,
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
    # Agregado por campanha (base)
    camp_agg = camp.groupby("Nome", as_index=False).agg(
        Investimento=("Investimento\n(Moeda local)", "sum"),
        Receita=("Receita\n(Moeda local)", "sum"),
        Vendas=("Vendas por publicidade\n(Diretas + Indiretas)", "sum"),
        CVR=("CVR\n(Conversion rate)", "mean"),
        ROAS=("ROAS\n(Receitas / Investimento)", "mean"),
        Perdidas_Orc=("% de impressões perdidas por orçamento", "mean"),
        Perdidas_Class=("% de impressões perdidas por classificação", "mean")
    )

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

    # KPIs
    invest_total = float(camp["Investimento\n(Moeda local)"].sum())
    receita_total = float(camp["Receita\n(Moeda local)"].sum())
    vendas_total = int(camp["Vendas por publicidade\n(Diretas + Indiretas)"].sum())
    roas_total = (receita_total / invest_total) if invest_total else 0.0

    kpis = {
        "Campanhas únicas": int(camp["Nome"].nunique()),
        "IDs patrocinados únicos": int(pat["ID"].nunique()),
        "Investimento Ads (R$)": invest_total,
        "Receita Ads (R$)": receita_total,
        "Vendas Ads": vendas_total,
        "ROAS": roas_total,
    }

    return kpis, camp_agg, pause, enter, scale, acos


def build_daily(camp: pd.DataFrame) -> pd.DataFrame:
    daily = camp.groupby("Desde", as_index=False).agg(
        Investimento=("Investimento\n(Moeda local)", "sum"),
        Receita=("Receita\n(Moeda local)", "sum"),
        Vendas=("Vendas por publicidade\n(Diretas + Indiretas)", "sum"),
        Cliques=("Cliques","sum"),
        Impressoes=("Impressões","sum"),
    )
    return daily.sort_values("Desde")


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
