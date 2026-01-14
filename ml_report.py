import pandas as pd
from io import BytesIO

EMOJI_GREEN = "\U0001F7E2"
EMOJI_YELLOW = "\U0001F7E1"
EMOJI_BLUE = "\U0001F535"
EMOJI_RED = "\U0001F534"


def load_organico(file):
    df = pd.read_excel(file, header=4)
    df.columns = [
        "ID","Titulo","Status","Variacao","SKU",
        "Visitas","Qtd_Vendas","Compradores",
        "Unidades","Vendas_Brutas","Participacao",
        "Conv_Visitas_Vendas","Conv_Visitas_Compradores"
    ]
    df = df[df["ID"] != "ID do anúncio"]
    df["ID"] = df["ID"].astype(str).str.replace("MLB", "", regex=False)
    return df


def load_patrocinados(file):
    df = pd.read_excel(file, sheet_name="Relatório Anúncios patrocinados", header=1)
    df["ID"] = df["Código do anúncio"].astype(str).str.replace("MLB", "", regex=False)
    return df


def load_campanhas_consolidado(file):
    return pd.read_excel(file, sheet_name="Relatório de campanha", header=1)


def load_campanhas_diario(file):
    df = pd.read_excel(file, sheet_name="Relatório de campanha", header=1)
    df["Desde"] = pd.to_datetime(df["Desde"], errors="coerce")
    return df


def build_campaign_agg(df, modo):
    if modo == "diario":
        return df.groupby("Nome", as_index=False).agg(
            Receita=("Receita\n(Moeda local)", "sum"),
            Investimento=("Investimento\n(Moeda local)", "sum"),
            Vendas=("Vendas por publicidade\n(Diretas + Indiretas)", "sum"),
            ROAS=("ROAS\n(Receitas / Investimento)", "mean"),
            Perdidas_Orc=("% de impressões perdidas por orçamento", "mean"),
            Perdidas_Class=("% de impressões perdidas por classificação", "mean"),
            Orçamento=("Orçamento", "last"),
            ACOS_Objetivo=("ACOS Objetivo", "last")
        )

    return df.rename(columns={
        "Receita\n(Moeda local)": "Receita",
        "Investimento\n(Moeda local)": "Investimento",
        "Vendas por publicidade\n(Diretas + Indiretas)": "Vendas",
        "% de impressões perdidas por orçamento": "Perdidas_Orc",
        "% de impressões perdidas por classificação": "Perdidas_Class",
    })


def build_tables(org, camp, pat, **_):
    camp["ROAS_Real"] = camp["Receita"] / camp["Investimento"]
    camp["Acao_Recomendada"] = camp["ROAS_Real"].apply(
        lambda r: f"{EMOJI_GREEN} Escalar" if r >= 7 else f"{EMOJI_RED} Revisar"
    )

    kpis = {
        "Investimento Ads (R$)": camp["Investimento"].sum(),
        "Receita Ads (R$)": camp["Receita"].sum(),
        "Vendas Ads": camp["Vendas"].sum(),
        "ROAS": camp["Receita"].sum() / camp["Investimento"].sum()
    }

    return kpis, camp, org, camp, camp, camp


def build_executive_diagnosis(df, daily=None):
    return {
        "ROAS": df["ROAS_Real"].mean(),
        "ACOS_real": 1 / df["ROAS_Real"].mean(),
        "Veredito": "Estamos deixando dinheiro na mesa.",
        "Tendencias": {"cpc_proxy_up": None, "ticket_down": None, "roas_down": None}
    }


def build_control_panel(df):
    return df[["Nome","ROAS_Real","Perdidas_Orc","Perdidas_Class","Acao_Recomendada"]]


def build_opportunity_highlights(df):
    return {
        "Locomotivas": df.sort_values("Receita", ascending=False).head(5),
        "Minas": df[df["ROAS_Real"] >= 7]
    }


def build_7_day_plan(df):
    return df[["Nome","Acao_Recomendada"]]


def gerar_excel(*_):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame([{"OK": "Relatorio gerado"}]).to_excel(writer, index=False)
    out.seek(0)
    return out.read()
