# -*- coding: utf-8 -*-
import pandas as pd
from io import BytesIO
import unicodedata


def _norm(s: str) -> str:
    """Lower + remove accents/diacritics + trim."""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()


def _find_sheet(xls: pd.ExcelFile, keywords):
    """Find a sheet by keywords ignoring accents."""
    names = xls.sheet_names
    norm_names = [_norm(n) for n in names]
    for i, nn in enumerate(norm_names):
        ok = True
        for kw in keywords:
            if _norm(kw) not in nn:
                ok = False
                break
        if ok:
            return names[i]
    return None


# =========================
# LOADERS
# =========================
def load_organico(organico_file) -> pd.DataFrame:
    # Usually first sheet
    df = pd.read_excel(organico_file, header=4)
    # Standardize columns if possible; otherwise keep and map later
    # Expect: ID, title, status, visits, sales, etc.
    cols = list(df.columns)

    # Try to rename common Portuguese headers to stable names
    rename_map = {}
    for c in cols:
        cn = _norm(c)
        if "id" in cn and "anuncio" in cn:
            rename_map[c] = "ID"
        elif cn in ("titulo", "título"):
            rename_map[c] = "Titulo"
        elif "visitas" in cn and "unica" in cn:
            rename_map[c] = "Visitas"
        elif cn == "visitas":
            rename_map[c] = "Visitas"
        elif "qtd" in cn and "venda" in cn:
            rename_map[c] = "Qtd_Vendas"
        elif "vendas brutas" in cn or ("vendas" in cn and "bruta" in cn):
            rename_map[c] = "Vendas_Brutas"
        elif "conversao" in cn and "visitas" in cn and "vendas" in cn:
            rename_map[c] = "Conv_Visitas_Vendas"

    df = df.rename(columns=rename_map)

    # Remove header-like row if present
    if "ID" in df.columns:
        df = df[df["ID"].astype(str).str.lower().str.contains("id") == False].copy()

    # Coerce
    if "ID" in df.columns:
        df["ID"] = df["ID"].astype(str).str.replace("MLB", "", regex=False)

    for c in ["Visitas", "Qtd_Vendas", "Vendas_Brutas", "Conv_Visitas_Vendas"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Keep safe columns
    keep = []
    for c in ["ID", "Titulo", "Visitas", "Qtd_Vendas", "Vendas_Brutas", "Conv_Visitas_Vendas"]:
        if c in df.columns:
            keep.append(c)
    if not keep:
        # fallback: return raw
        return df
    return df[keep].copy()


def load_patrocinados(patrocinados_file) -> pd.DataFrame:
    xls = pd.ExcelFile(patrocinados_file)
    # Find sheet by keywords (no accents)
    sheet = _find_sheet(xls, ["anuncios", "patrocinados"])
    if sheet is None:
        sheet = xls.sheet_names[0]

    df = pd.read_excel(patrocinados_file, sheet_name=sheet, header=1)

    # Find ad code column
    code_col = None
    for c in df.columns:
        if "codigo" in _norm(c) and "anuncio" in _norm(c):
            code_col = c
            break

    if code_col is not None:
        df["ID"] = df[code_col].astype(str).str.replace("MLB", "", regex=False)
    else:
        # fallback
        df["ID"] = pd.NA

    # Coerce common numerics (if present)
    num_cols = [
        "Impressoes", "Cliques", "Receita", "Investimento", "Vendas"
    ]
    # Map by fuzzy match
    for c in list(df.columns):
        cn = _norm(c)
        if "impresso" in cn:
            df["Impressoes"] = pd.to_numeric(df[c], errors="coerce")
        elif cn == "cliques":
            df["Cliques"] = pd.to_numeric(df[c], errors="coerce")
        elif "receita" in cn:
            df["Receita"] = pd.to_numeric(df[c], errors="coerce")
        elif "investimento" in cn:
            df["Investimento"] = pd.to_numeric(df[c], errors="coerce")
        elif "vendas" in cn and "public" in cn:
            df["Vendas"] = pd.to_numeric(df[c], errors="coerce")

    # Keep minimum
    keep = ["ID"]
    for c in ["Impressoes", "Cliques", "Receita", "Investimento", "Vendas"]:
        if c in df.columns:
            keep.append(c)
    return df[keep].copy()


def _coerce_campaign_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        cn = _norm(c)
        if any(k in cn for k in ["impresso", "clique", "receita", "investimento", "vendas", "roas", "cvr", "perdidas", "orcamento", "acos"]):
            df[c] = pd.to_numeric(df[c], errors="ignore")
    return df


def load_campanhas_diario(campanhas_file) -> pd.DataFrame:
    xls = pd.ExcelFile(campanhas_file)
    sheet = _find_sheet(xls, ["campanha"])
    if sheet is None:
        sheet = xls.sheet_names[0]
    df = pd.read_excel(campanhas_file, sheet_name=sheet, header=1)

    # date col
    if "Desde" in df.columns:
        df["Desde"] = pd.to_datetime(df["Desde"], errors="coerce")
    else:
        # try find date column
        for c in df.columns:
            if "desde" in _norm(c):
                df["Desde"] = pd.to_datetime(df[c], errors="coerce")
                break

    df = _coerce_campaign_numeric(df)
    return df


def load_campanhas_consolidado(campanhas_file) -> pd.DataFrame:
    xls = pd.ExcelFile(campanhas_file)
    sheet = _find_sheet(xls, ["campanha"])
    if sheet is None:
        sheet = xls.sheet_names[0]
    df = pd.read_excel(campanhas_file, sheet_name=sheet, header=1)
    df = _coerce_campaign_numeric(df)
    return df


def build_daily_from_diario(camp_diario: pd.DataFrame) -> pd.DataFrame:
    # Try to locate column names by fuzzy match
    def col_like(keywords):
        for c in camp_diario.columns:
            cn = _norm(c)
            ok = True
            for kw in keywords:
                if _norm(kw) not in cn:
                    ok = False
                    break
            if ok:
                return c
        return None

    col_inv = col_like(["investimento"])
    col_rec = col_like(["receita"])
    col_vend = col_like(["vendas"])
    col_cli = col_like(["cliques"])
    col_imp = col_like(["impresso"])

    daily = camp_diario.groupby("Desde", as_index=False).agg(
        Investimento=(col_inv, "sum") if col_inv else ("Desde", "size"),
        Receita=(col_rec, "sum") if col_rec else ("Desde", "size"),
        Vendas=(col_vend, "sum") if col_vend else ("Desde", "size"),
        Cliques=(col_cli, "sum") if col_cli else ("Desde", "size"),
        Impressoes=(col_imp, "sum") if col_imp else ("Desde", "size"),
    )
    return daily.sort_values("Desde")


def build_campaign_agg(camp: pd.DataFrame, modo: str) -> pd.DataFrame:
    # Fuzzy column getter
    def col_like(keywords):
        for c in camp.columns:
            cn = _norm(c)
            ok = True
            for kw in keywords:
                if _norm(kw) not in cn:
                    ok = False
                    break
            if ok:
                return c
        return None

    col_nome = col_like(["nome"]) or "Nome"
    col_status = col_like(["status"])
    col_orc = col_like(["orcamento"]) or col_like(["orçamento"])
    col_acos = col_like(["acos", "objetivo"])
    col_imp = col_like(["impresso"])
    col_cli = col_like(["cliques"])
    col_rec = col_like(["receita"])
    col_inv = col_like(["investimento"])
    col_vend = col_like(["vendas", "public"])
    col_roas = col_like(["roas"])
    col_cvr = col_like(["cvr"])
    col_lost_b = col_like(["perdidas", "orcamento"]) or col_like(["perdidas", "orçamento"])
    col_lost_r = col_like(["perdidas", "classificacao"]) or col_like(["perdidas", "classificação"])

    if modo == "diario":
        g = camp.groupby(col_nome, as_index=False).agg(
            Nome=(col_nome, "first"),
            Status=(col_status, "last") if col_status else (col_nome, "first"),
            Orcamento=(col_orc, "last") if col_orc else (col_nome, "first"),
            Acos_Objetivo=(col_acos, "last") if col_acos else (col_nome, "first"),
            Impressoes=(col_imp, "sum") if col_imp else (col_nome, "size"),
            Cliques=(col_cli, "sum") if col_cli else (col_nome, "size"),
            Receita=(col_rec, "sum") if col_rec else (col_nome, "size"),
            Investimento=(col_inv, "sum") if col_inv else (col_nome, "size"),
            Vendas=(col_vend, "sum") if col_vend else (col_nome, "size"),
            ROAS=(col_roas, "mean") if col_roas else (col_nome, "size"),
            CVR=(col_cvr, "mean") if col_cvr else (col_nome, "size"),
            Perdidas_Orc=(col_lost_b, "mean") if col_lost_b else (col_nome, "size"),
            Perdidas_Class=(col_lost_r, "mean") if col_lost_r else (col_nome, "size"),
        )
        return g[[
            "Nome","Status","Orcamento","Acos_Objetivo","Impressoes","Cliques",
            "Receita","Investimento","Vendas","ROAS","CVR","Perdidas_Orc","Perdidas_Class"
        ]].copy()

    # Consolidado: assume 1 linha por campanha
    df = pd.DataFrame()
    df["Nome"] = camp[col_nome].astype(str)
    df["Status"] = camp[col_status] if col_status else pd.NA
    df["Orcamento"] = camp[col_orc] if col_orc else pd.NA
    df["Acos_Objetivo"] = camp[col_acos] if col_acos else pd.NA
    df["Impressoes"] = camp[col_imp] if col_imp else pd.NA
    df["Cliques"] = camp[col_cli] if col_cli else pd.NA
    df["Receita"] = camp[col_rec] if col_rec else pd.NA
    df["Investimento"] = camp[col_inv] if col_inv else pd.NA
    df["Vendas"] = camp[col_vend] if col_vend else pd.NA
    df["ROAS"] = camp[col_roas] if col_roas else pd.NA
    df["CVR"] = camp[col_cvr] if col_cvr else pd.NA
    df["Perdidas_Orc"] = camp[col_lost_b] if col_lost_b else pd.NA
    df["Perdidas_Class"] = camp[col_lost_r] if col_lost_r else pd.NA

    for c in ["Impressoes","Cliques","Receita","Investimento","Vendas","ROAS","CVR","Perdidas_Orc","Perdidas_Class","Orcamento","Acos_Objetivo"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def build_tables(
    org: pd.DataFrame,
    camp_agg: pd.DataFrame,
    pat: pd.DataFrame,
    enter_visitas_min: int = 50,
    enter_conv_min: float = 0.05,
    pause_invest_min: float = 100.0,
    pause_cvr_max: float = 0.01,
    scale_lost_budget_min: float = 20.0,
    scale_cvr_min: float = 0.02,
    scale_roas_min: float = 6.0,
    acos_lost_rank_min: float = 30.0,
    acos_roas_min: float = 7.0
):
    # Normalize numeric
    for c in ["Investimento","Receita","Vendas","ROAS","CVR","Perdidas_Orc","Perdidas_Class"]:
        if c in camp_agg.columns:
            camp_agg[c] = pd.to_numeric(camp_agg[c], errors="coerce")

    # Pause
    pause = camp_agg[
        (camp_agg["Investimento"] > pause_invest_min) &
        ((camp_agg["Vendas"] <= 0) | (camp_agg["CVR"] < pause_cvr_max))
    ].copy()
    pause["Acao"] = "PAUSAR"
    pause = pause.sort_values("Investimento", ascending=False)

    # Enter ads
    ads_ids = set(pat["ID"].dropna().astype(str).unique()) if "ID" in pat.columns else set()
    if "ID" in org.columns:
        org_ids = org["ID"].astype(str)
    else:
        org_ids = pd.Series([], dtype=str)

    enter = org.copy()
    if "Visitas" in enter.columns:
        enter = enter[enter["Visitas"] >= enter_visitas_min]
    if "Conv_Visitas_Vendas" in enter.columns:
        enter = enter[enter["Conv_Visitas_Vendas"] > enter_conv_min]
    if "ID" in enter.columns:
        enter = enter[~enter["ID"].astype(str).isin(ads_ids)]

    if "ID" in enter.columns:
        enter["Codigo_MLB"] = "MLB" + enter["ID"].astype(str)

    # Scale budget
    scale = camp_agg[
        (camp_agg["Perdidas_Orc"] > scale_lost_budget_min) &
        (camp_agg["CVR"] >= scale_cvr_min) &
        (camp_agg["ROAS"] >= scale_roas_min)
    ].copy()
    scale["Acao"] = "AUMENTAR_ORCAMENTO"
    scale = scale.sort_values("Perdidas_Orc", ascending=False)

    # Increase acos
    acos = camp_agg[
        (camp_agg["Perdidas_Class"] > acos_lost_rank_min) &
        (camp_agg["ROAS"] >= acos_roas_min)
    ].copy()
    acos["Acao"] = "AUMENTAR_ACOS"
    acos = acos.sort_values("Perdidas_Class", ascending=False)

    invest_total = float(pd.to_numeric(camp_agg["Investimento"], errors="coerce").fillna(0).sum())
    receita_total = float(pd.to_numeric(camp_agg["Receita"], errors="coerce").fillna(0).sum())
    vendas_total = int(pd.to_numeric(camp_agg["Vendas"], errors="coerce").fillna(0).sum())
    roas_total = (receita_total / invest_total) if invest_total else 0.0

    kpis = {
        "campaigns_unique": int(camp_agg["Nome"].nunique()) if "Nome" in camp_agg.columns else 0,
        "sponsored_ids_unique": int(pat["ID"].nunique()) if "ID" in pat.columns else 0,
        "investment": invest_total,
        "revenue": receita_total,
        "sales": vendas_total,
        "roas": roas_total,
    }

    return kpis, pause, enter, scale, acos


def gerar_excel(kpis, camp_agg, pause, enter, scale, acos) -> bytes:
    resumo = pd.DataFrame([kpis])
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo.to_excel(writer, index=False, sheet_name="RESUMO")
        pause.to_excel(writer, index=False, sheet_name="PAUSAR_CAMPANHAS")
        enter.to_excel(writer, index=False, sheet_name="ENTRAR_EM_ADS")
        scale.to_excel(writer, index=False, sheet_name="ESCALAR_ORCAMENTO")
        acos.to_excel(writer, index=False, sheet_name="AJUSTAR_ACOS")
        camp_agg.to_excel(writer, index=False, sheet_name="BASE_CAMPANHAS")
    out.seek(0)
    return out.read()
