"""
Courbes de présentation aéroportuaires — fichier unique.

Lancement interface :   streamlit run app.py
Lancement en batch  :   python app.py chemin/vers/presentation_BD.xlsx
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------- Paramètres

BUCKETS = list(range(0, 300, 10))                      # 0 à 290 par pas de 10
BUCKET_LABELS = [f"{x:03d}-{x + 10:03d} min" for x in BUCKETS]
LABEL_BY_BUCKET = dict(zip(BUCKETS, BUCKET_LABELS))
FENETRE_MAX_MIN = 300                                  # 5 h avant le départ

COL_JOUR = "Jour"
COL_PAX = "nb pax"
COL_VOL = "Numéro de vol"
COL_HORAIRE = "Horaire théorique"
COL_TRANCHE = "Tranche 10 minutes passage"
COL_GEO = "Faisceau géographique"
COL_INTER = "Faisceau inter/sch"
COL_CIE = "Code IATA compagnie"

REQUIRED_COLUMNS = [COL_JOUR, COL_PAX, COL_VOL, COL_HORAIRE,
                    COL_TRANCHE, COL_GEO, COL_INTER, COL_CIE]

CLES_VOL = [COL_JOUR, COL_VOL]
ATTRS_VOL = [COL_GEO, COL_INTER, COL_CIE]

# Nom de feuille -> colonne de regroupement (None = global)
AGREGATS = {
    "global": None,
    "faisceau_geo": COL_GEO,
    "faisceau_inter_sch": COL_INTER,
    "compagnie": COL_CIE,
}

# ------------------------------------------------------------------ Parsing

def parse_tranche_debut(value) -> float:
    """Début de tranche en minutes depuis minuit. Gère 'HH.MM.SS - HH.MM.SS' et 'HH:MM'."""
    m = re.search(r"(\d{1,2})[.:](\d{2})(?:[.:](\d{2}))?", str(value))
    if not m:
        return np.nan
    return int(m.group(1)) * 60 + int(m.group(2)) + int(m.group(3) or 0) / 60


def minutes_horaire(v) -> float:
    """Heure de départ en minutes depuis minuit."""
    if pd.isna(v):
        return np.nan
    if isinstance(v, pd.Timestamp):
        return v.hour * 60 + v.minute + v.second / 60
    t = pd.to_datetime(v, errors="coerce")
    return np.nan if pd.isna(t) else t.hour * 60 + t.minute + t.second / 60


# --------------------------------------------------------------- Préparation

def prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    """Calcule le delta avant départ et la tranche de rattachement, ligne à ligne."""
    manquantes = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if manquantes:
        raise ValueError("Colonnes manquantes : " + ", ".join(manquantes))

    x = df[REQUIRED_COLUMNS].copy()
    x[COL_JOUR] = pd.to_datetime(x[COL_JOUR], errors="coerce")
    x[COL_PAX] = pd.to_numeric(x[COL_PAX], errors="coerce").fillna(0)

    debut = x[COL_TRANCHE].map(parse_tranche_debut)
    depart = x[COL_HORAIRE].map(minutes_horaire)

    # Delta positif = présentation avant le départ, avec recalage sur minuit
    delta = depart - debut
    delta = np.where(delta < -720, delta + 1440,
                     np.where(delta > 720, delta - 1440, delta))
    x["delta_min"] = delta

    dans = (x["delta_min"] >= 0) & (x["delta_min"] <= FENETRE_MAX_MIN)
    x["tranche_min"] = pd.Series(pd.NA, index=x.index, dtype="Int64")
    x.loc[dans, "tranche_min"] = (np.floor(x.loc[dans, "delta_min"] / 10) * 10).astype("Int64")
    x.loc[x["tranche_min"] == 300, "tranche_min"] = 290      # borne haute incluse
    x["dans_perimetre"] = x["tranche_min"].notna()
    x["tranche"] = x["tranche_min"].map(LABEL_BY_BUCKET)
    return x


# ------------------------------------------------------------------ Profils

def profiles_by_flight(prepared: pd.DataFrame) -> pd.DataFrame:
    """Profil de chaque vol, renormalisé sur les seuls pax retenus : somme = 100 %."""
    colonnes = CLES_VOL + ATTRS_VOL + ["tranche_min", "tranche",
                                       "passagers_tranche", "passagers_vol", "proportion"]
    v = prepared[prepared["dans_perimetre"]]
    if v.empty:
        return pd.DataFrame(columns=colonnes)

    attrs = v.groupby(CLES_VOL, dropna=False)[ATTRS_VOL].first().reset_index()
    totaux = v.groupby(CLES_VOL, dropna=False)[COL_PAX].sum().rename("passagers_vol").reset_index()
    par_tranche = (v.groupby(CLES_VOL + ["tranche_min"], dropna=False)[COL_PAX]
                    .sum().rename("passagers_tranche").reset_index())

    # Grille complète vols x tranches pour ne pas perdre les tranches vides
    grille = totaux[CLES_VOL].merge(pd.DataFrame({"tranche_min": BUCKETS}), how="cross")
    out = (grille
           .merge(par_tranche, on=CLES_VOL + ["tranche_min"], how="left")
           .merge(totaux, on=CLES_VOL)
           .merge(attrs, on=CLES_VOL))

    out["passagers_tranche"] = out["passagers_tranche"].fillna(0)
    out["proportion"] = np.where(out["passagers_vol"] > 0,
                                 out["passagers_tranche"] / out["passagers_vol"], 0.0)
    out["tranche"] = out["tranche_min"].map(LABEL_BY_BUCKET)
    return out[colonnes].sort_values(CLES_VOL + ["tranche_min"]).reset_index(drop=True)


def aggregate_profiles(profils_vol: pd.DataFrame, group_col: str | None = None) -> pd.DataFrame:
    """Moyenne arithmétique SIMPLE des profils de vols : chaque vol pèse 1."""
    cles = [group_col] if group_col else []
    colonnes = cles + ["tranche_min", "tranche", "nb_vols", "proportion"]
    if profils_vol.empty:
        return pd.DataFrame(columns=colonnes)

    p = profils_vol.copy()
    if group_col:
        p[group_col] = p[group_col].fillna("(non renseigné)").astype(str)

    agg = (p.groupby(cles + ["tranche_min"], dropna=False)
             .agg(proportion=("proportion", "mean"))
             .reset_index())

    nb = (p.groupby(cles, dropna=False)[CLES_VOL]
            .apply(lambda d: d.drop_duplicates().shape[0])
            .rename("nb_vols").reset_index()) if group_col else None

    if group_col:
        agg = agg.merge(nb, on=group_col)
    else:
        agg["nb_vols"] = p[CLES_VOL].drop_duplicates().shape[0]

    agg["tranche"] = agg["tranche_min"].map(LABEL_BY_BUCKET)
    return agg[colonnes].sort_values(cles + ["tranche_min"]).reset_index(drop=True)


def cumulative(df: pd.DataFrame, group_col: str | None = None) -> pd.DataFrame:
    """Part des passagers DÉJÀ présentés à un instant donné.

    Cumul du plus loin du départ (290-300 min) vers le plus proche (0-10 min) :
    0 % à H-5h, 100 % à l'heure de départ.
    """
    if df.empty:
        return df.assign(proportion_cumulee=[])
    out = df.sort_values(([group_col] if group_col else []) + ["tranche_min"],
                         ascending=[True] * bool(group_col) + [False]).copy()
    out["proportion_cumulee"] = (out.groupby(group_col)["proportion"].cumsum()
                                 if group_col else out["proportion"].cumsum())
    return out.sort_values(([group_col] if group_col else []) + ["tranche_min"]).reset_index(drop=True)


def compute_all(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Retourne toutes les tables : préparation, profils par vol et agrégats."""
    prepared = prepare_data(df)
    profils = profiles_by_flight(prepared)
    res = {"prepared": prepared, "profils_par_vol": profils}
    for nom, col in AGREGATS.items():
        res[nom] = aggregate_profiles(profils, col)
    return res


# -------------------------------------------------------------------- Export

def build_workbook(res: dict[str, pd.DataFrame]) -> bytes:
    """Classeur Excel : tables brutes, agrégats et versions cumulées."""
    buf = io.BytesIO()
    ordre = ["prepared", "profils_par_vol"] + list(AGREGATS)
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as w:
        for nom in ordre:
            res[nom].to_excel(w, sheet_name=nom, index=False)
        for nom, col in AGREGATS.items():
            cumulative(res[nom], col).to_excel(w, sheet_name=f"{nom}_cumulee"[:31], index=False)

        pct = w.book.add_format({"num_format": "0.00%"})
        ent = w.book.add_format({"num_format": "0"})
        for nom in ordre + [f"{n}_cumulee"[:31] for n in AGREGATS]:
            base = nom[:-8] if nom.endswith("_cumulee") else nom
            frame = res[base]
            colonnes = list(frame.columns) + (["proportion_cumulee"] if nom.endswith("_cumulee") else [])
            ws = w.sheets[nom]
            ws.freeze_panes(1, 0)
            ws.set_column(0, max(len(colonnes) - 1, 0), 18)
            for i, name in enumerate(colonnes):
                if "proportion" in name:
                    ws.set_column(i, i, 15, pct)
                elif name == "nb_vols":
                    ws.set_column(i, i, 10, ent)
    return buf.getvalue()


# ---------------------------------------------------------------- Streamlit

def main() -> None:
    import streamlit as st
    import plotly.express as px

    st.set_page_config(page_title="Courbes de présentation", layout="wide")
    st.title("Courbes de présentation aéroportuaires")

    up = st.sidebar.file_uploader("Fichier Excel", type=["xlsx", "xls"])
    defaut = Path("presentation_BD.xlsx")
    source = up if up is not None else (defaut if defaut.exists() else None)
    if source is None:
        st.info("Chargez votre fichier Excel dans le panneau de gauche pour démarrer.")
        return

    try:
        df = pd.read_excel(source)
    except Exception as e:
        st.error(f"Lecture impossible : {e}")
        return

    manquantes = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if manquantes:
        st.error("Colonnes manquantes : " + ", ".join(manquantes))
        return

    df[COL_JOUR] = pd.to_datetime(df[COL_JOUR], errors="coerce")

    st.sidebar.header("Filtres")
    jours = sorted(df[COL_JOUR].dropna().dt.strftime("%Y-%m-%d").unique())
    cies = sorted(df[COL_CIE].dropna().astype(str).unique())
    geos = sorted(df[COL_GEO].dropna().astype(str).unique())
    sel_j = st.sidebar.multiselect("Jour", jours, default=jours)
    sel_c = st.sidebar.multiselect("Compagnie", cies, default=cies)
    sel_g = st.sidebar.multiselect("Faisceau géographique", geos, default=geos)

    x = df[df[COL_JOUR].dt.strftime("%Y-%m-%d").isin(sel_j)
           & df[COL_CIE].astype(str).isin(sel_c)
           & df[COL_GEO].astype(str).isin(sel_g)]
    if x.empty:
        st.warning("Aucune ligne ne correspond aux filtres.")
        return

    res = compute_all(x)
    prep = res["prepared"]
    nb_vols = res["profils_par_vol"][CLES_VOL].drop_duplicates().shape[0]
    hors = int((~prep["dans_perimetre"]).sum())
    pax_hors = float(prep.loc[~prep["dans_perimetre"], COL_PAX].sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Lignes sélectionnées", f"{len(x):,}".replace(",", " "))
    c2.metric("Vols retenus", f"{nb_vols:,}".replace(",", " "))
    c3.metric("Lignes hors 0–5 h", f"{hors:,}".replace(",", " "), f"{pax_hors:,.0f} pax exclus".replace(",", " "))
    st.caption("Profil de chaque vol renormalisé à 100 % sur les passagers retenus, "
               "puis moyenne simple des vols (chaque vol pèse 1, sans pondération pax).")

    onglets = st.tabs(["Global", "Faisceau géographique", "Inter / Schengen", "Compagnie"])
    for onglet, (nom, col) in zip(onglets, AGREGATS.items()):
        with onglet:
            a = res[nom]
            cc = cumulative(a, col)
            f1 = px.line(a, x="tranche", y="proportion", color=col, markers=True,
                         title="Courbe de présentation")
            f1.update_yaxes(tickformat=".0%", title="Part des passagers")
            f1.update_xaxes(title="Minutes avant le départ", autorange="reversed")
            st.plotly_chart(f1, use_container_width=True)

            f2 = px.line(cc, x="tranche", y="proportion_cumulee", color=col, markers=True,
                         title="Courbe cumulée")
            f2.update_yaxes(tickformat=".0%", title="Part cumulée")
            f2.update_xaxes(title="Minutes avant le départ", autorange="reversed")
            st.plotly_chart(f2, use_container_width=True)
            st.dataframe(a, use_container_width=True)

    st.download_button("Télécharger l'export Excel", build_workbook(res),
                       "courbes_presentation.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def run_batch(chemin: str) -> None:
    """Génère l'export sans interface : python app.py fichier.xlsx"""
    res = compute_all(pd.read_excel(chemin))
    sortie = Path(chemin).with_name("courbes_presentation.xlsx")
    sortie.write_bytes(build_workbook(res))
    print(f"Export écrit : {sortie}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_batch(sys.argv[1])
    else:
        main()
