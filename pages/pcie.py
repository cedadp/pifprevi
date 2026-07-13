# ============================================================
# app.py  —  Consolidation multi-sources CDG vers CSV unifié
# ============================================================
# Dépendances : streamlit  pandas  openpyxl  xlrd
# Lancer     : streamlit run app.py

import streamlit as st
import pandas as pd
import re
from datetime import datetime

# ── Constantes ─────────────────────────────────────────────────────────────
OUTPUT_COLS = [
    "ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
    "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT",
]

# ── Helpers partagés ────────────────────────────────────────────────────────

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        c.strip().lower() if isinstance(c, str) else str(c)
        for c in df.columns
    ]
    return df


def parse_date(val):
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d.%m.%y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return s


def finalize_output(dfs: list) -> pd.DataFrame:
    if not dfs:
        return pd.DataFrame(columns=OUTPUT_COLS)
    out = pd.concat(dfs, ignore_index=True)
    out = out[OUTPUT_COLS]
    out = out[out["NbPaxTOT"] != 0].reset_index(drop=True)
    return out


# ── Transformations par source ───────────────────────────────────────────────

def transform_af(uploaded) -> pd.DataFrame:
    """
    Air France — Programme brut (onglets multiples).
    Cherche l'en-tête "cieope" automatiquement dans les lignes,
    filtre CieOpe == "AF", NumVol entier.
    """
    xl = pd.ExcelFile(uploaded)
    for sheet in xl.sheet_names:
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        for i, row in df_raw.iterrows():
            if any(str(v).strip().lower() == "cieope" for v in row):
                header_row = i
                df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
                break
        else:
            continue
        break

    df = normalize_columns(df)
    col_map = {}
    for col in df.columns:
        cl = col.strip().lower() if isinstance(col, str) else str(col)
        if cl == "cieope":
            col_map[col] = "cieope"
        elif cl in ("numvol", "flight number", "fltnbr"):
            col_map[col] = "numvol"
        elif cl in ("pax tot", "ttl"):
            col_map[col] = "nbpax_tot"
        elif cl == "arrdep":
            col_map[col] = "arrdep"
        elif cl == "escd ep":
            col_map[col] = "escd ep"
        elif cl == "escar r":
            col_map[col] = "escar r"
        elif cl == "datelocalemvt":
            col_map[col] = "datelocalemvt"
    df = df.rename(columns=col_map)

    df = df[df["cieope"].astype(str).str.strip() == "AF"].copy()
    df["numvol"] = pd.to_numeric(df["numvol"], errors="coerce").fillna(0).astype(int)
    df["datelocalemvt"] = df["datelocalemvt"].apply(parse_date)
    df["nbpax_tot"] = pd.to_numeric(df["nbpax_tot"], errors="coerce").fillna(0).astype(int)

    out = pd.DataFrame({
        "ArrDep":        df.get("arrdep",   "D"),
        "CieOpe":        df["cieope"],
        "NumVol":        df["numvol"].astype(int),
        "EscDep":        df.get("escd ep", "CDG").astype(str).str.strip().str.upper(),
        "EscArr":        df.get("escar r", "CDG").astype(str).str.strip().str.upper(),
        "DateLocaleMvt": df["datelocalemvt"],
        "NbPaxCNT":      0,
        "NbPaxTOT":      df["nbpax_tot"].astype(int),
    })
    return out


def transform_lh(inbound_uploaded, outbound_uploaded) -> pd.DataFrame:
    """
    LH — Lufthansa.  inbound + outbound, formats différents.
    inbound  : "LHG inbound 90 days 06JUL.xlsx" — col A dates fusionnées, ffill.
    outbound : "LHG outbound 90 jours 2.xlsx"   — en-tête row 0, Estimated PAX.
    """
    dfs_out = []

    # ── inbound ────────────────────────────────────────────────────────────────
    xl_in = pd.ExcelFile(inbound_uploaded)
    df_in = pd.read_excel(xl_in, sheet_name=0, header=0)
    df_in = normalize_columns(df_in)

    pax_col = [c for c in df_in.columns if "estimated pax" in c][0]
    dt_col  = [c for c in df_in.columns if c.strip().lower() in ("arr date", "date")][0]
    org_col = [c for c in df_in.columns if c.strip().lower() == "origin"][0]
    dst_col = [c for c in df_in.columns if c.strip().lower() == "dest"][0]
    flt_col = [c for c in df_in.columns if "flt" in c][0]

    df_in[pax_col] = pd.to_numeric(df_in[pax_col], errors="coerce").fillna(0)
    df_in[dt_col]  = df_in[dt_col].ffill()
    df_in["_date"] = df_in[dt_col].apply(parse_date)
    df_in["_cie"]  = df_in[flt_col].astype(str).str.strip().str.upper().str[:2]
    df_in["_nv"]   = pd.to_numeric(
        df_in[flt_col].astype(str).str.strip().str.replace(r"^[A-Z]{2}", "", regex=True),
        errors="coerce"
    ).fillna(0).astype(int)

    d_in = pd.DataFrame({
        "ArrDep":        "A",
        "CieOpe":        df_in["_cie"],
        "NumVol":        df_in["_nv"],
        "EscDep":        df_in[dst_col].astype(str).str.strip().str.upper(),
        "EscArr":        df_in[org_col].astype(str).str.strip().str.upper(),
        "DateLocaleMvt": df_in["_date"],
        "NbPaxCNT":      0,
        "NbPaxTOT":      df_in[pax_col].astype(int),
    })
    dfs_out.append(d_in)

    # ── outbound ──────────────────────────────────────────────────────────────
    xl_out = pd.ExcelFile(outbound_uploaded)
    df_out = pd.read_excel(xl_out, sheet_name=0, header=0)
    df_out = normalize_columns(df_out)

    pax_col_o = [c for c in df_out.columns if "estimated pax" in c][0]
    dt_col_o  = [c for c in df_out.columns if c.strip().lower() in ("dep date", "date")][0]
    org_col_o = [c for c in df_out.columns if c.strip().lower() == "origin"][0]
    dst_col_o = [c for c in df_out.columns if c.strip().lower() == "dest"][0]
    flt_col_o = [c for c in df_out.columns if "flt" in c][0]

    # Exclure lignes à 0 pax ( footer à zéros )
    df_out = df_out[df_out[pax_col_o].apply(
        lambda x: pd.notna(x) and str(x).strip() not in ("", "0")
    )]
    df_out[pax_col_o] = pd.to_numeric(df_out[pax_col_o], errors="coerce").fillna(0)
    df_out["_date"]   = df_out[dt_col_o].apply(parse_date)
    df_out["_cie"]    = df_out[flt_col_o].astype(str).str.strip().str.upper().str[:2]
    df_out["_nv"]     = pd.to_numeric(
        df_out[flt_col_o].astype(str).str.strip().str.replace(r"^[A-Z]{2}", "", regex=True),
        errors="coerce"
    ).fillna(0).astype(int)

    d_out = pd.DataFrame({
        "ArrDep":        "D",
        "CieOpe":        df_out["_cie"],
        "NumVol":        df_out["_nv"],
        "EscDep":        df_out[org_col_o].astype(str).str.strip().str.upper(),
        "EscArr":        df_out[dst_col_o].astype(str).str.strip().str.upper(),
        "DateLocaleMvt": df_out["_date"],
        "NbPaxCNT":      0,
        "NbPaxTOT":      df_out[pax_col_o].astype(int),
    })
    dfs_out.append(d_out)

    return pd.concat(dfs_out, ignore_index=True)


def transform_ez(uploaded) -> pd.DataFrame:
    """
    EZ — easyJet (.xls via xlrd).
    Lecture brute (header=None), en-tête détectée en row 5 :
    [DATE, FLT, TYPE, nan, REG, AC, DEP, ARR, STD, STA, ETD, nan, EXP]
    Données dès row 7.  Footer ("Total Record(s): ...") starts row 593+ —
    détecté par FLT non-matching et expurity de EXP numérique.
    """
    df_raw = pd.read_excel(uploaded, sheet_name="Sheet", header=None, engine="xlrd")
    # Nommer colonnes avec row 5
    header = df_raw.iloc[5]
    df = df_raw.iloc[7:].copy()
    df.columns = header
    df = df.dropna(how="all")

    # Garder uniquement les lignes avec FLT non-vide et EXP numérique
    def valid_row(r):
        flt = r.get("FLT")
        exp = r.get("EXP")
        if pd.isna(flt):
            return False
        flt_s = str(flt).strip()
        # Rejeter le footer : pas de flight number réel
        if re.match(r"^[A-Z]{2}\d", flt_s.upper()) is None and not flt_s.isdigit():
            return False
        if pd.isna(exp):
            return False
        exp_s = str(exp).strip()
        if not (exp_s.replace('.', '', 1).isdigit()):
            return False
        return True

    df = df[df.apply(valid_row, axis=1)]

    def eju_or_ezy(flt):
        s = str(flt).strip().upper()
        m = re.match(r"^(EJU)(\d+)$", s)
        if m:
            return ("EJU", int(m.group(2)))
        num = re.sub(r"^[A-Z]{2,4}", "", s)
        return ("EZY", int(num) if num.isdigit() else 0)

    df[["CieOpe", "NumVol"]] = df["FLT"].apply(
        lambda x: pd.Series(eju_or_ezy(x))
    )
    df["ArrDep"] = df.apply(
        lambda r: "A" if str(r.get("ARR", "")).strip().upper() == "CDG"
        else ("D" if str(r.get("DEP", "")).strip().upper() == "CDG" else ""),
        axis=1,
    )
    df["EscDep"] = df["DEP"].astype(str).str.strip().str.upper()
    df["EscArr"] = df["ARR"].astype(str).str.strip().str.upper()
    df["DateLocaleMvt"] = df["DATE"].apply(parse_date)
    df["NbPaxTOT"] = pd.to_numeric(df["EXP"], errors="coerce").fillna(0).astype(int)
    df["NbPaxCNT"] = 0

    return pd.DataFrame({
        "ArrDep":        df["ArrDep"],
        "CieOpe":        df["CieOpe"],
        "NumVol":        df["NumVol"].astype(int),
        "EscDep":        df["EscDep"],
        "EscArr":        df["EscArr"],
        "DateLocaleMvt": df["DateLocaleMvt"],
        "NbPaxCNT":      df["NbPaxCNT"].astype(int),
        "NbPaxTOT":      df["NbPaxTOT"].astype(int),
    })


def transform_ai(uploaded) -> pd.DataFrame:
    """
    AI — Air India.
    Onglet "Masque Prévisions CDG".  Format déjà cible (11 colonnes).
    Mapping direct des 8 colonnes OUTPUT_COLS.
    """
    df = pd.read_excel(uploaded, sheet_name="Masque Prévisions CDG")
    df = normalize_columns(df)

    out = pd.DataFrame({
        "ArrDep":        df["arrdep"].astype(str).str.strip().str.upper(),
        "CieOpe":        df["cieope"].astype(str).str.strip().str.upper(),
        "NumVol":        pd.to_numeric(df["numvol"], errors="coerce").fillna(0).astype(int),
        "EscDep":        df["escd ep"].astype(str).str.strip().str.upper(),
        "EscArr":        df["escar r"].astype(str).str.strip().str.upper(),
        "DateLocaleMvt": df["datelocalemvt"].apply(parse_date),
        "NbPaxCNT":      pd.to_numeric(df.get("nbpaxcnt", 0), errors="coerce").fillna(0).astype(int),
        "NbPaxTOT":      pd.to_numeric(df["nbpax_tot"], errors="coerce").fillna(0).astype(int),
    })
    return out


def transform_ei(uploaded) -> pd.DataFrame:
    """
    EI — Aer Lingus (fichier 202678 Masque Prévisions CDG.xlsx).
    Onglet "Masque Prévisions CDG".
    Format quasi-cible : 7 colonnes (pas de NbPaxCNT → 0).
    Colonnes : ArrDep, CieOpe(=EI), NumVol, EscDep, EscArr, DateLocaleMvt, NbPaxTOT.
    """
    df = pd.read_excel(uploaded, sheet_name="Masque Prévisions CDG")
    df = normalize_columns(df)

    out = pd.DataFrame({
        "ArrDep":        df["arrdep"].astype(str).str.strip().str.upper(),
        "CieOpe":        df["cieope"].astype(str).str.strip().str.upper(),
        "NumVol":        pd.to_numeric(df["numvol"], errors="coerce").fillna(0).astype(int),
        "EscDep":        df["escd ep"].astype(str).str.strip().str.upper(),
        "EscArr":        df["escar r"].astype(str).str.strip().str.upper(),
        "DateLocaleMvt": df["datelocalemvt"].apply(parse_date),
        "NbPaxCNT":      0,
        "NbPaxTOT":      pd.to_numeric(df["nbpax_tot"], errors="coerce").fillna(0).astype(int),
    })
    return out


# ── Interface Streamlit ──────────────────────────────────────────────────────

st.set_page_config(page_title="CDG — Consolidateur multi-sources", layout="centered")
st.title("✈  Consolidateur CDG")

st.markdown(
    "**Dépendances à installer :** `pip install streamlit pandas openpyxl xlrd`"
)

uploaded = {}
with st.expander("📁 Charger les fichiers", expanded=True):
    # ── AF ──────────────────────────────────────────────────────────────────
    st.markdown("**AF — Air France**")
    uploaded["AF"] = st.file_uploader(
        "  Air France (Programme brut, .xlsx)", type=["xlsx"], key="file_af"
    )

    # ── LH ─────────────────────────────────────────────────────────────────
    st.markdown("**LH — Lufthansa**")
    uploaded["LH_inbound"]  = st.file_uploader(
        "  inbound  (.xlsx)", type=["xlsx"], key="file_lh_inbound"
    )
    uploaded["LH_outbound"] = st.file_uploader(
        "  outbound (.xlsx)", type=["xlsx"], key="file_lh_outbound"
    )

    # ── EZ ─────────────────────────────────────────────────────────────────
    st.markdown("**EZ — easyJet**")
    uploaded["EZ"] = st.file_uploader(
        "  easyJet (Job 80C, .xls)", type=["xls", "xlsx"], key="file_ez"
    )

    # ── AI ─────────────────────────────────────────────────────────────────
    st.markdown("**AI — Air India**")
    uploaded["AI"] = st.file_uploader(
        "  Air India (Masque Prévisions CDG, .xlsx)", type=["xlsx"], key="file_ai"
    )

    # ── EI ─────────────────────────────────────────────────────────────────
    st.markdown("**EI — Aer Lingus**")
    uploaded["EI"] = st.file_uploader(
        "  Aer Lingus (202678 Masque Prévisions CDG, .xlsx)", type=["xlsx"], key="file_ei"
    )


if st.button("▶  GO", type="primary"):
    errors = []
    dfs    = []

    # AF
    if uploaded.get("AF") is not None:
        try:
            df = transform_af(uploaded["AF"])
            dfs.append(df)
            st.success(f"AF  : {len(df)} lignes")
        except Exception as e:
            errors.append(f"AF : {e}")

    # LH
    ib = uploaded.get("LH_inbound")
    ob = uploaded.get("LH_outbound")
    if ib is not None or ob is not None:
        try:
            df = transform_lh(ib, ob)
            dfs.append(df)
            st.success(f"LH  : {len(df)} lignes")
        except Exception as e:
            errors.append(f"LH : {e}")

    # EZ
    if uploaded.get("EZ") is not None:
        try:
            df = transform_ez(uploaded["EZ"])
            dfs.append(df)
            st.success(f"EZ  : {len(df)} lignes")
        except Exception as e:
            errors.append(f"EZ : {e}")

    # AI
    if uploaded.get("AI") is not None:
        try:
            df = transform_ai(uploaded["AI"])
            dfs.append(df)
            st.success(f"AI  : {len(df)} lignes")
        except Exception as e:
            errors.append(f"AI : {e}")

    # EI
    if uploaded.get("EI") is not None:
        try:
            df = transform_ei(uploaded["EI"])
            dfs.append(df)
            st.success(f"EI  : {len(df)} lignes")
        except Exception as e:
            errors.append(f"EI : {e}")

    if errors:
        st.error("Erreurs : " + ";  ".join(errors))
        st.stop()

    if not dfs:
        st.warning("Aucun fichier chargé.")
        st.stop()

    result = finalize_output(dfs)

    st.dataframe(result.head(50), use_container_width=True)
    st.caption(f"{len(result)} lignes consolidées")

    csv_bytes = result.to_csv(
        sep=";", encoding="utf-8", index=False, lineterminator="\n"
    ).encode("utf-8")

    st.download_button(
        "⬇  Télécharger le CSV consolidé",
        data=csv_bytes,
        file_name="consolidation_cdg.csv",
        mime="text/csv",
    )
