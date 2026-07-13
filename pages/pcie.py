import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Concaténateur Prévisions Cies", layout="wide")

OUTPUT_COLS = ["ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
               "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT"]

# ---------------------------------------------------------------
# UTILITAIRES
# ---------------------------------------------------------------
def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df

def parse_lh_date(serie):
    d = pd.to_datetime(serie, format="%d.%m.%y", errors="coerce")
    d = d.fillna(pd.to_datetime(serie, errors="coerce", dayfirst=True))
    return d.dt.strftime("%d/%m/%Y")

def split_flight(serie):
    s = serie.astype(str).str.replace(r"\s+", "", regex=True).str.upper()
    cie = s.str.extract(r"^([A-Z0-9]{2})")[0]
    num = s.str.extract(r"^[A-Z0-9]{2}(\d+)")[0]
    return cie, num

def finalize_output(out):
    out = out.copy()
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()
    out["NbPaxCNT"] = pd.to_numeric(out["NbPaxCNT"], errors="coerce").fillna(0).astype(int)
    out["NbPaxTOT"] = pd.to_numeric(out["NbPaxTOT"], errors="coerce").fillna(0).astype(int)
    out = out[out["NbPaxTOT"] != 0]
    out = out[out["CieOpe"].notna() & (out["CieOpe"] != "") & (out["CieOpe"].str.lower() != "nan")]
    return out[OUTPUT_COLS]

# ---------------------------------------------------------------
# AF (mapping générique)
# ---------------------------------------------------------------
def transform_af(file, conf, label="AF"):
    df = pd.read_excel(file, sheet_name=conf["sheet"])
    df = normalize_columns(df)
    mapping = conf["mapping"]

    missing = [c for c in mapping if c not in df.columns]
    if missing:
        st.error(f"[{label}] Colonnes manquantes : {missing}")
        st.write(f"[{label}] Colonnes trouvées :", list(df.columns))
        return None

    out = pd.DataFrame()
    for src_col, dst_col in mapping.items():
        out[dst_col] = df[src_col]

    if conf.get("filter_cie"):
        out = out[out["CieOpe"].astype(str).str.strip() == conf["filter_cie"]]

    dcol = conf.get("date_col")
    if dcol and dcol in out.columns:
        d = pd.to_datetime(out[dcol], errors="coerce", dayfirst=True)
        out[dcol] = d.dt.strftime("%d/%m/%Y")

    return out

# ---------------------------------------------------------------
# LH
# ---------------------------------------------------------------
def transform_lh_inbound(file):
    df = pd.read_excel(file, sheet_name="Sheet 1", header=0)
    df = normalize_columns(df)
    df["Arr Date"] = df["Arr Date"].ffill()
    df = df[df["Flt Nbr"].notna() & (df["Flt Nbr"].astype(str).str.strip() != "")]
    cie, num = split_flight(df["Flt Nbr"])
    return pd.DataFrame({
        "ArrDep": "A",
        "CieOpe": cie.values,
        "NumVol": num.values,
        "EscDep": df["Origin"].astype(str).str.strip().values,
        "EscArr": "CDG",
        "DateLocaleMvt": parse_lh_date(df["Arr Date"]).values,
        "NbPaxCNT": 0,
        "NbPaxTOT": pd.to_numeric(df["Estimated PAX"], errors="coerce").fillna(0).astype(int).values,
    })

def transform_lh_outbound(file):
    df = pd.read_excel(file, sheet_name="Input", header=0)
    df = normalize_columns(df)
    df = df[df["Flt Nbr"].notna() & (df["Flt Nbr"].astype(str).str.strip() != "")]
    cie, num = split_flight(df["Flt Nbr"])
    return pd.DataFrame({
        "ArrDep": "D",
        "CieOpe": cie.values,
        "NumVol": num.values,
        "EscDep": "CDG",
        "EscArr": df["Dest"].astype(str).str.strip().values,
        "DateLocaleMvt": parse_lh_date(df["Dep Date"]).values,
        "NbPaxCNT": 0,
        "NbPaxTOT": pd.to_numeric(df["Estimated PAX"], errors="coerce").fillna(0).astype(int).values,
    })

# ---------------------------------------------------------------
# AI / EI  (format cible, mapping PAR NOM normalisé)
# ---------------------------------------------------------------
def _target_format(file, sheet_name, cie_fixe=None, cnt_present=True):
    df = pd.read_excel(file, sheet_name=sheet_name, header=0)
    df = normalize_columns(df)

    df = df[df["ArrDep"].astype(str).str.strip().isin(["A", "D"])]

    out = pd.DataFrame({
        "ArrDep": df["ArrDep"].astype(str).str.strip().values,
        "CieOpe": (cie_fixe if cie_fixe else df["CieOpe"].astype(str).str.strip()),
        "NumVol": pd.to_numeric(df["NumVol"], errors="coerce").fillna(0).astype(int).values,
        "EscDep": df["EscDep"].astype(str).str.strip().values,
        "EscArr": df["EscArr"].astype(str).str.strip().values,
        "DateLocaleMvt": pd.to_datetime(df["DateLocaleMvt"], errors="coerce")
                           .dt.strftime("%d/%m/%Y").values,
        "NbPaxCNT": (pd.to_numeric(df["NbPaxCNT"], errors="coerce").fillna(0).astype(int).values
                     if cnt_present else 0),
        "NbPaxTOT": pd.to_numeric(df["NbPaxTOT"], errors="coerce").fillna(0).astype(int).values,
    })
    return out

def transform_ai(file):
    # AI a des colonnes Business/Premium/Economy : le mapping par NOM les ignore
    return _target_format(file, "Masque Prévisions CDG", cie_fixe=None, cnt_present=True)

def transform_ei(file):
    # EI (Aer Lingus) : pas de colonne NbPaxCNT
    return _target_format(file, "Masque Prévisions CDG", cie_fixe="EI", cnt_present=False)

# ---------------------------------------------------------------
# DECLARATION DES SOURCES
# ---------------------------------------------------------------
SOURCES = {
    "AF": {
        "input_type": "excel",
        "label": "AF — Programme brut (Excel)",
        "sheet": "Programme brut",
        "mapping": {
            "A/D": "ArrDep",
            "Cie Ope": "CieOpe",
            "Num Vol": "NumVol",
            "Esc Dep": "EscDep",
            "Esc Arr": "EscArr",
            "Local Date": "DateLocaleMvt",
            "Pax CNT TOT": "NbPaxCNT",
            "PAX TOT": "NbPaxTOT",
        },
        "date_col": "DateLocaleMvt",
        "filter_cie": "AF",
        "af": True,
    },
    "LH_IN":  {"input_type": "excel", "label": "LH — Arrivées (inbound)", "custom": transform_lh_inbound},
    "LH_OUT": {"input_type": "excel", "label": "LH — Départs (outbound)", "custom": transform_lh_outbound},
    "AI":     {"input_type": "excel", "label": "AI — Air India (Masque Prévisions CDG)", "custom": transform_ai},
    "EI":     {"input_type": "excel", "label": "EI — Aer Lingus (Masque Prévisions CDG)", "custom": transform_ei},
}

# ---------------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------------
st.title("✈️ Concaténateur de prévisions compagnies")
st.markdown("Déposez les fichiers Excel puis cliquez sur **GO**.")

excel_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "excel"}
uploaded = {}

st.header("📁 Fichiers Excel")
for name, conf in excel_sources.items():
    uploaded[name] = st.file_uploader(conf.get("label", name),
                                      type=["xlsx", "xls"], key=f"file_{name}")

st.divider()

if st.button("🚀 GO", type="primary", use_container_width=True):
    frames = []
    for name, conf in excel_sources.items():
        up = uploaded.get(name)
        if up is None:
            continue
        try:
            if conf.get("af"):
                res = transform_af(up, conf, label=name)
            else:
                res = conf["custom"](up)
            if res is None:
                continue
            res = finalize_output(res)
            if not res.empty:
                frames.append(res)
                st.success(f"[{name}] {len(res)} lignes intégrées.")
            else:
                st.warning(f"[{name}] 0 ligne après filtrage.")
        except Exception as e:
            st.error(f"[{name}] Erreur : {e}")

    if not frames:
        st.warning("Aucune donnée valide fournie.")
    else:
        final = pd.concat(frames, ignore_index=True)
        st.subheader(f"📦 Fichier final ({len(final)} lignes)")
        st.dataframe(final.head(100))

        header = ";".join(OUTPUT_COLS)
        body = final.to_csv(sep=";", index=False, header=False, lineterminator="\n")
        csv_out = header + "\n" + body

        st.download_button(
            "💾 Télécharger le CSV",
            data=csv_out.encode("utf-8"),
            file_name="Previs_cies.csv",
            mime="text/csv",
        )
