import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Concaténateur Prévisions Cies", layout="wide")

# ---------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------
OUTPUT_COLS = ["ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
               "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT"]


# ---------------------------------------------------------------
# UTILITAIRES COMMUNS
# ---------------------------------------------------------------
def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def parse_lh_date(serie):
    """Dates au format D.M.YY (ex '6.7.26') -> JJ/MM/AAAA."""
    d = pd.to_datetime(serie, format="%d.%m.%y", errors="coerce")
    d = d.fillna(pd.to_datetime(serie, errors="coerce", dayfirst=True))
    return d.dt.strftime("%d/%m/%Y")


def split_flight(serie):
    """'SN3631' -> ('SN', '3631'). Compagnie = 2 premiers caractères."""
    s = serie.astype(str).str.replace(r"\s+", "", regex=True).str.upper()
    cie = s.str.extract(r"^([A-Z0-9]{2})")[0]
    num = s.str.extract(r"^[A-Z0-9]{2}(\d+)")[0]
    return cie, num


def finalize_output(out):
    """Nettoyage commun + exclusion 0 pax (même critère que AF)."""
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()
    out = out[out["NbPaxTOT"] != 0]                       # exclusion 0 pax
    out = out[out["CieOpe"].notna() & (out["CieOpe"] != "")]
    return out[OUTPUT_COLS]


# ---------------------------------------------------------------
# TRANSFORMATION AF (mapping générique)
# ---------------------------------------------------------------
def read_excel_source(file, conf):
    return pd.read_excel(file, sheet_name=conf["sheet"])


def read_paste_source(txt):
    return pd.read_csv(io.StringIO(txt), sep=None, engine="python")


def transform(df, conf, label=""):
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

    # filtre compagnie
    if conf.get("filter_cie"):
        out = out[out["CieOpe"].astype(str).str.strip() == conf["filter_cie"]]

    # date au format JJ/MM/AAAA
    dcol = conf.get("date_col")
    if dcol and dcol in out.columns:
        d = pd.to_datetime(out[dcol], errors="coerce", dayfirst=True)
        out[dcol] = d.dt.strftime("%d/%m/%Y")

    # pax numériques
    for c in ["NbPaxCNT", "NbPaxTOT"]:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0).astype(int)

    # exclusion 0 pax
    if conf.get("exclude_zero_pax"):
        out = out[out["NbPaxTOT"] != 0]

    # nettoyage num vol
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip()

    return out[OUTPUT_COLS]


# ---------------------------------------------------------------
# TRANSFORMATIONS LH (parseurs dédiés)
# ---------------------------------------------------------------
def transform_lh_inbound(file):
    """Arrivées LHG : Arr Date fusionnée (ffill), EscArr = CDG."""
    df = pd.read_excel(file, sheet_name="Sheet 1", header=0)
    df = normalize_columns(df)

    df["Arr Date"] = df["Arr Date"].ffill()
    df = df[df["Flt Nbr"].notna() & (df["Flt Nbr"].astype(str).str.strip() != "")]

    cie, num = split_flight(df["Flt Nbr"])
    out = pd.DataFrame({
        "ArrDep": "A",
        "CieOpe": cie.values,
        "NumVol": num.values,
        "EscDep": df["Origin"].astype(str).str.strip().values,
        "EscArr": "CDG",
        "DateLocaleMvt": parse_lh_date(df["Arr Date"]).values,
        "NbPaxCNT": 0,
        "NbPaxTOT": pd.to_numeric(df["Estimated PAX"], errors="coerce").fillna(0).astype(int).values,
    })
    return out


def transform_lh_outbound(file):
    """Départs LHG : onglet 'Input', EscDep = CDG."""
    df = pd.read_excel(file, sheet_name="Input", header=0)
    df = normalize_columns(df)

    df = df[df["Flt Nbr"].notna() & (df["Flt Nbr"].astype(str).str.strip() != "")]

    cie, num = split_flight(df["Flt Nbr"])
    out = pd.DataFrame({
        "ArrDep": "D",
        "CieOpe": cie.values,
        "NumVol": num.values,
        "EscDep": "CDG",
        "EscArr": df["Dest"].astype(str).str.strip().values,
        "DateLocaleMvt": parse_lh_date(df["Dep Date"]).values,
        "NbPaxCNT": 0,
        "NbPaxTOT": pd.to_numeric(df["Estimated PAX"], errors="coerce").fillna(0).astype(int).values,
    })
    return out


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
        "exclude_zero_pax": True,
    },
    "LH_IN": {
        "input_type": "excel",
        "label": "LH — Arrivées (inbound)",
        "custom": transform_lh_inbound,
    },
    "LH_OUT": {
        "input_type": "excel",
        "label": "LH — Départs (outbound)",
        "custom": transform_lh_outbound,
    },
}


# ---------------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------------
st.title("✈️ Concaténateur de prévisions compagnies")
st.markdown(
    "Déposez les fichiers Excel et/ou collez les données, puis cliquez sur **GO**."
)

excel_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "excel"}
paste_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "paste"}

uploaded = {}
pasted = {}

st.header("📁 Fichiers Excel")
for name, conf in excel_sources.items():
    label = conf.get("label", f"Fichier {name}")
    uploaded[name] = st.file_uploader(label, type=["xlsx", "xls"], key=f"file_{name}")

if paste_sources:
    st.header("📋 Données à coller")
    for name, conf in paste_sources.items():
        label = conf.get("label", f"Données {name}")
        pasted[name] = st.text_area(label, height=150, key=f"paste_{name}")

st.divider()

# ---------------------------------------------------------------
# TRAITEMENT
# ---------------------------------------------------------------
if st.button("🚀 GO", type="primary", use_container_width=True):
    frames = []

    # Sources Excel
    for name, conf in excel_sources.items():
        up = uploaded.get(name)
        if up is None:
            continue
        try:
            if "custom" in conf:                       # sources LH
                res = finalize_output(conf["custom"](up))
            else:                                       # sources type AF (mapping)
                df_in = read_excel_source(up, conf)
                res = transform(df_in, conf, label=name)
            if res is not None and not res.empty:
                frames.append(res)
                st.success(f"[{name}] {len(res)} lignes intégrées.")
        except Exception as e:
            st.error(f"[{name}] Erreur de lecture : {e}")

    # Sources à coller
    for name, conf in paste_sources.items():
        txt = pasted.get(name, "")
        if txt and txt.strip():
            try:
                df_in = read_paste_source(txt)
                res = transform(df_in, conf, label=name)
                if res is not None and not res.empty:
                    frames.append(res)
                    st.success(f"[{name}] {len(res)} lignes intégrées.")
            except Exception as e:
                st.error(f"[{name}] Erreur de lecture : {e}")

    # Résultat final
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
