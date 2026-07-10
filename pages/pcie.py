import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Concaténateur Prévisions Cies", layout="wide")

# ---------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------
OUTPUT_COLS = ["ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
               "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT"]

# input_type : "excel" ou "paste"
SOURCES = {
    "AF": {
        "input_type": "excel",
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
    # Exemple de future source à coller (à adapter quand tu auras le format) :
    # "XX": {
    #     "input_type": "paste",
    #     "sheet": None,
    #     "mapping": { ... },
    #     "date_col": "DateLocaleMvt",
    #     "filter_cie": "XX",
    #     "exclude_zero_pax": True,
    # },
}


def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def transform(df, conf, label=""):
    df = normalize_columns(df)
    mapping = conf["mapping"]

    missing = [c for c in mapping if c not in df.columns]
    if missing:
        st.error(f"[{label}] Colonnes manquantes : {missing}")
        st.write(f"[{label}] Colonnes trouvées :", list(df.columns))
        return None

    out = df[list(mapping.keys())].rename(columns=mapping)

    # Date -> JJ/MM/AAAA
    dcol = conf["date_col"]
    out[dcol] = pd.to_datetime(out[dcol], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y")

    # Pax en entiers
    for c in ["NbPaxCNT", "NbPaxTOT"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0).astype(int)

    # Num vol en entier
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)

    # Nettoyage textes
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()

    # --- Filtres ---
    n0 = len(out)

    if conf.get("filter_cie"):
        out = out[out["CieOpe"] == conf["filter_cie"]]

    if conf.get("exclude_zero_pax"):
        out = out[out["NbPaxTOT"] != 0]

    out = out[out["DateLocaleMvt"].notna()]

    st.caption(f"[{label}] {n0} lignes lues → {len(out)} lignes conservées.")

    return out[OUTPUT_COLS].reset_index(drop=True)


def read_excel_source(uploaded_file, conf):
    xls = pd.ExcelFile(uploaded_file)
    sheet = conf["sheet"] if conf["sheet"] in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet)


def read_paste_source(text):
    sep = "\t" if "\t" in text else ";"
    return pd.read_csv(io.StringIO(text), sep=sep)




# ===============================================================
#  AJOUTS POUR LH (inbound + outbound)
# ===============================================================

def parse_lh_date(serie):
    """Dates au format D.M.YY (ex '6.7.26') -> JJ/MM/AAAA."""
    d = pd.to_datetime(serie, format="%d.%m.%y", errors="coerce")
    # secours si certaines cellules sont déjà des datetime Excel
    d = d.fillna(pd.to_datetime(serie, errors="coerce", dayfirst=True))
    return d.dt.strftime("%d/%m/%Y")


def split_flight(serie):
    """'SN3631' -> ('SN', '3631'). Compagnie = 2 premières lettres."""
    s = serie.astype(str).str.replace(r"\s+", "", regex=True).str.upper()
    cie = s.str.extract(r"^([A-Z0-9]{2})")[0]
    num = s.str.extract(r"^[A-Z0-9]{2}(\d+)")[0]
    return cie, num


def transform_lh_inbound(file):
    """Arrivées LHG : Arr Date fusionnée (ffill), EscArr = CDG."""
    df = pd.read_excel(file, sheet_name="Sheet 1", header=0)
    df = normalize_columns(df)

    # ffill sur la date fusionnée
    df["Arr Date"] = df["Arr Date"].ffill()

    # ne garder que les lignes réelles (Flt Nbr renseigné)
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

    # lignes réelles uniquement
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


def finalize_output(out):
    """Nettoyage commun + exclusion 0 pax (même critère que AF)."""
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()
    out = out[out["NbPaxTOT"] != 0]          # exclusion 0 pax
    out = out[out["CieOpe"].notna() & (out["CieOpe"] != "")]
    return out[OUTPUT_COLS]





# ---------------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------------
st.title("🛫 Concaténateur de prévisions Compagnies")

excel_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "excel"}
paste_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "paste"}

uploaded = {}
pasted = {}

# --- Bloc fichiers (glisser / déposer) ---
if excel_sources:
    st.header("📁 Compagnies avec fichier Excel")
    for name in excel_sources:
        uploaded[name] = st.file_uploader(
            f"Fichier {name}", type=["xlsx", "xls"], key=f"file_{name}"
        )

# --- Bloc coller les données ---
if paste_sources:
    st.header("📋 Compagnies à coller")
    for name in paste_sources:
        pasted[name] = st.text_area(
            f"Données {name} (en-tête inclus, séparateur Tab ou ;)",
            height=180, key=f"paste_{name}"
        )

st.divider()

# ---------------------------------------------------------------
# GO
# ---------------------------------------------------------------
if st.button("🚀 GO — Générer le fichier", type="primary"):
    frames = []

    # Sources Excel
    for name, conf in excel_sources.items():
        up = uploaded.get(name)
        if up is not None:
            try:
                df_in = read_excel_source(up, conf)
                res = transform(df_in, conf, label=name)
                if res is not None and not res.empty:
                    frames.append(res)
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
            except Exception as e:
                st.error(f"[{name}] Erreur de lecture : {e}")

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
