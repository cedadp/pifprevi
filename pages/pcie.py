import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Concaténateur Prévisions Cies", layout="wide")

# ---------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------
OUTPUT_COLS = ["ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
               "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT"]

SOURCES = {
    "AF": {
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
        "filter_cie": "AF",       # ne garder que cette compagnie
        "exclude_zero_pax": True,  # exclure NbPaxTOT == 0
    },
    # Les futures sources (formats/libellés différents) s'ajouteront ici
}


def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def transform(df, conf):
    df = normalize_columns(df)
    mapping = conf["mapping"]

    missing = [c for c in mapping if c not in df.columns]
    if missing:
        st.error(f"Colonnes manquantes : {missing}")
        st.write("Colonnes trouvées :", list(df.columns))
        return None

    out = df[list(mapping.keys())].rename(columns=mapping)

    # Date -> JJ/MM/AAAA
    dcol = conf["date_col"]
    out[dcol] = pd.to_datetime(out[dcol], errors="coerce").dt.strftime("%d/%m/%Y")

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

    # Retirer lignes sans date valide
    out = out[out["DateLocaleMvt"].notna()]

    st.caption(f"{n0} lignes lues → {len(out)} lignes conservées après filtrage.")

    return out[OUTPUT_COLS].reset_index(drop=True)


# ---------------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------------
st.title("🛫 Concaténateur de prévisions Compagnies")

if "accumulated" not in st.session_state:
    st.session_state.accumulated = []

source_name = st.selectbox("Source des données", list(SOURCES.keys()))
conf = SOURCES[source_name]

mode = st.radio("Mode d'import", ["Fichier Excel", "Coller les données"], horizontal=True)

df_input = None

if mode == "Fichier Excel":
    up = st.file_uploader("Déposer le fichier Excel", type=["xlsx", "xls"])
    if up:
        xls = pd.ExcelFile(up)
        if conf["sheet"] in xls.sheet_names:
            sheet = conf["sheet"]
        else:
            st.warning(f"Onglet « {conf['sheet']} » introuvable. Sélectionnez-en un :")
            sheet = st.selectbox("Onglet", xls.sheet_names)
        df_input = pd.read_excel(xls, sheet_name=sheet)

else:
    txt = st.text_area("Collez les données (en-tête inclus, séparateur Tab ou ;)",
                        height=250)
    if txt.strip():
        sep = "\t" if "\t" in txt else ";"
        df_input = pd.read_csv(io.StringIO(txt), sep=sep)

if df_input is not None:
    with st.expander("Aperçu des données source"):
        st.dataframe(df_input.head(20))

    result = transform(df_input, conf)
    if result is not None and not result.empty:
        st.subheader("Résultat formaté")
        st.dataframe(result.head(30))

        if st.button("➕ Ajouter au fichier de sortie"):
            st.session_state.accumulated.append(result)
            st.success(f"{len(result)} lignes ajoutées.")

# ---------------------------------------------------------------
# FICHIER FINAL
# ---------------------------------------------------------------
if st.session_state.accumulated:
    final = pd.concat(st.session_state.accumulated, ignore_index=True)

    st.divider()
    st.subheader(f"📦 Fichier final ({len(final)} lignes)")
    st.dataframe(final.head(50))

    header = ";".join(OUTPUT_COLS)
    body = final.to_csv(sep=";", index=False, header=False, lineterminator="\n")
    csv_out = header + "\n" + body

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "💾 Télécharger le CSV",
            data=csv_out.encode("utf-8"),
            file_name="Previs_cies.csv",
            mime="text/csv",
        )
    with col2:
        if st.button("🗑️ Réinitialiser le fichier final"):
            st.session_state.accumulated = []
            st.rerun()
