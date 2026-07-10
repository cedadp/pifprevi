import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Concaténateur Prévisions Cies", layout="wide")

# ---------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------
OUTPUT_COLS = ["ArrDep", "CieOpe", "NumVol", "EscDep", "EscArr",
               "DateLocaleMvt", "NbPaxCNT", "NbPaxTOT"]

SOURCES = {
    "AF": {
        "input_type": "excel",
        "label": "AF — Prévisions d'activité",
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
        "custom": None,  # sera assigné après définition de la fonction
    },
    "LH_OUT": {
        "input_type": "excel",
        "label": "LH — Départs (outbound)",
        "custom": None,  # sera assigné après définition de la fonction
    },
}


# ---------------------------------------------------------------
# FONCTIONS UTILITAIRES
# ---------------------------------------------------------------

def normalize_columns(df):
    """Normalise les noms de colonnes : espaces multiples, majuscules, accents."""
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


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


def transform(df, conf, label=""):
    """Transformation générique avec mapping (type AF)."""
    df = normalize_columns(df)
    mapping = conf["mapping"]

    missing = [c for c in mapping if c not in df.columns]
    if missing:
        st.error(f"[{label}] Colonnes manquantes : {missing}")
        st.write(f"[{label}] Colonnes trouvées :", list(df.columns))
        return None

    out = df[list(mapping.keys())].copy()
    out.columns = list(mapping.values())

    # Conversion numériques
    for col in ["NbPaxCNT", "NbPaxTOT"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).astype(int)

    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)

    # Format date
    date_col = conf.get("date_col", "DateLocaleMvt")
    out["DateLocaleMvt"] = pd.to_datetime(out[date_col], errors="coerce").dt.strftime("%d/%m/%Y")

    # Nettoyage colonnes texte
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()

    # Filtre compagnie
    filter_cie = conf.get("filter_cie")
    if filter_cie:
        out = out[out["CieOpe"] == filter_cie]

    # Exclusion 0 pax
    if conf.get("exclude_zero_pax", False):
        out = out[out["NbPaxTOT"] != 0]

    return out[OUTPUT_COLS] if not out.empty else None


def transform_lh_inbound(file):
    """Arrivées LHG : Arr Date fusionnée (ffill), EscArr = CDG."""
    df = pd.read_excel(file, sheet_name=0, header=0)
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
    """Nettoyage commun + exclusion 0 pax."""
    out["NumVol"] = pd.to_numeric(out["NumVol"], errors="coerce").fillna(0).astype(int)
    for c in ["ArrDep", "CieOpe", "EscDep", "EscArr"]:
        out[c] = out[c].astype(str).str.strip()
    out = out[out["NbPaxTOT"] != 0]
    out = out[out["CieOpe"].notna() & (out["CieOpe"] != "")]
    return out[OUTPUT_COLS]


def read_excel_source(file, conf):
    """Lecture fichier Excel selon config (sheet standard)."""
    sheet = conf.get("sheet", 0)
    return pd.read_excel(file, sheet_name=sheet, header=0)


# ===============================================================
# ASSIGNATION DES FONCTIONS CUSTOM
# ===============================================================
SOURCES["LH_IN"]["custom"] = transform_lh_inbound
SOURCES["LH_OUT"]["custom"] = transform_lh_outbound


# ===============================================================
# INTERFACE STREAMLIT
# ===============================================================

st.title("📊 Concaténateur Prévisions Compagnies")
st.write("Fusionnez les données de plusieurs sources (AF, LH, etc.) en un seul fichier CSV.")

with st.form("form_concatenate"):
    st.subheader("📥 Chargement des données")

    uploaded = {}
    pasted = {}
    excel_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "excel"}
    paste_sources = {k: v for k, v in SOURCES.items() if v["input_type"] == "paste"}

    # ===== SOURCES EXCEL =====
    if excel_sources:
        st.write("**Fichiers Excel :**")
        cols = st.columns(len(excel_sources))
        for i, (name, conf) in enumerate(excel_sources.items()):
            with cols[i]:
                label = conf.get("label", f"Fichier {name}")
                uploaded[name] = st.file_uploader(label, type=["xlsx", "xls"], key=f"file_{name}")

    # ===== SOURCES À COLLER =====
    if paste_sources:
        st.write("**Données à coller directement :**")
        for name, conf in paste_sources.items():
            label = conf.get("label", f"Données {name}")
            pasted[name] = st.text_area(label, key=f"paste_{name}", height=150)

    # ===== BOUTON GO =====
    go_button = st.form_submit_button("🚀 GO — Générer le CSV", use_container_width=True)

if go_button:
    frames = []

    # ===== TRAITEMENT SOURCES EXCEL =====
    for name, conf in excel_sources.items():
        up = uploaded.get(name)
        if up is None:
            continue
        try:
            if "custom" in conf and conf["custom"] is not None:
                # sources LH (parsing custom)
                res = finalize_output(conf["custom"](up))
            else:
                # sources type AF (mapping)
                df_in = read_excel_source(up, conf)
                res = transform(df_in, conf, label=name)

            if res is not None and not res.empty:
                frames.append(res)
                st.success(f"✅ [{name}] {len(res)} lignes chargées")
            else:
                st.warning(f"⚠️ [{name}] Aucune donnée valide après traitement")

        except Exception as e:
            st.error(f"❌ [{name}] Erreur de lecture : {e}")

    # ===== TRAITEMENT SOURCES À COLLER =====
    for name, conf in paste_sources.items():
        txt = pasted.get(name, "")
        if txt and txt.strip():
            try:
                # À adapter selon le format du texte collé (CSV, TSV, etc.)
                from io import StringIO
                df_in = pd.read_csv(StringIO(txt), sep=None, engine="python")
                res = transform(df_in, conf, label=name)

                if res is not None and not res.empty:
                    frames.append(res)
                    st.success(f"✅ [{name}] {len(res)} lignes chargées")
                else:
                    st.warning(f"⚠️ [{name}] Aucune donnée valide après traitement")

            except Exception as e:
                st.error(f"❌ [{name}] Erreur de lecture : {e}")

    # ===== CONCATÉNATION ET SORTIE =====
    if not frames:
        st.warning("⚠️ Aucune donnée valide fournie.")
    else:
        final = pd.concat(frames, ignore_index=True)
        final = final.sort_values(by=["DateLocaleMvt", "CieOpe", "NumVol"]).reset_index(drop=True)

        st.subheader(f"📦 Fichier final ({len(final)} lignes)")
        st.dataframe(final.head(100), use_container_width=True)

        # ===== TÉLÉCHARGEMENT CSV =====
        header = ";".join(OUTPUT_COLS)
        body = final.to_csv(sep=";", index=False, header=False, lineterminator="\n", encoding="utf-8")
        csv_out = header + "\n" + body

        timestamp = datetime.now().strftime("%d_%m_%Y")
        filename = f"Previs_cies_{timestamp}.csv"

        st.download_button(
            "💾 Télécharger le CSV",
            data=csv_out.encode("utf-8"),
            file_name=filename,
            mime="text/csv",
            use_container_width=True,
        )

        st.info(f"💡 Fichier : `{filename}` | Encodage : UTF-8 | Séparateur : `;`")
