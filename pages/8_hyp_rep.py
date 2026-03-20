import streamlit as st 
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Transformation Correspondances", layout="wide")
st.title("🛫 Transformation des données de correspondances aéroportuaires")

uploaded_file = st.file_uploader("📂 Déposez votre fichier Excel ici", type=["xlsx"])

def transform_data(df):
    # --- Renommage colonnes pour faciliter le traitement ---
    col_map = {
        df.columns[0]: "Vol_apport_Terminal",
        df.columns[1]: "Vol_apport_Salle_deb",
        df.columns[2]: "Vol_emport_Terminal",
        df.columns[3]: "Vol_emport_Salle_emb",
        df.columns[4]: "AIBT",
        df.columns[5]: "AOBT",
        df.columns[6]: "Nb_pax_correspondance"
    }
    df = df.rename(columns=col_map)

    # --- Conversion des dates ---
    df["AIBT"] = pd.to_datetime(df["AIBT"])
    df["AOBT"] = pd.to_datetime(df["AOBT"])

    # --- Transformation Vol apport - Terminal ---
    def transform_terminal_apport(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val in ["T1", "T2A", "T2B", "T2C", "T2D"]:
            return "ABCDT1"
        elif val == "T2E":
            return "T2E"
        elif val == "T2F":
            return "C2F"
        elif val == "T2G":
            return "C2G"
        return val

    df["Vol_apport_Terminal"] = df["Vol_apport_Terminal"].apply(transform_terminal_apport)

    # --- Transformation Vol apport - Salle de débarquement ---
    def transform_salle_deb(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        mapping = {
            "C2E-JETE": "salle_K",
            "C2E-S3": "salle_L",
            "C2E-S4": "salle_M"
        }
        return mapping.get(val, val)

    df["Vol_apport_Salle_deb"] = df["Vol_apport_Salle_deb"].apply(transform_salle_deb)

    # --- Transformation Vol emport - Terminal ---
    def transform_terminal_emport(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val in ["T1", "T2A", "T2B", "T2C", "T2D"]:
            return "ABCDT1"
        elif val == "T2E":
            return "T2E"
        elif val == "T2F":
            return "C2F"
        elif val == "T2G":
            return "C2G"
        return val

    df["Vol_emport_Terminal"] = df["Vol_emport_Terminal"].apply(transform_terminal_emport)

    # --- Transformation Vol emport - Salle d'embarquement ---
    def transform_salle_emb(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        mapping = {
            "C2E-JETE": "salle_K",
            "C2E-S3": "salle_L",
            "C2E-S4": "Salle_M"   # S majuscule pour emport
        }
        return mapping.get(val, val)

    df["Vol_emport_Salle_emb"] = df["Vol_emport_Salle_emb"].apply(transform_salle_emb)

    # --- Filtre : conserver uniquement les lignes où AIBT et AOBT sont le même jour ---
    df = df[df["AIBT"].dt.date == df["AOBT"].dt.date].copy()

    # --- Remise en forme des colonnes de dates pour affichage ---
    df["AIBT"] = df["AIBT"].dt.strftime("%Y-%m-%d %H:%M:%S")
    df["AOBT"] = df["AOBT"].dt.strftime("%Y-%m-%d %H:%M:%S")

    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transformé")
    return output.getvalue()

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, header=0)
        st.subheader("📋 Aperçu du fichier original")
        st.dataframe(df_raw.head(20), use_container_width=True)
        st.write(f"**Nombre de lignes originales :** {len(df_raw)}")

        df_transformed = transform_data(df_raw.copy())

        st.subheader("✅ Résultat après transformation et filtrage")
        st.dataframe(df_transformed, use_container_width=True)
        st.write(f"**Nombre de lignes après filtrage (même jour AIBT/AOBT) :** {len(df_transformed)}")

        # --- Statistiques rapides ---
        st.subheader("📊 Statistiques rapides")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Lignes supprimées", len(df_raw) - len(df_transformed))
        with col2:
            st.metric("Total passagers (filtrés)", int(df_transformed["Nb_pax_correspondance"].sum()))
        with col3:
            st.metric("Lignes conservées", len(df_transformed))

        # --- Téléchargement ---
        excel_data = to_excel(df_transformed)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel transformé",
            data=excel_data,
            file_name="hyp_rep_1_transformed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erreur lors du traitement : {e}")
else:
    st.info("👆 Veuillez déposer un fichier Excel pour commencer.")
