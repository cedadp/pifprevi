import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Transformation fichier correspondances", layout="wide")
st.title("📊 Transformation & Analyse des correspondances aéroportuaires")

uploaded_file = st.file_uploader("📁 Déposez votre fichier Excel ici", type=["xlsx"])

if uploaded_file is not None:
    # ─────────────────────────────────────────────
    # LECTURE
    # ─────────────────────────────────────────────
    df = pd.read_excel(uploaded_file)

    st.subheader("1️⃣ Aperçu du fichier original")
    st.dataframe(df.head(20), use_container_width=True)
    st.write(f"**Dimensions originales :** {df.shape[0]} lignes × {df.shape[1]} colonnes")

    # ─────────────────────────────────────────────
    # RENOMMAGE DES COLONNES (robuste)
    # ─────────────────────────────────────────────
    col_mapping = {}
    for col in df.columns:
        col_strip = col.strip()
        if "terminal" in col_strip.lower() and "apport" in col_strip.lower():
            col_mapping[col] = "terminal_apport"
        elif "salle de débarquement" in col_strip.lower() or "débarquement" in col_strip.lower():
            col_mapping[col] = "salle_debarquement"
        elif "terminal" in col_strip.lower() and "emport" in col_strip.lower():
            col_mapping[col] = "terminal_emport"
        elif "salle d'embarquement" in col_strip.lower() or "embarquement" in col_strip.lower():
            col_mapping[col] = "salle_embarquement"
        elif "aibt" in col_strip.lower():
            col_mapping[col] = "AIBT"
        elif "aobt" in col_strip.lower():
            col_mapping[col] = "AOBT"
        elif "passager" in col_strip.lower() or "nombre" in col_strip.lower():
            col_mapping[col] = "nb_passagers"
        elif "plage" in col_strip.lower():
            col_mapping[col] = "plage"

    df.rename(columns=col_mapping, inplace=True)

    # ─────────────────────────────────────────────
    # TRANSFORMATIONS TERMINAUX APPORT
    # ─────────────────────────────────────────────
    def transform_terminal_apport(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val in ["T1", "T2A", "T2B", "T2C", "T2D"]:
            return "ABCDT1"
        elif val == "T2F":
            return "C2F"
        elif val == "T2G":
            return "C2G"
        elif val == "T2E":
            return "T2E"
        return val

    def transform_terminal_emport(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val in ["T1", "T2A", "T2B", "T2C", "T2D"]:
            return "ABCDT1"
        elif val == "T2F":
            return "C2F"
        elif val == "T2G":
            return "C2G"
        elif val == "T2E":
            return "T2E"
        return val

    # ─────────────────────────────────────────────
    # TRANSFORMATIONS SALLES DÉBARQUEMENT
    # ─────────────────────────────────────────────
    def transform_salle_debarquement(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val == "C2E-JETE":
            return "salle_K"
        elif val == "C2E-S3":
            return "salle_L"
        elif val in ["C2E-S4", "C2E_S4"]:
            return "salle_M"
        return val

    # ─────────────────────────────────────────────
    # TRANSFORMATIONS SALLES EMBARQUEMENT
    # ─────────────────────────────────────────────
    def transform_salle_embarquement(val):
        if pd.isna(val):
            return val
        val = str(val).strip()
        if val == "C2E-JETE":
            return "salle_K"
        elif val == "C2E-S3":
            return "salle_L"
        elif val in ["C2E-S4", "C2E_S4"]:
            return "Salle_M"   # S majuscule intentionnel
        return val

    if "terminal_apport" in df.columns:
        df["terminal_apport"] = df["terminal_apport"].apply(transform_terminal_apport)
    if "terminal_emport" in df.columns:
        df["terminal_emport"] = df["terminal_emport"].apply(transform_terminal_emport)
    if "salle_debarquement" in df.columns:
        df["salle_debarquement"] = df["salle_debarquement"].apply(transform_salle_debarquement)
    if "salle_embarquement" in df.columns:
        df["salle_embarquement"] = df["salle_embarquement"].apply(transform_salle_embarquement)

    # ─────────────────────────────────────────────
    # FILTRE MÊME JOUR AIBT / AOBT
    # ─────────────────────────────────────────────
    nb_avant = len(df)

    if "AIBT" in df.columns and "AOBT" in df.columns:
        df["AIBT"] = pd.to_datetime(df["AIBT"], errors="coerce")
        df["AOBT"] = pd.to_datetime(df["AOBT"], errors="coerce")
        df = df[df["AIBT"].dt.date == df["AOBT"].dt.date]

    nb_apres = len(df)

    st.subheader("2️⃣ Données transformées et filtrées")
    st.dataframe(df, use_container_width=True)

    col1, col2, col3 = st.columns(3)
    col1.metric("Lignes originales", nb_avant)
    col2.metric("Lignes conservées", nb_apres)
    col3.metric("Lignes supprimées", nb_avant - nb_apres)

    if "nb_passagers" in df.columns:
        st.metric("Total passagers en correspondance", int(df["nb_passagers"].sum()))

    # ─────────────────────────────────────────────
    # EXPORT FICHIER TRANSFORMÉ
    # ─────────────────────────────────────────────
    output1 = io.BytesIO()
    with pd.ExcelWriter(output1, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transformé")
    output1.seek(0)
    st.download_button(
        label="⬇️ Télécharger le fichier transformé",
        data=output1,
        file_name="fichier_transforme.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ═════════════════════════════════════════════
    # PARTIE 2 — TAUX DE CORRESPONDANCE
    # ═════════════════════════════════════════════
    st.markdown("---")
    st.header("📈 Taux de correspondance")

    # ─────────────────────────────────────────────
    # Création de la colonne "origine" et "destination"
    # L'origine correspond à :
    #   - salle_debarquement si terminal_apport == T2E (salles K/L/M)
    #   - terminal_apport sinon (ABCDT1, C2F, C2G)
    # La destination correspond à :
    #   - salle_embarquement si terminal_emport == T2E (salles K/L/M)
    #   - terminal_emport sinon (ABCDT1, C2F, C2G)
    # ─────────────────────────────────────────────

    def get_origine(row):
        t = str(row.get("terminal_apport", "")).strip()
        s = str(row.get("salle_debarquement", "")).strip()
        # Si le terminal apport est T2E, on utilise la salle
        if t == "T2E":
            return s if s not in ["", "nan"] else t
        # Sinon on utilise directement le terminal transformé (ABCDT1, C2F, C2G)
        return t

    def get_destination(row):
        t = str(row.get("terminal_emport", "")).strip()
        s = str(row.get("salle_embarquement", "")).strip()
        if t == "T2E":
            return s if s not in ["", "nan"] else t
        return t

    df["origine"]      = df.apply(get_origine, axis=1)
    df["destination"]  = df.apply(get_destination, axis=1)

    # ─────────────────────────────────────────────
    # Origines et destinations d'intérêt
    # ─────────────────────────────────────────────
    ORIGINES      = ["salle_K", "salle_L", "salle_M", "C2F", "C2G"]
    DESTINATIONS  = ["ABCDT1", "C2F", "C2G", "salle_K", "salle_L", "Salle_M"]

    # ─────────────────────────────────────────────
    # Construction de la matrice OD (passagers)
    # ─────────────────────────────────────────────
    pax_col = "nb_passagers" if "nb_passagers" in df.columns else None

    rows_matrix = []
    for orig in ORIGINES:
        df_orig = df[df["origine"] == orig]
        total_orig = int(df_orig[pax_col].sum()) if pax_col else len(df_orig)
        row_data = {"Origine \\ Destination": orig, "Total arrivée": total_orig}
        for dest in DESTINATIONS:
            df_od = df_orig[df_orig["destination"] == dest]
            pax_od = int(df_od[pax_col].sum()) if pax_col else len(df_od)
            row_data[dest] = pax_od
        rows_matrix.append(row_data)

    matrix_pax = pd.DataFrame(rows_matrix).set_index("Origine \\ Destination")

    # ─────────────────────────────────────────────
    # Construction de la matrice des taux (%)
    # ─────────────────────────────────────────────
    rows_taux = []
    for orig in ORIGINES:
        df_orig = df[df["origine"] == orig]
        total_orig = int(df_orig[pax_col].sum()) if pax_col else len(df_orig)
        row_data = {"Origine \\ Destination": orig, "Total arrivée": total_orig}
        for dest in DESTINATIONS:
            df_od = df_orig[df_orig["destination"] == dest]
            pax_od = int(df_od[pax_col].sum()) if pax_col else len(df_od)
            taux = round(pax_od / total_orig * 100, 1) if total_orig > 0 else 0.0
            row_data[dest] = taux
        rows_taux.append(row_data)

    matrix_taux = pd.DataFrame(rows_taux).set_index("Origine \\ Destination")

    # ─────────────────────────────────────────────
    # Affichage
    # ─────────────────────────────────────────────
    st.subheader("🔢 Matrice des passagers (nombre absolu)")
    st.dataframe(
        matrix_pax.style.background_gradient(cmap="Blues", subset=DESTINATIONS),
        use_container_width=True
    )

    st.subheader("📊 Matrice des taux de correspondance (%)")
    st.markdown(
        "*Lecture : pour 100 passagers arrivant à l'origine, X% repartent vers la destination.*"
    )

    # Formatage avec % et gradient de couleur
    styled_taux = (
        matrix_taux[DESTINATIONS]
        .style
        .background_gradient(cmap="RdYlGn", axis=None, vmin=0, vmax=100)
        .format("{:.1f}%")
    )
    st.dataframe(styled_taux, use_container_width=True)

    # ─────────────────────────────────────────────
    # Graphique en barres empilées par origine
    # ─────────────────────────────────────────────
    st.subheader("📉 Visualisation graphique des taux (%)")

    import matplotlib.pyplot as plt
    import numpy as np

    fig, ax = plt.subplots(figsize=(12, 5))
    x = np.arange(len(ORIGINES))
    width = 0.13
    colors = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948"]

    for i, dest in enumerate(DESTINATIONS):
        vals = [matrix_taux.loc[orig, dest] if orig in matrix_taux.index else 0
                for orig in ORIGINES]
        ax.bar(x + i * width, vals, width, label=dest, color=colors[i % len(colors)])

    ax.set_xlabel("Origine (salle / terminal d'arrivée)", fontsize=11)
    ax.set_ylabel("Taux de correspondance (%)", fontsize=11)
    ax.set_title("Taux de correspondance par origine et destination", fontsize=13, fontweight="bold")
    ax.set_xticks(x + width * (len(DESTINATIONS) - 1) / 2)
    ax.set_xticklabels(ORIGINES, fontsize=10)
    ax.legend(title="Destination", bbox_to_anchor=(1.01, 1), loc="upper left")
    ax.set_ylim(0, 110)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)

    # ─────────────────────────────────────────────
    # EXPORT FICHIER ANALYSE
    # ─────────────────────────────────────────────
    output2 = io.BytesIO()
    with pd.ExcelWriter(output2, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Données transformées")
        matrix_pax.to_excel(writer, sheet_name="Matrice passagers")
        matrix_taux.to_excel(writer, sheet_name="Matrice taux (%)")
    output2.seek(0)
    st.download_button(
        label="⬇️ Télécharger l'analyse complète (Excel)",
        data=output2,
        file_name="analyse_correspondances.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Veuillez déposer un fichier Excel pour commencer l'analyse.")
