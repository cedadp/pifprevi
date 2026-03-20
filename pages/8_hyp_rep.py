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
    nb_apres = len(df_trans)

    st.subheader("✅ Données transformées")
    st.dataframe(df_trans)

    col1, col2, col3 = st.columns(3)
    col1.metric("Lignes conservées", nb_apres)
    col2.metric("Lignes supprimées", nb_avant - nb_apres)
    col3.metric("Total passagers", int(df_trans[col_pax].sum()))

    # ============================================================
    # SECTION 2 : TAUX DE CORRESPONDANCE PAR COMBINAISON
    # ============================================================
    st.markdown("---")
    st.header("📊 Taux de Correspondance par Combinaison")

    st.markdown("""
    Cette section calcule le **nombre de passagers en correspondance** et le **taux de correspondance**
    pour toutes les combinaisons possibles entre les entités suivantes :
    `ABCDT1`, `C2F`, `C2G`, `salle_K`, `salle_L`, `salle_M / Salle_M`
    """)

    # Définir les entités cibles
    entites_apport = ['ABCDT1', 'C2F', 'C2G', 'salle_K', 'salle_L', 'salle_M']
    entites_emport = ['ABCDT1', 'C2F', 'C2G', 'salle_K', 'salle_L', 'Salle_M']

    # Normaliser salle_M / Salle_M pour le calcul
    # On crée des colonnes "depuis" et "vers" qui combinent terminal + salle selon les règles
    # Règle : 
    #   - Pour ABCDT1, C2F, C2G => on prend le terminal (pas la salle)
    #   - Pour salle_K, salle_L, salle_M/Salle_M => on prend la salle (pas le terminal)

    def get_depuis(row):
        terminal = str(row[col_apport_terminal]).strip()
        salle = str(row[col_apport_salle]).strip()
        if terminal in ['ABCDT1', 'C2F', 'C2G']:
            # Vérifier si la salle est une salle spéciale
            if salle in ['salle_K', 'salle_L', 'salle_M']:
                return salle
            else:
                return terminal
        else:
            return terminal

    def get_vers(row):
        terminal = str(row[col_emport_terminal]).strip()
        salle = str(row[col_emport_salle]).strip()
        if terminal in ['ABCDT1', 'C2F', 'C2G']:
            # Vérifier si la salle est une salle spéciale
            if salle in ['salle_K', 'salle_L', 'Salle_M']:
                # Normaliser Salle_M -> salle_M pour regroupement
                return 'salle_M' if salle == 'Salle_M' else salle
            else:
                return terminal
        else:
            return terminal

    df_trans['_depuis'] = df_trans.apply(get_depuis, axis=1)
    df_trans['_vers'] = df_trans.apply(get_vers, axis=1)

    # Toutes les entités présentes
    toutes_entites = sorted(set(df_trans['_depuis'].unique()) | set(df_trans['_vers'].unique()))
    
    # Filtrer uniquement les entités qui nous intéressent
    entites_cibles = ['ABCDT1', 'C2F', 'C2G', 'salle_K', 'salle_L', 'salle_M']
    entites_presentes_depuis = [e for e in entites_cibles if e in df_trans['_depuis'].values]
    entites_presentes_vers = [e for e in entites_cibles if e in df_trans['_vers'].values]

    st.markdown("#### 🔧 Paramètres d'analyse")
    
    col_param1, col_param2 = st.columns(2)
    with col_param1:
        entites_depuis_sel = st.multiselect(
            "Entités DEPUIS (apport)",
            options=entites_cibles,
            default=entites_presentes_depuis
        )
    with col_param2:
        entites_vers_sel = st.multiselect(
            "Entités VERS (emport)",
            options=entites_cibles,
            default=entites_presentes_vers
        )

    # Filtre optionnel par plage
    filtre_plage = None
    if col_plage is not None:
        st.markdown("#### 🕐 Filtre par plage horaire")
        plages_dispo = ['Toutes'] + sorted(df_trans[col_plage].dropna().unique().tolist())
        filtre_plage = st.selectbox("Sélectionner une plage", plages_dispo)

    df_calcul = df_trans.copy()
    if col_plage is not None and filtre_plage and filtre_plage != 'Toutes':
        df_calcul = df_calcul[df_calcul[col_plage] == filtre_plage]

    if entites_depuis_sel and entites_vers_sel:
        # --- Tableau croisé : nombre de passagers ---
        st.markdown("### 🔢 Nombre de passagers en correspondance (matrice Depuis → Vers)")
        
        df_filtre = df_calcul[
            df_calcul['_depuis'].isin(entites_depuis_sel) & 
            df_calcul['_vers'].isin(entites_vers_sel)
        ]

        pivot_pax = df_filtre.pivot_table(
            index='_depuis',
            columns='_vers',
            values=col_pax,
            aggfunc='sum',
            fill_value=0
        )
        
        # Réindexer pour avoir toutes les entités sélectionnées
        pivot_pax = pivot_pax.reindex(index=entites_depuis_sel, columns=entites_vers_sel, fill_value=0)
        
        # Ajouter totaux
        pivot_pax['TOTAL'] = pivot_pax.sum(axis=1)
        pivot_pax.loc['TOTAL'] = pivot_pax.sum(axis=0)
        
        st.dataframe(pivot_pax.style.format("{:.0f}").background_gradient(cmap='Blues', subset=pd.IndexSlice[entites_depuis_sel, entites_vers_sel]))

        # --- Tableau croisé : taux de correspondance (% ligne) ---
        st.markdown("### 📈 Taux de correspondance (% du total passagers par entité DEPUIS)")
        
        pivot_pax_sans_total = pivot_pax.loc[entites_depuis_sel, entites_vers_sel]
        total_par_depuis = pivot_pax_sans_total.sum(axis=1)
        
        pivot_taux = pivot_pax_sans_total.div(total_par_depuis, axis=0) * 100
        pivot_taux = pivot_taux.fillna(0)
        pivot_taux['TOTAL %'] = pivot_taux.sum(axis=1)
        
        st.dataframe(pivot_taux.style.format("{:.1f}%").background_gradient(cmap='Greens', subset=pd.IndexSlice[entites_depuis_sel, entites_vers_sel]))

        # --- Tableau croisé : taux de correspondance (% colonne) ---
        st.markdown("### 📉 Taux de correspondance (% du total passagers par entité VERS)")
        
        total_par_vers = pivot_pax_sans_total.sum(axis=0)
        pivot_taux_col = pivot_pax_sans_total.div(total_par_vers, axis=1) * 100
        pivot_taux_col = pivot_taux_col.fillna(0)
        pivot_taux_col.loc['TOTAL %'] = pivot_taux_col.sum(axis=0)
        
        st.dataframe(pivot_taux_col.style.format("{:.1f}%").background_gradient(cmap='Oranges', subset=pd.IndexSlice[entites_depuis_sel, entites_vers_sel]))

        # --- Tableau détail par combinaison ---
        st.markdown("### 📋 Détail par combinaison Depuis → Vers")
        
        detail = df_filtre.groupby(['_depuis', '_vers'])[col_pax].sum().reset_index()
        detail.columns = ['Depuis', 'Vers', 'Nb Passagers']
        total_global = detail['Nb Passagers'].sum()
        detail['% Global'] = (detail['Nb Passagers'] / total_global * 100).round(1)
        detail['% Depuis'] = detail.groupby('Depuis')['Nb Passagers'].transform(lambda x: (x / x.sum() * 100).round(1))
        detail = detail.sort_values('Nb Passagers', ascending=False).reset_index(drop=True)
        
        st.dataframe(detail.style.format({'Nb Passagers': '{:.0f}', '% Global': '{:.1f}%', '% Depuis': '{:.1f}%'})
                     .background_gradient(cmap='Purples', subset=['Nb Passagers']))

        # Métriques globales
        st.markdown("### 📌 Métriques globales")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total passagers analysés", int(total_global))
        m2.metric("Nb combinaisons actives", len(detail))
        m3.metric("Combinaison top", f"{detail.iloc[0]['Depuis']} → {detail.iloc[0]['Vers']}")
        m4.metric("Pax top combinaison", int(detail.iloc[0]['Nb Passagers']))

        # --- Export Excel enrichi ---
        st.markdown("---")
        st.subheader("💾 Export Excel")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_trans.drop(columns=['_depuis', '_vers']).to_excel(writer, sheet_name='Données transformées', index=False)
            pivot_pax.to_excel(writer, sheet_name='Matrice Pax')
            pivot_taux.to_excel(writer, sheet_name='Taux % Depuis')
            pivot_taux_col.to_excel(writer, sheet_name='Taux % Vers')
            detail.to_excel(writer, sheet_name='Détail combinaisons', index=False)
        
        output.seek(0)
        st.download_button(
            label="📥 Télécharger le fichier Excel complet",
            data=output,
            file_name="hyp_rep_1_analyse_correspondances.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Veuillez sélectionner au moins une entité DEPUIS et une entité VERS.")
else:
    st.info("👆 Veuillez déposer un fichier Excel pour commencer l'analyse.")
