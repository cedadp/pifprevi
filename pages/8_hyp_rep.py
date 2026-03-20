import streamlit as st
import pandas as pd
import io
from itertools import product

st.set_page_config(page_title="Analyse Correspondances Aéroport", layout="wide")
st.title("✈️ Analyse des Correspondances Aéroport")

uploaded_file = st.file_uploader("📂 Déposez votre fichier Excel ici", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Données originales (20 premières lignes)")
    st.dataframe(df.head(20))
    st.write(f"Dimensions originales : {df.shape[0]} lignes × {df.shape[1]} colonnes")
    
    col_apport_terminal = df.columns[0]
    col_apport_salle = df.columns[1]
    col_emport_terminal = df.columns[2]
    col_emport_salle = df.columns[3]
    col_aibt = df.columns[4]
    col_aobt = df.columns[5]
    col_pax = df.columns[6]

    # Try to detect "plage" column if it exists
    col_plage = None
    for c in df.columns:
        if 'plage' in str(c).lower():
            col_plage = c
            break

    df_trans = df.copy()

    # --- Transformations Terminal Apport ---
    regrouper_abcdt1 = ['T1', 'T2A', 'T2B', 'T2C', 'T2D']
    df_trans[col_apport_terminal] = df_trans[col_apport_terminal].apply(
        lambda x: 'ABCDT1' if str(x).strip() in regrouper_abcdt1 else (
            'C2F' if str(x).strip() == 'T2F' else (
            'C2G' if str(x).strip() == 'T2G' else x))
    )

    # --- Transformations Salle Débarquement ---
    salle_debar_map = {
        'C2E-JETE': 'salle_K',
        'C2E-S3': 'salle_L',
        'C2E-S4': 'salle_M'
    }
    df_trans[col_apport_salle] = df_trans[col_apport_salle].apply(
        lambda x: salle_debar_map.get(str(x).strip(), x)
    )

    # --- Transformations Terminal Emport ---
    df_trans[col_emport_terminal] = df_trans[col_emport_terminal].apply(
        lambda x: 'ABCDT1' if str(x).strip() in regrouper_abcdt1 else (
            'C2F' if str(x).strip() == 'T2F' else (
            'C2G' if str(x).strip() == 'T2G' else x))
    )

    # --- Transformations Salle Embarquement ---
    salle_embar_map = {
        'C2E-JETE': 'salle_K',
        'C2E-S3': 'salle_L',
        'C2E-S4': 'Salle_M'
    }
    df_trans[col_emport_salle] = df_trans[col_emport_salle].apply(
        lambda x: salle_embar_map.get(str(x).strip(), x)
    )

    # --- Filtre même jour AIBT / AOBT ---
    df_trans[col_aibt] = pd.to_datetime(df_trans[col_aibt], errors='coerce')
    df_trans[col_aobt] = pd.to_datetime(df_trans[col_aobt], errors='coerce')
    
    mask_same_day = df_trans[col_aibt].dt.date == df_trans[col_aobt].dt.date
    nb_avant = len(df_trans)
    df_trans = df_trans[mask_same_day].reset_index(drop=True)
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
