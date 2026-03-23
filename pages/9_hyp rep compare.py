import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import io
from itertools import product

st.set_page_config(page_title="Comparaison des fichiers de correspondance", layout="wide")
st.title("🔄 Comparaison de deux fichiers de correspondances")

st.markdown("""
Déposez deux versions du fichier de correspondances (format généré par le module précédent).  
L'outil compare les taux par OD et par plage, et met en évidence les évolutions.
""")

# ─────────────────────────────────────────────
# UPLOAD DES 2 FICHIERS
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("📂 Fichier V1 (référence)")
    file1 = st.file_uploader("Déposez le fichier V1", type=["xlsx"], key="file1")
    label1 = st.text_input("Nom de la version 1", value="Version 1")

with col2:
    st.subheader("📂 Fichier V2 (nouvelle version)")
    file2 = st.file_uploader("Déposez le fichier V2", type=["xlsx"], key="file2")
    label2 = st.text_input("Nom de la version 2", value="Version 2")

# ─────────────────────────────────────────────
# LECTURE DES FICHIERS
# ─────────────────────────────────────────────
def read_od_file(uploaded_file):
    """
    Lit un fichier Excel multi-onglets au format OD.
    Retourne un dict : {od_key: DataFrame avec colonnes [plage, taux]}
    """
    xl = pd.ExcelFile(uploaded_file)
    od_data = {}
    for sheet in xl.sheet_names:
        df_sheet = xl.parse(sheet, header=0)  # ligne 1 = header (heure, Métropole, ...)
        # Colonne A = plage (P1..P7), colonne B = taux (on prend B = Métropole comme valeur de référence)
        df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
        # Première colonne = labels de plage
        plage_col = df_sheet.columns[0]
        taux_col = df_sheet.columns[1]  # colonne B = Métropole (toutes colonnes identiques)
        df_clean = df_sheet[[plage_col, taux_col]].copy()
        df_clean.columns = ["plage", "taux"]
        df_clean = df_clean.dropna(subset=["plage", "taux"])
        df_clean["taux"] = pd.to_numeric(df_clean["taux"], errors="coerce")
        df_clean = df_clean.dropna(subset=["taux"])
        od_data[sheet] = df_clean.reset_index(drop=True)
    return od_data

if file1 and file2:
    data1 = read_od_file(file1)
    data2 = read_od_file(file2)

    ods_v1 = set(data1.keys())
    ods_v2 = set(data2.keys())
    ods_common = sorted(ods_v1 & ods_v2)
    ods_only_v1 = sorted(ods_v1 - ods_v2)
    ods_only_v2 = sorted(ods_v2 - ods_v1)

    # ─────────────────────────────────────────────
    # RÉSUMÉ DES ONGLETS
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📋 Inventaire des OD")
    c1, c2, c3 = st.columns(3)
    c1.metric("OD communes", len(ods_common))
    c2.metric(f"Uniquement dans {label1}", len(ods_only_v1))
    c3.metric(f"Uniquement dans {label2}", len(ods_only_v2))

    if ods_only_v1:
        st.warning(f"OD présentes uniquement dans {label1} : {', '.join(ods_only_v1)}")
    if ods_only_v2:
        st.info(f"OD présentes uniquement dans {label2} : {', '.join(ods_only_v2)}")

    # ─────────────────────────────────────────────
    # CONSTRUCTION DU TABLEAU DE COMPARAISON GLOBAL
    # ─────────────────────────────────────────────
    rows = []
    for od in ods_common:
        df1 = data1[od].set_index("plage")["taux"]
        df2 = data2[od].set_index("plage")["taux"]
        all_plages = sorted(set(df1.index) | set(df2.index))
        for plage in all_plages:
            t1 = df1.get(plage, np.nan)
            t2 = df2.get(plage, np.nan)
            delta = t2 - t1 if not (np.isnan(t1) or np.isnan(t2)) else np.nan
            delta_pct = (delta / t1 * 100) if (not np.isnan(delta) and t1 != 0) else np.nan
            rows.append({
                "OD": od,
                "Plage": plage,
                label1: t1,
                label2: t2,
                "Δ absolu": delta,
                "Δ relatif (%)": delta_pct
            })

    df_compare = pd.DataFrame(rows)

    # ─────────────────────────────────────────────
    # FILTRES INTERACTIFS
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🔍 Tableau de comparaison détaillé")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        selected_ods = st.multiselect("Filtrer par OD", options=ods_common, default=ods_common)
    with col_f2:
        all_plages_list = sorted(df_compare["Plage"].unique())
        selected_plages = st.multiselect("Filtrer par plage", options=all_plages_list, default=all_plages_list)
    with col_f3:
        seuil = st.slider("Afficher uniquement si |Δ absolu| ≥", min_value=0.0, max_value=0.5, value=0.0, step=0.005, format="%.3f")

    df_filtered = df_compare[
        (df_compare["OD"].isin(selected_ods)) &
        (df_compare["Plage"].isin(selected_plages)) &
        (df_compare["Δ absolu"].abs() >= seuil)
    ].copy()

    # Formatage visuel
    def color_delta(val):
        if pd.isna(val):
            return "color: grey"
        if val > 0:
            return "color: green; font-weight: bold"
        if val < 0:
            return "color: red; font-weight: bold"
        return ""

    st.dataframe(
        df_filtered.style
            .applymap(color_delta, subset=["Δ absolu", "Δ relatif (%)"])
            .format({
                label1: "{:.4f}",
                label2: "{:.4f}",
                "Δ absolu": "{:+.4f}",
                "Δ relatif (%)": "{:+.1f}%"
            }, na_rep="—"),
        use_container_width=True,
        height=400
    )

    # ─────────────────────────────────────────────
    # VUE PAR OD INDIVIDUELLE
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📊 Comparaison graphique par OD")

    od_to_plot = st.selectbox("Choisir une OD à visualiser", options=ods_common)

    if od_to_plot:
        df_od = df_compare[df_compare["OD"] == od_to_plot].copy()
        plages = df_od["Plage"].tolist()
        v1_vals = df_od[label1].tolist()
        v2_vals = df_od[label2].tolist()
        delta_vals = df_od["Δ absolu"].tolist()

        x = np.arange(len(plages))
        width = 0.3

        fig, axes = plt.subplots(1, 2, figsize=(14, 5))

        # Graphique 1 : comparaison V1 vs V2
        ax1 = axes[0]
        bars1 = ax1.bar(x - width/2, v1_vals, width, label=label1, color="#4C72B0", alpha=0.85)
        bars2 = ax1.bar(x + width/2, v2_vals, width, label=label2, color="#DD8452", alpha=0.85)
        ax1.set_xticks(x)
        ax1.set_xticklabels(plages)
        ax1.set_title(f"{od_to_plot}\n{label1} vs {label2}")
        ax1.set_ylabel("Taux de correspondance")
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.3f}"))
        ax1.legend()
        ax1.grid(axis="y", alpha=0.3)

        # Graphique 2 : delta absolu par plage
        ax2 = axes[1]
        colors = ["green" if d >= 0 else "red" for d in delta_vals]
        ax2.bar(x, delta_vals, color=colors, alpha=0.8)
        ax2.axhline(0, color="black", linewidth=0.8)
        ax2.set_xticks(x)
        ax2.set_xticklabels(plages)
        ax2.set_title(f"Δ absolu par plage\n({label2} − {label1})")
        ax2.set_ylabel("Δ taux")
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.3f}"))
        ax2.grid(axis="y", alpha=0.3)

        plt.tight_layout()
        st.pyplot(fig)

    # ─────────────────────────────────────────────
    # HEATMAP GLOBALE DES DELTAS
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🗺️ Heatmap des évolutions (Δ absolu)")

    pivot = df_compare.pivot_table(index="OD", columns="Plage", values="Δ absolu", aggfunc="mean")
    pivot = pivot[sorted(pivot.columns)]  # trier les plages

    fig2, ax = plt.subplots(figsize=(max(8, len(pivot.columns) * 1.2), max(6, len(pivot) * 0.5)))
    vmax = max(abs(pivot.values[~np.isnan(pivot.values)]).max(), 0.001)
    cmap = mcolors.TwoSlopeNorm(vmin=-vmax, vcenter=0, vmax=vmax)
    im = ax.imshow(pivot.values, cmap="RdYlGn", norm=cmap, aspect="auto")

    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels(pivot.columns, fontsize=9)
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels(pivot.index, fontsize=8)
    ax.set_title(f"Δ absolu moyen par OD et plage\n({label2} − {label1})", fontsize=12)

    # Annotations
    for i in range(len(pivot.index)):
        for j in range(len(pivot.columns)):
            val = pivot.values[i, j]
            if not np.isnan(val):
                ax.text(j, i, f"{val:+.3f}", ha="center", va="center", fontsize=7,
                        color="black" if abs(val) < vmax * 0.6 else "white")

    plt.colorbar(plt.cm.ScalarMappable(norm=cmap, cmap="RdYlGn"), ax=ax, label="Δ taux")
    plt.tight_layout()
    st.pyplot(fig2)

    # ─────────────────────────────────────────────
    # TOP ÉVOLUTIONS
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🏆 Top évolutions")

    col_top1, col_top2 = st.columns(2)
    n_top = st.slider("Nombre de lignes à afficher", 5, 20, 10)

    df_sorted = df_compare.dropna(subset=["Δ absolu"]).copy()

    with col_top1:
        st.markdown("**📈 Plus fortes hausses**")
        top_up = df_sorted.nlargest(n_top, "Δ absolu")[["OD", "Plage", label1, label2, "Δ absolu"]]
        st.dataframe(top_up.style.format({label1: "{:.4f}", label2: "{:.4f}", "Δ absolu": "{:+.4f}"}),
                     use_container_width=True)

    with col_top2:
        st.markdown("**📉 Plus fortes baisses**")
        top_down = df_sorted.nsmallest(n_top, "Δ absolu")[["OD", "Plage", label1, label2, "Δ absolu"]]
        st.dataframe(top_down.style.format({label1: "{:.4f}", label2: "{:.4f}", "Δ absolu": "{:+.4f}"}),
                     use_container_width=True)

    # ─────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────
    st.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_compare.to_excel(writer, index=False, sheet_name="Comparaison complète")
        df_sorted.nlargest(20, "Δ absolu").to_excel(writer, index=False, sheet_name="Top hausses")
        df_sorted.nsmallest(20, "Δ absolu").to_excel(writer, index=False, sheet_name="Top baisses")
        pivot.to_excel(writer, sheet_name="Heatmap delta")
    output.seek(0)

    st.download_button(
        label="⬇️ Télécharger le rapport de comparaison (Excel)",
        data=output,
        file_name="comparaison_correspondances.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Veuillez déposer les deux fichiers Excel pour démarrer la comparaison.")
