# app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

st.set_page_config(page_title="File d'Attente", layout="wide")

st.title("🛫 File d'Attente")
st.markdown("---")

# --- CHARGEMENT DES DONNÉES ---
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    # Normaliser les noms de colonnes
    df.columns = [c.strip().lower() for c in df.columns]
    
    # Construire un datetime complet à partir de jour et heure
    # heure peut être un timedelta ou un time
    def build_datetime(row):
        jour = pd.Timestamp(row['jour'])
        h = row['heure']
        if isinstance(h, pd.Timedelta) or isinstance(h, np.timedelta64):
            return jour + pd.Timedelta(h)
        elif isinstance(h, str):
            parts = h.split(':')
            return jour + pd.Timedelta(hours=int(parts[0]), minutes=int(parts[1]), seconds=int(parts[2]) if len(parts) > 2 else 0)
        else:
            try:
                return jour + pd.Timedelta(hours=h.hour, minutes=h.minute, seconds=h.second)
            except:
                return jour + pd.Timedelta(h)
    
    df['datetime'] = df.apply(build_datetime, axis=1)
    df = df.sort_values(['site', 'datetime']).reset_index(drop=True)
    return df

uploaded_file = st.file_uploader("📂 Charger le fichier Excel", type=['xlsx', 'xls'])

if uploaded_file is None:
    st.info("Veuillez charger le fichier Excel.")
    st.stop()

df = load_data(uploaded_file)

sites = sorted(df['site'].unique())
jours = sorted(df['jour'].unique())

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Paramètres de simulation")

#--# Débits par défaut par site (en pax/h)
DEFAULT_DEBITS = {'K CTRCNT' : 0,'K CTR' : 1660,'K CNT' : 300 , 'L CTR' : 2080, 'L CNT' :  1520, 'M CTR' : 1960, 'Galerie EF' : 1820, 'C2F' : 1960, 'C2G' : 910, 'Liaison AC' : 1960,'Liaison BD' : 2320,
            'T3': 1260, 'Terminal 1' : 2140, 'Terminal 1_5' : 390, 'Terminal 1_6' : 520, 

            '2E_Arr' :3948 , '2E_Dep' : 4314 ,'Galerie E > F' : 2976 , 'Galerie F > E' : 1356, 'F > S3' :  2482 , 'S3 > F': 1084 ,  '2G_Emport' : 1350, 
                     'AC_Dep' : 2848, 'AC_Arr' :  4268, 'C_Arr' :  1808, 'BD_Arr' : 1825, 'BD_Dep' : 1544, 'T1_Arr' : 3012, 'T1_Dep' : 2071,'T3_Arr' : 1056, 'T3_Dep' : 825}

# Valeur par défaut si un site n'est pas dans le dictionnaire
DEFAULT_DEBIT_FALLBACK = 500



# --- SÉLECTION SITE ET JOUR ---
selected_site = st.sidebar.selectbox("Site", sites)
jour_labels = [pd.Timestamp(j).strftime('%A %d/%m/%Y') for j in jours]
jour_map = dict(zip(jour_labels, jours))
selected_jour_label = st.sidebar.selectbox("Jour", jour_labels)
selected_jour = jour_map[selected_jour_label]

# Filtrer les données
df_site_jour = df[(df['site'] == selected_site) & (df['jour'] == selected_jour)].copy()
df_site_jour = df_site_jour.sort_values('datetime').reset_index(drop=True)

if df_site_jour.empty:
    st.warning("Aucune donnée pour cette sélection.")
    st.stop()

# --- DÉBIT DE SORTIE ---
st.sidebar.markdown("---")
st.sidebar.header("⚙️ Débit de sortie (pax/h)")

mode_debit = st.sidebar.radio(
    "Mode de configuration du débit",
    ["Débit constant sur la journée", "Débit par tranche horaire"]
)

# Heures min/max de la journée
heure_min = df_site_jour['datetime'].min()
heure_max = df_site_jour['datetime'].max()
h_start = heure_min.hour
h_end = min(heure_max.hour + 1, 24)

if mode_debit == "Débit constant sur la journée":
    # Récupérer le débit par défaut pour le site sélectionné
    debit_default = DEFAULT_DEBITS.get(selected_site, DEFAULT_DEBIT_FALLBACK)

    debit_nominal = st.sidebar.number_input(
    f"Débit de sortie (pax/h) — défaut {selected_site}: {debit_default}",
    min_value=0,
    max_value=5000,
    value=debit_default,
    step=50
    )    

    taux_utilisation = st.sidebar.slider(
        "Taux d'utilisation (%)", 
        min_value=0, max_value=100, value=100, step=5
    )
    debit_effectif_paxh = debit_nominal * taux_utilisation / 100.0
    
    # Créer un dictionnaire heure -> débit constant
    debit_par_heure = {h: debit_effectif_paxh for h in range(24)}
    
    st.sidebar.info(f"Débit effectif : **{debit_effectif_paxh:.0f} pax/h** ({debit_effectif_paxh/6:.1f} pax/10min)")

else:
    st.sidebar.markdown("### Définir le débit par tranche")
    
    # Permettre de définir des tranches
    nb_tranches = st.sidebar.number_input("Nombre de tranches horaires", 1, 10, 3)
    
    tranches = []
    debit_par_heure = {}
    
    # Initialiser tout à 0
    for h in range(24):
        debit_par_heure[h] = 0
    
    for i in range(nb_tranches):
        st.sidebar.markdown(f"**Tranche {i+1}**")
        col1, col2 = st.sidebar.columns(2)
        with col1:
            h_deb = st.number_input(f"Début (h)", 0, 23, min(h_start + i * 6, 23), key=f"hdeb_{i}")
        with col2:
            h_fin = st.number_input(f"Fin (h)", 0, 23, min(h_start + (i+1) * 6, 23), key=f"hfin_{i}")
        
        col3, col4 = st.sidebar.columns(2)
        with col3:
            debit_t = st.number_input(f"Débit (pax/h)", 0, 5000, 500, 10, key=f"debit_{i}")
        with col4:
            taux_t = st.slider(f"Taux %", 0, 100, 100, 5, key=f"taux_{i}")
        
        eff = debit_t * taux_t / 100.0
        tranches.append((h_deb, h_fin, eff))
        
        for h in range(h_deb, h_fin + 1):
            if h < 24:
                debit_par_heure[h] = eff

# --- SIMULATION DE LA FILE D'ATTENTE ---
def simulate_queue(df_data, debit_par_heure):
    """
    Simulation pas à pas (10 min) :
    - arrivées = charge sur le créneau
    - capacité de traitement sur 10 min = débit(heure) / 6
    - file = max(0, file_précédente + arrivées - capacité)
    - traités = min(file_précédente + arrivées, capacité)
    - temps d'attente estimé = file / (débit/60) en minutes
    """
    results = []
    queue = 0.0  # nombre de personnes en file
    
    # Créer la grille complète de 10 min
    all_times = pd.date_range(start=df_data['datetime'].min(), 
                               end=df_data['datetime'].max(), 
                               freq='10min')
    
    # Indexer les données existantes
    charge_map = df_data.set_index('datetime')['charge'].to_dict()
    
    for t in all_times:
        arrivees = charge_map.get(t, 0.0)
        heure = t.hour
        debit_h = debit_par_heure.get(heure, 0)
        capacite_10min = debit_h / 6.0  # pax par 10 min
        
        # Personnes disponibles = file + nouvelles arrivées
        disponible = queue + arrivees
        
        # Traités sur ce créneau
        traites = min(disponible, capacite_10min)
        
        # Nouvelle file
        queue = max(0, disponible - capacite_10min)
        
        # Temps d'attente estimé (minutes)
        if debit_h > 0:
            temps_attente = (queue / debit_h) * 60  # en minutes
        else:
            temps_attente = 0 if queue == 0 else float('inf')
        
        results.append({
            'datetime': t,
            'heure_str': t.strftime('%H:%M'),
            'arrivees': arrivees,
            'debit_paxh': debit_h,
            'capacite_10min': capacite_10min,
            'traites': traites,
            'file_attente': queue,
            'temps_attente_min': temps_attente
        })
    
    return pd.DataFrame(results)

df_sim = simulate_queue(df_site_jour, debit_par_heure)

# --- AFFICHAGE DES RÉSULTATS ---
st.header(f"📊 {selected_site} — {pd.Timestamp(selected_jour).strftime('%A %d/%m/%Y')}")

# KPIs
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total arrivées file", f"{df_sim['arrivees'].sum():.0f} pax")
with col2:
    st.metric("File max", f"{df_sim['file_attente'].max():.0f} pax")
with col3:
    max_attente = df_sim['temps_attente_min'].replace(float('inf'), np.nan).max()
    st.metric("Attente max", f"{max_attente:.0f} min" if pd.notna(max_attente) else "N/A")
with col4:
    st.metric("Total traités", f"{df_sim['traites'].sum():.0f} pax")

# --- GRAPHIQUE PRINCIPAL ---
fig = make_subplots(
    rows=3, cols=1,
    shared_xaxes=True,
    vertical_spacing=0.06,
    subplot_titles=(
        "Arrivées file vs Capacité de traitement (pax / 10 min)",
        "File d'attente (pax)",
        #"Temps d'attente estimé (min)"
    ),
    row_heights=[0.35, 0.35, 0.3]
)

# Graph 1 : Arrivées vs Capacité
fig.add_trace(
    go.Bar(
        x=df_sim['datetime'], y=df_sim['arrivees'],
        name='Arrivées file (pax/10min)',
        marker_color='steelblue', opacity=0.7
    ), row=1, col=1
)
fig.add_trace(
    go.Scatter(
        x=df_sim['datetime'], y=df_sim['capacite_10min'],
        name='Capacité (pax/10min)',
        line=dict(color='red', width=2, dash='dash')
    ), row=1, col=1
)
fig.add_trace(
    go.Scatter(
        x=df_sim['datetime'], y=df_sim['traites'],
        name='Traités (pax/10min)',
        line=dict(color='green', width=1.5),
        fill='tozeroy', fillcolor='rgba(0,200,0,0.1)'
    ), row=1, col=1
)

# Graph 2 : File d'attente
fig.add_trace(
    go.Scatter(
        x=df_sim['datetime'], y=df_sim['file_attente'],
        name="File d'attente",
        fill='tozeroy',
        line=dict(color='orange', width=2),
        fillcolor='rgba(255,165,0,0.3)'
    ), row=2, col=1
)

#Graph 3 : Temps d'attente
temps_display = df_sim['temps_attente_min'].replace(float('inf'), np.nan)
fig.add_trace(
    go.Scatter(
        x=df_sim['datetime'], y=temps_display,
        name='Temps attente (min)',
        fill='tozeroy',
        line=dict(color='crimson', width=2),
        fillcolor='rgba(220,20,60,0.2)'
    ), row=3, col=1
)

fig.update_layout(
    height=900,
    showlegend=True,
   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    hovermode='x unified'
)
fig.update_xaxes(title_text="Heure", row=3, col=1)
fig.update_yaxes(title_text="pax/10min", row=1, col=1)
fig.update_yaxes(title_text="pax", row=2, col=1)
#fig.update_yaxes(title_text="minutes", row=3, col=1)

st.plotly_chart(fig, use_container_width=True)

#--- TABLEAU DÉTAILLÉ ---
with st.expander("📋 Voir le tableau détaillé"):
    df_display = df_sim[['heure_str', 'arrivees', 'debit_paxh', 'capacite_10min', 
                          'traites', 'file_attente', 'temps_attente_min']].copy()
    df_display.columns = ['Heure', 'Arrivées', 'Débit (pax/h)', 'Capacité/10min', 
                           'Traités', 'File attente', 'Attente (min)']
    df_display = df_display.round(1)
    st.dataframe(df_display, use_container_width=True, height=400)


#--- EXPORT ---
st.markdown("---")
if st.button("💾 Exporter les résultats en CSV"):
    csv = df_sim.to_csv(index=False)
    st.download_button(
        label="Télécharger CSV",
        data=csv,
        file_name=f"simulation_{selected_site}_{pd.Timestamp(selected_jour).strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
