# -*- coding: utf-8 -*-
"""
Created on Wed Feb  7 13:52:23 2024

@author: DEMANET
"""

##### Version en cours
### ajout de la possbilité de selectionner une période


import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import altair as alt
from itertools import product
import locale
from datetime import datetime, timedelta

###############
#Chargement du fichier seuil par défaut
###############
@st.cache_data
def charger_df_seuils(chemin="Seuils.xlsx"):
    df = pd.read_excel(chemin)
    df["site"] = df["site"].astype(str).str.strip()
    return df

###############
#Upload optionnel du fichier seuil
###############

fichier = st.sidebar.file_uploader("Charger un fichier de seuils (optionnel)", type=["xlsx"])
if fichier is not None:
    # Si fichier téléversé:
    df_seuils = pd.read_excel(fichier)
    df_seuils["site"] = df_seuils["site"].astype(str).str.strip() 
    st.sidebar.success("Seuils personnalisés chargés ✅") 
    st.sidebar.success("Seuils personnalisés chargés ✅")
else:
    # Sinon: fichier par défaut
    df_seuils = charger_df_seuils() 
    st.sidebar.info("Seuils par défaut (GitHub)")

###############
#Dictionnaire utilisé par la fonction seuil()
###############
SEUILS = dict(zip(df_seuils["site"], df_seuils["seuil"]))

###############
#Affichage du tableau des seuils dans la barre latérale
###############
st.sidebar.subheader("📋 Seuils en vigueur") 
st.sidebar.dataframe(df_seuils, use_container_width=True, hide_index=True)

def seuil(site): 
    return SEUILS.get(str(site).strip(), 0)




def main(): 
    st.set_page_config(page_title="Vérif Seuil", page_icon="📊", layout="centered", initial_sidebar_state="auto", menu_items=None)
    
    hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True) 
    st.title("📊 Vérif Seuil - Charge horaire")
    st.divider()
    uploaded_file = st.file_uploader("Choisir un fichier :", type=["xls", "xlsx"])
    
    #uploaded_file = "C:/Users/demanet/Downloads/export_pif_du_2024-07-04_au_2024-07-14.xlsx"
    
    
    if uploaded_file is not None: 
        
        df = pd.read_excel(uploaded_file)
        
       

# création d'un dataframe contenant toutes les combinaisons jour/heure/site
        jours= df['jour'].unique()
        heures = pd.date_range("00:00:00", "23:50:00", freq="10min").strftime('%H:%M:%S')
        sites= df['site'].unique()
    
        combinaisons = pd.DataFrame(list(product(jours,heures,sites)), columns=['jour', 'heure','site'])
       
# jointure entre le df_final et le df combinaisons pour corriger le problème d'omission de colonne lorsque charge = 0
       
        df_complet =pd.merge(df, combinaisons, on = ['jour', 'heure','site'], how = "right")
        df_complet['charge'].fillna(0, inplace=True)
       
        df = df_complet
        df['jour']= pd.to_datetime(df['jour'])
        df = df.sort_values(by=['site', 'jour', 'heure'])

#transforme le dataframe dans un format "large"        
        
        df_pivot = df.pivot (index=['site', 'jour'], columns='heure', values='charge').reset_index() 
        
        
# calcul du cumul de charge horaire glissant à partir de données 10 min
        
        def calculate_sums(row): 
            sums = {}
            times = df_pivot.columns[2:]
            for i in range(len(times)-5):
                sum_value=sum(row[times[j]] for j in range (i,i+6))
                sums[times[i+5]] = sum_value
            return pd.Series(sums)
        
        new_df = df_pivot.apply (calculate_sums, axis=1)
            
        new_df.insert(0, 'site', df_pivot['site'])
        new_df.insert(1, 'jour', df_pivot['jour'])        
    
    # retourne le dataframe dans un format "long"
    
        df_depivote = pd.melt(new_df, id_vars=['site', 'jour'], var_name='heure', value_name='charge')

         
                                  
#new_df.to_excel("C:/Users/demanet/Downloads/test_cumul.xlsx", index= False)         

#créer une liste dynamique des sites (revient à faire unique(), mais permet ici de supprmier les sites fermées)
        df_config =df.groupby(['site'], as_index=False)['charge'].sum()
        df_config =df_config[df_config['charge'] > 0]
        df_config = df_config.drop(columns=['charge'])
        
                  
        custom_dict = {'K CTRCNT' : 0,'K CTR' : 1,'K CNT' : 2 , 'L CTR' : 3, 'L CNT' : 4, 'M CTR' : 5, 'Galerie EF' : 6, 'C2F' : 7, 'C2G' : 8, 'Liaison AC' : 9,'Liaison BD' : 10,
        'T3': 11, 'Terminal 1' : 12, 'Terminal 1_5' : 13, 'Terminal 1_6' : 14}
        
        df_config = df_config.sort_values(by=['site'], key=lambda x: x.map(custom_dict)).reset_index(drop=True)
    
        
        
        
        
        
        sites = df_config['site'].unique() # permet de créer un tableau simple à partir du dataframe df_complet (à revoir)
        
        
    # seuils de saturation    
        
        
        def seuil(site):
            #seuils = {'K CTRCNT' : 0,'K CTR' : 1900,'K CNT' : 300 , 'L CTR' : 2240, 'L CNT' :  1520, 'M CTR' : 1960, 'Galerie EF' : 1820, 'C2F' : 2400, 'C2G' : 700, 'Liaison AC' : 1960,'Liaison BD' : 2460,
            #'T3': 1260, 'Terminal 1' : 2580, 'Terminal 1_5' : 390, 'Terminal 1_6' : 520, 

            #'2E_Arr' :3164 , '2E_Dep' : 4251 ,'Galerie E > F' : 2272 , 'Galerie F > E' : 1314, 'F > S3' :  2422 , 'S3 > F': 892 ,  '2G_Emport' : 1314, 
                     #'AC_Dep' : 2848, 'AC_Arr' :  4152, 'C_Arr' :  1808, 'BD_Arr' : 1993, 'BD_Dep' : 1544, 'T1_Arr' : 2478, 'T1_Dep' : 2518,'T3_Arr' : 1140, 'T3_Dep' : 1374}
        
            #return seuils.get(site,0)
            return SEUILS.get(str(site).strip(), 0)





####### filtre le df sur la semaine suivante entière
        on = st.toggle("Mode auto (désactiver pour sélectionner une semaine à tracer)",value=True)   
         
        if on : 
                jour_deb = df['jour'].min().weekday()
                jour_a_ajouter = (7 - jour_deb) % 7
                deb_semaine_deux= df['jour'].min() + pd.Timedelta(days=jour_a_ajouter)
               
                fin_semaine_deux = deb_semaine_deux+ pd.Timedelta(days=6)
               
                df_depivote = df_depivote[(df_depivote['jour'] >= deb_semaine_deux) & (df_depivote['jour']<= fin_semaine_deux)]
                
              
                # df['jour']= pd.to_datetime(df['jour'])
                
                for site in sites: 
                    
                    st.subheader(f" {site}")
                    
                    st.write("seuil max: ", seuil(site))
                    df_site = df_depivote.loc[df_depivote['site']==site]
                    locale.setlocale(locale.LC_ALL, 'fr_FR')
                    df_site['Date']= df_site['jour'].dt.strftime('%A %d %b')
                   
                    fig = px.line(df_site, x= 'heure', y= 'charge',  color = 'Date',
                                    labels={'jour', 'date'})
                    
                    ligne_seuil = seuil(site)
                    fig.add_hline(y=ligne_seuil, line_dash='dash', line_color="red")
                    
                    
                    
                    st.plotly_chart(fig)
    
        
  
        else:
    
    
    ## mode auto pour J-7 ou mode manuel
    
            col1, col2 = st.columns(2)
            with col1:
                  debut = st.date_input("Date de début :",min_value=jours.min(),max_value= jours.max() - timedelta(days=1), key=10)
            with col2:    
                  fin = st.date_input("Date de fin :",value=min( pd.to_datetime(debut + timedelta(days=6)), jours.max()- timedelta(days=1) ), min_value = jours.min(), max_value = min( pd.to_datetime(debut + timedelta(days=6)), jours.max()- timedelta(days=1) ) ,help=" 7 jours max.", key=2) #-1 car il y a toujours quelques heures du début de la dernière journée (méthode à revoir)

            # start_date = pd.to_datetime(debut)
            # end_date = pd.to_datetime(fin)
            #df_depivote['jour']= pd.to_datetime(df['jour'])
            df_depivote = df_depivote[(df_depivote['jour'] >= pd.to_datetime(debut)) & (df_depivote['jour']<= pd.to_datetime(fin))]


# graphique   
            if st.button('Tracer les flux sur la période choisie'):
                
                for site in sites: 
                    
                    st.subheader(f" {site}")
                    
                    st.write("seuil max: ", seuil(site))
                    df_site = df_depivote.loc[df_depivote['site']==site]
                    locale.setlocale(locale.LC_ALL, 'fr_FR')
                    df_site['Date']= df_site['jour'].dt.strftime('%A %d %b')
                   
                    fig = px.line(df_site, x= 'heure', y= 'charge',  color = 'Date',
                                    labels={'jour', 'date'})
                    
                    ligne_seuil = seuil(site)
                    fig.add_hline(y=ligne_seuil, line_dash='dash', line_color="red")
                    
                    
                    
                    st.plotly_chart(fig)

                    
        

 
       
if __name__=="__main__":
    main()
    





