import pandas as pd  
import streamlit as st
import numpy as np
import datetime
from functools import reduce
import time as tm
import openpyxl


# Configuration Streamlit

st.set_page_config(page_title="EquiPif", page_icon="👩‍✈️", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

#########################

st.title('👩‍✈️ EquiPif')
st.subheader("Programme complet :")

# Import

uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache_data(ttl=90)
    def df():
        with st.spinner('Chargemement Programme complet ...'):
            df = pd.read_excel(uploaded_file, "pgrm_complet")
            sat5 = ['FI', 'LO', 'A3', 'SK', 'DY', 'D8', 'GQ', 'S4']
            sat6 = ['LH', 'LX', 'OS', 'EW', 'SN']
            df.loc[df['Cie Ope'].isin(sat6), 'Libellé terminal'] = 'Terminal 1_6'
            df.loc[df['Cie Ope'].isin(sat5), 'Libellé terminal'] = 'Terminal 1_5'

            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_Inter","Terminal 1")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_5","Terminal 1_5")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_6","Terminal 1_6")
        st.success("Programme complet chargée !")
        return df

    df_pgrm = df()         
    start_all = tm.time()
    l_date = pd.to_datetime(df_pgrm['Local Date'].unique().tolist()).date
    l_date = sorted(l_date)

    # Merge utile pour le traitement

    table_faisceau_iata = pd.read_excel('table_faisceau_IATA.xlsx')
    table_faisceau_iata = table_faisceau_iata[['Prov Dest','Faisceau géographique']]
    df_pgrm = df_pgrm.merge(table_faisceau_iata,how='left', left_on='Prov Dest', right_on='Prov Dest')
    
    eff_type_avion = pd.read_excel('effectif_type_avion.xlsx')
    eff_type_avion = eff_type_avion.rename(columns={"TypeAvion":"Sous-type avion"})
    eff_type_avion['Sous-type avion'] = eff_type_avion['Sous-type avion'].astype(str)
    df_pgrm['Sous-type avion'] = df_pgrm['Sous-type avion'].astype(str)

    df_pgrm = df_pgrm.merge(eff_type_avion,how='left')
    df_pgrm = df_pgrm.loc[(df_pgrm['Cie Ope'] != 'AF') & (df_pgrm['Cie Ope'] != 'EC')]

    #Gestion des effectifs nul

    df_pgrm['Effectif'].fillna(8, inplace=True)
    nan_eff = df_pgrm[~df_pgrm['Effectif'].notnull()]
    # st.write(list(nan_eff['Sous-type avion'].unique()))

    st.subheader("Choix de la période :")
    col1, col2 = st.columns(2)
    with col1:
        debut = st.date_input("Date de début :", key=10)
    with col2:    
        fin = st.date_input("Date de fin :", key=2)
    
    start_date = pd.to_datetime(debut)
    end_date = pd.to_datetime(fin) 

    if st.button('Créer Export EquiPif'):

        st.warning('La requête a bien été prise en compte, début du traitement.\nNe tentez pas de fermer la fenêtre même si celle-ci semble figée')

    #        Entre pgrm ADP et pgrm AF les heures ne sont pas au même format. On les transforme ici. A terme migrer cette fonction dans Concat
        def STR_TO_DT(df):
            df_temp = df
            l_dt = []
            for t in range(df.shape[0]):
                TSTR =  str(df['Horaire théorique'][t])
                if len(TSTR)<10:
                    l = [int(i) for i in TSTR.split(':')]
                    l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
                else:
                    TSTR = TSTR[10:]
                    l = [int(i) for i in TSTR.split(':')]
                    l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
            
            df['Horaire théorique'] = l_dt
               
            return df_temp

        df_pgrm.reset_index(inplace=True)
        df_pgrm = STR_TO_DT(df_pgrm)
        df_pgrm_dt = df_pgrm.loc[(df_pgrm['Local Date'] >= start_date) & (df_pgrm['Local Date'] <= end_date)]
        df_pgrm_dt.reset_index(inplace=True, drop=True)
        df_pgrm_dt['Unnamed: 0'] = df_pgrm_dt.index
        del df_pgrm_dt['Unnamed: 0']
        df_pgrm_dt = df_pgrm_dt.drop_duplicates()

        df_pgrm_dt = df_pgrm_dt.sort_values(by=['Local Date', 'Horaire théorique'])


        df_pgrm_dt['Porteur'].loc[df_pgrm_dt['PAX TOT'] < 170] = 'MP'
        df_pgrm_dt['Porteur'].loc[df_pgrm_dt['PAX TOT'] >= 170] = 'GP'
        df_pgrm_dt_GP = df_pgrm_dt.loc[df_pgrm_dt['Porteur'] == 'GP']
        df_pgrm_dt_MP = df_pgrm_dt.loc[df_pgrm_dt['Porteur'] == 'MP']

        #df_pgrm_dt = df_pgrm_dt.drop_duplicates(subset=['Cie Ope', 'Prov Dest', 'Local Date'], keep='first')

        L = ['Terminal 2D', 'Terminal 2B', 'Terminal 3', 'Terminal 1', 'Terminal 1_5', 'Terminal 1_6']
        L1 = ['Terminal 1', 'Terminal 1_5', 'Terminal 1_6']

        df_pgrm_dt_MP=pd.concat([df_pgrm_dt_MP.loc[(df_pgrm_dt_MP['Libellé terminal'].isin(L))],
        df_pgrm_dt_MP.loc[~df_pgrm_dt_MP['Libellé terminal'].isin(L)]
        .drop_duplicates(subset=['Cie Ope', 'Prov Dest', 'Local Date'],keep='first')]).sort_index()

        df_pgrm_dt_MP.reset_index(inplace=True)

        del df_pgrm_dt_MP['level_0']
        del df_pgrm_dt_MP['index']

        df_pgrm_dt_MP.drop_duplicates(inplace=True)


             
        df_pgrm_dt_GP['Horaire théorique'] = pd.to_datetime(df_pgrm_dt_GP['Horaire théorique'],format='%H:%M:%S')
        df_pgrm_dt_GP['Horaire théorique'] = df_pgrm_dt_GP['Horaire théorique'].round("10min").dt.strftime('%H:%M:%S')
        df_pgrm_dt_GP['Horaire théorique'] = pd.to_datetime(pd.to_datetime(df_pgrm_dt_GP['Horaire théorique']) 
                    - pd.Timedelta(minutes=90),format='%Y-%m-%d %H:%M:%S').round("10min").dt.strftime('%H:%M:%S')


        del df_pgrm_dt_GP['index']
        df_pgrm_dt_GP.drop_duplicates(inplace=True)
        
       

        from datetime import datetime, timedelta


        # Gestion de la date

        df_pgrm_dt_MP['Horaire théorique'] = pd.to_datetime(df_pgrm_dt_MP['Horaire théorique'],format='%H:%M:%S')
        df_pgrm_dt_MP['Horaire théorique'] = df_pgrm_dt_MP['Horaire théorique'].round("10min").dt.strftime('%H:%M:%S')

        # Ici on affecte le temps de presentation au PIF en focntion du Porteur et du Terminal
        # TO DO : à refacto

        

        df_pgrm_dt_MP['Horaire théorique'].loc[(df_pgrm_dt_MP['Libellé terminal'].isin(L1))] = pd.to_datetime(pd.to_datetime(df_pgrm_dt_MP['Horaire théorique'].loc[(df_pgrm_dt_MP['Libellé terminal'].isin(L1))]) - pd.Timedelta(minutes=75),format='%Y-%m-%d %H:%M:%S').round("10min").dt.strftime('%H:%M:%S')
        df_pgrm_dt_MP['Horaire théorique'].loc[~(df_pgrm_dt_MP['Libellé terminal'].isin(L1))] = pd.to_datetime(pd.to_datetime(df_pgrm_dt_MP['Horaire théorique'].loc[~(df_pgrm_dt_MP['Libellé terminal'].isin(L1))]) - pd.Timedelta(minutes=50),format='%Y-%m-%d %H:%M:%S').round("10min").dt.strftime('%H:%M:%S')
        
        df_pgrm_dt = pd.concat([df_pgrm_dt_MP, df_pgrm_dt_GP])

        df_pgrm_dt.reset_index(inplace=True)
        
        df_pgrm_dt = df_pgrm_dt.drop_duplicates(subset=df_pgrm_dt.columns.difference(['Unnamed: 0']))

            ### DISPATCH ###

        def DISPATCH_NEW(df):
            """Permet la création d'un DF dispatch qui facilite le tri par batterie de PIF"""
            col = ['Local Date', 'Horaire théorique', 'Prov Dest', 'A/D', 'Sous-type avion', 'Libellé terminal', 
                    'L CTR', 
                    'M CTR', 
                    'Galerie EF',
                    'C2F', 
                    'C2G', 
                    'Liaison AC', 
                    'Liaison BD', 
                    'T3',
                    'Terminal 1',
                    'Terminal 1_5',
                    'Terminal 1_6']

            dispatch_df = pd.DataFrame(columns = col)

            dispatch_df['Local Date'] = df['Local Date']
            dispatch_df['Horaire théorique'] = df['Horaire théorique']
            dispatch_df['Prov Dest'] = df['Prov Dest']
            dispatch_df['A/D'] = df['A/D']
            dispatch_df['Libellé terminal'] = df['Libellé terminal']
            dispatch_df['Sous-type avion'] = df['Sous-type avion']


            def dispatch(terminal):                
                temp = df.loc[df['Libellé terminal'] == terminal].copy()
                return temp['Effectif']


        #    L CTR
            dispatch_df['L CTR'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("EL")])

        #    M CTR
            dispatch_df['M CTR'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("EM")])

        #    Galerie EF
            dispatch_df['Galerie EF'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("EK")])

        #    LAC
            dispatch_df['Liaison AC'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 2A"),
                        dispatch("Terminal 2C")])

        #    LBD
            dispatch_df['Liaison BD'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 2B"),
                        dispatch("Terminal 2D")])

        #    T3
            dispatch_df['T3'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 3")])
            

        #    Terminal 1 Jonction
            dispatch_df['Terminal 1'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 1")])

            
        #    Terminal 1 Schengen (5 et 6)
            dispatch_df['Terminal 1_5'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 1_5")])
            
            dispatch_df['Terminal 1_6'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch("Terminal 1_6")])         



            dispatch_df.fillna(0, inplace=True)

            return dispatch_df


        dispatch = DISPATCH_NEW(df_pgrm_dt)

        dispatch = dispatch.loc[dispatch['A/D'] == 'D']

        dispatch = dispatch.groupby(by=['Local Date', 'Horaire théorique']).sum().reset_index()

        from scipy import signal 
            
        def CREATE_DF_SITE(dispatch_df, site):
            """Permet de créer le format de l'export pif final"""
            c = ['jour', 'heure', 'site', 'charge', 'type']
            l_pas10min = pd.date_range(datetime(2022,1,1), periods=144, freq="10T").time.tolist()
            df = pd.DataFrame(columns=c)
            l_jour = dispatch_df['Local Date'].sort_values(ascending = True).unique().tolist()
            nb_jour = len(l_jour)
            df['heure'] = l_pas10min * nb_jour
            df['site'] = site
            df['charge'] = 0
            df['type'] = "pifbi_python"
            for i in range(len(l_jour)):
                df.iloc[144*i:144*(i+1), 0] = pd.to_datetime(l_jour[i])

            df['heure'] = pd.to_datetime(df['heure'],format='%H:%M:%S').dt.strftime('%H:%M:%S')    
            return df
        
        site = ['L CTR', 
                'M CTR', 
                'Galerie EF', 
                'Liaison AC', 
                'Liaison BD', 
                'T3',
                'Terminal 1',
                'Terminal 1_5',
                'Terminal 1_6']


        L_df = []
        for i in site:
            temp_final = CREATE_DF_SITE(df_pgrm_dt, i)
            df_temp = dispatch[['Local Date', 'Horaire théorique', i]]
            df_temp.rename(columns={"Local Date": "jour", "Horaire théorique": "heure"}, inplace = True)
            temp_final['jour'] = pd.to_datetime(temp_final['jour'])
            df_final_temp = temp_final.merge(df_temp, on=['jour','heure'], how='left')
            df_final_temp['charge'] = df_final_temp[i]
            del df_final_temp[i]
            df_final_temp['charge'].fillna(0, inplace = True)
            L_df += [df_final_temp]
    
        x = pd.concat(L_df)
        

        directory_exp = "export_equipif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"

        st.info("Export EquiPIF créé avec succès !" + "\n\nPour lancer une nouvelle étude, lancer uniquement 'Choix de la période'")
        
        import io
        from pyxlsb import open_workbook as open_xlsb

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            x.to_excel(writer, sheet_name="export_equipif")
            writer.close()

            st.download_button(
            label="Télécharger fichier Export EquiPIF",
            data=buffer,
            file_name=directory_exp,
            mime="application/vnd.ms-excel"
            )

