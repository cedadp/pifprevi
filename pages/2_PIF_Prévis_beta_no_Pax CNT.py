import pandas as pd  
import streamlit as st
import numpy as np
import datetime
from functools import reduce
import time as tm
import openpyxl
from datetime import datetime, timedelta 
from io import BytesIO  
from pyxlsb import open_workbook as open_xlsb
### Version du 24 - 07 - 2024 ##########################################

###V2 - Dans cette version le traitement des heures au format homogène entre AF et SARIA est réalisé dans Concat et non plus dans Prévis
# uploaded_file = "C:/Users/demanet/Downloads/pgrm_complet_2024-08-28.xlsx"
# df = pd.read_excel(uploaded_file) 

st.set_page_config(page_title="PIF Prévis", page_icon="🛫", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

def download_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, sheet_name=name_output, index=False)
            writer.close()
            processed_data = output.getvalue()
            return processed_data



st.title('🛫 PIF Prévis')
st.subheader("Programme complet :")
uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache_data(ttl=90)
    def df():
        with st.spinner('Chargemement Programme complet ...'):
            df = pd.read_excel(uploaded_file, "pgrm_complet", converters={'Local Date': lambda x: pd.to_datetime(x, dayfirst=True, errors = "coerce")})
        # ajouter filtre T1
        # df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_Inter","Terminal 1")
        #df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_5","Terminal 1_5")
        # df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_6","Terminal 1_6")
        #partie déplacée dans Concat_V2
            st.success("Programme complet chargée !")
        return df

           
    df = df()

            
     # Si le fichier est un "REPLAY", on applique la logique de calcul
    if 'REPLAY' in uploaded_file.name:
        st.info("Fichier 'REPLAY' détecté. Calcul de  'Pax CNT TOT' théorique pour AM, KE, KL, LG, MF, MU.")
        coefficients = {
            'AM': 0.5, # coefficient pour la compagnie AM
            'KE': 0.5, # coefficient pour la compagnie KE
            'KL': 0.5, # coefficient pour la compagnie KL
            'LG': 0.5, # coefficient pour la compagnie LG
            'MF': 0.5, # coefficient pour la compagnie MF
            'MU': 0.5  # coefficient pour la compagnie MU
        }    

            # Conditions pour le calcul
        mask = (
                    (df['Pax CNT TOT'].isna()) &
                    (df['Affectation'].isin(['E', 'F', 'G'])) &
                    (df['A/D'] == 'A') &
                    (df['Cie Ope'].isin(list(coefficients.keys()))) &
                    (df['PAX TOT'].notna())
                )

               
            # Créer une série de coefficients et appliquer le calcul
        coeff_series = df['Cie Ope'].map(coefficients)
        df.loc[mask, 'Pax CNT TOT'] = df.loc[mask, 'PAX TOT'] * coeff_series[mask]
        
        st.success(f"Calcul REPLAY appliqué sur {mask.sum()} lignes")

            # Masque pour identifier les lignes à traiter (AF et DL)
        mask2 = (
                    (df['Pax CNT TOT'].isna()) &                            # Pax CNT TOT vide
                    (df['Affectation'].isin(['E', 'F', 'G'])) &            # Affectation E, F ou G
                    (df['A/D'] == 'A') &                                   # A/D = A
                    (df['Cie Ope'].isin(['AF', 'DL'])) &                   # Compagnies AF ou DL
                    (df['PAX TOT'].notna())                                # PAX TOT non vide
                )
                
        st.info(f"Lignes à traiter avec coefficients moyens par vol : {mask2.sum()}")
    

        
    
    








                
        #  afficher un aperçu des lignes modifiées
        if mask.sum() > 0:
            st.write("Aperçu des lignes modifiées :")
            st.dataframe(df[mask][['Cie Ope', 'Affectation', 'A/D', 'PAX TOT', 'Pax CNT TOT']])
            name_output = "dataset_replay"   
            excel_data = download_excel(df)
            st.download_button(
            label="📥 Télécharger le dataset modifié",
            data=excel_data,
            file_name=f'dataset_replay_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
    
    else:
        st.info("Fichier non-REPLAY chargé, aucun calcul spécifique appliqué")
        
   
    
            

    df_pgrm = df
    
    #start_all = tm.time()
    l_date = pd.to_datetime(df_pgrm['Local Date'].unique().tolist()).date
    l_date = sorted(l_date)

    def get_pif_in_fichier_config(pif):
        return pd.read_excel("fichier_config_PIF.xlsx", sheet_name=pif)
    
    def get_pif():
        df = pd.read_excel("fichier_config_PIF.xlsx", sheet_name="Config")
        df = df.fillna("XXXXX")
        return list(df['PIF'])

    L_pif = get_pif()

    table_faisceau_iata = pd.read_excel("table_faisceau_IATA (2).xlsx")
    table_faisceau_iata.rename(columns={"Code aéroport IATA":"Prov Dest"}, inplace=True)
    table_faisceau_iata = table_faisceau_iata[['Prov Dest','Faisceau géographique']]
    df_pgrm = df_pgrm.merge(table_faisceau_iata,how='left', left_on='Prov Dest', right_on='Prov Dest')
    df_pgrm['Faisceau géographique'].fillna('Autre Afrique', inplace=True)

    saison_IATA = st.radio(
    "Choix de la saison IATA",
    ["**Winter** :snowflake:", "**Summer** :sun_with_face:"],
    horizontal=True)

    def get_hyp_rep(pifs):
        return pd.read_excel("hyp_rep_V2.xlsx", sheet_name=pifs)

    def COURBE_PRES(t):
        if saison_IATA == "**Winter** :snowflake:":
            return pd.read_excel('courbes_presentation_Winter.xlsx', t)
        if saison_IATA == "**Summer** :sun_with_face:":
            return pd.read_excel('courbes_presentation_Summer.xlsx', t) 
        
    
    datemini = pd.to_datetime(l_date).min()
    datemaxi = pd.to_datetime(l_date).max()
    
    
    
    
    
    col1, col2 = st.columns(2)
    # with col1:
    #     debut = st.date_input("Date de début :", value = min ( pd.to_datetime(datemini + timedelta(days=2)), datemini) , min_value=datemini ,max_value=  datemaxi , key=10)
    # with col2:    
    #     fin = st.date_input("Date de fin :", value =  min ( pd.to_datetime(debut + timedelta(days=10)),datemaxi) ,   min_value = debut, max_value = datemaxi, key=2)
    with col1:
        debut = st.date_input("Date de début :" , min_value=datemini ,max_value=  datemaxi , key=10)
    with col2:    
        fin = st.date_input("Date de fin :",  value =  min ( pd.to_datetime(debut + timedelta(days=10)),datemaxi) , min_value = debut, max_value = datemaxi, key=2)
    start_date = pd.to_datetime(debut)
    end_date = pd.to_datetime(fin) 

    if st.button('Créer Export PIF'):
    
    #Fonction qui regroupe les sous fonctions de traitement
        
        st.warning('La requête a bien été prise en compte, début du traitement.\nNe tentez pas de fermer la fenêtre même si celle-ci semble figée.')
        placeholder = st.empty()
        my_bar2 = placeholder.progress(5)


    ### path files ###
        path_hyp = r"" + "hypotheses_repartition_correspondances.xlsx"
        name_hyp = "Feuil1"
        
        path_faisceaux = r"" + "faisceaux_escales.xlsx"
        name_faisceaux = "escales"
        
    #        ancienne courbes de prés globale, sans distinction de terminal
    #        path_courbes = r"" + source_outils_previ.chemin_fichier_source(4)
    #        name_courbes = "nouvellesCourbesPresentation"
        
        path_courbes_term = r"" + "nouvelles_courbes_presentation_PIF.xlsx"
        list_terminaux = ['Terminal 2A', 'Terminal 2B', 'Terminal 2C', 'Terminal 2D',
                        'EK', 'EL', 'EM', 'F', 'G', 'Terminal 3','Terminal 1',
                        'Terminal 1_5','Terminal 1_6']
        
        list_terminaux_P4 = ['Terminal 2A', 'Terminal 2B', 'Terminal 2C', 'Terminal 2D',
                        'Terminal 3','Terminal 1',
                        'Terminal 1_5','Terminal 1_6']
        
        path_output = r"" + "output_export_pif"
        name_output = "export_pif"
        


        
        def FAISCEAUX_IATA():
            df = pd.read_excel(path_faisceaux, name_faisceaux)
            del df['faisceau_facturation']
            del df['faisceau_commercial']
            del df['cl_long']
            del df['pays']
            del df['ville']
            del df['aeroport']
            del df['escale_OACI']
            del df['jour_ref']
            del df['statut']
            return df 
        
        df_faisceaux = FAISCEAUX_IATA()
        
        my_bar2.progress(10)
        
    #        Pour la courbe de pres unique, inutile
    #        def COURBE_PRESENTATION():
    #            return pd.read_excel(path_courbes, name_courbes)
        
    #        df_courbe_presentation = COURBE_PRESENTATION()
        
        
    # #        Entre pgrm ADP et pgrm AF les heures ne sont pas au même format. On les transforme ici. A terme migrer cette fonction dans Concat
    # déplacé dans Concat le 10/04/24
    #     def STR_TO_DT(df):
    #         df_temp = df
    #         l_dt = []
    #         for t in range(df.shape[0]):
    #             TSTR =  str(df['Horaire théorique'][t])
    #             if len(TSTR)<10:
    #                 try:
                        
    #                     l = [int(i) for i in TSTR.split(':')]
    #                     l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
    #                 except ValueError: 
    #                     st.write(f"Erreur horaire vol: {df.loc[t,'Cie Ope']}{df.loc[t,'Num Vol']}")
    #                     continue
    #             else:
    #                 TSTR = TSTR[10:] # on retire année mois jour
    #                 l = [int(i) for i in TSTR.split(':')]
    #                 l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
            
    #         df['Horaire théorique'] = l_dt
                
    #         return df_temp
        
    #     df_pgrm_dt = STR_TO_DT(df_pgrm)
        df_pgrm_dt = df_pgrm.loc[(df_pgrm['Local Date'] >= start_date) & (df_pgrm['Local Date'] <= end_date)]
        df_pgrm_dt.reset_index(inplace=True, drop=True)
        df_pgrm_dt['Unnamed: 0'] = df_pgrm_dt.index
        
    
        
        faisceaux = ['Métropole', 'Schengen', 'U.E. hors M & S', 'Afrique du Nord',
    'Amérique du Nord', 'Autre Afrique', 'Autre Europe', 'DOM TOM',
    'Extrême Orient', 'Moyen Orient', 'Amérique Centre + Sud']
        
            ### DISPATCH ###       


        import numpy
        from datetime import datetime, timedelta
        def HYP_REP(sheet):
            df = get_hyp_rep(sheet)
            # df['heure'] = pd.to_datetime(df['heure'].str[:8],format='%H.%M.%S')
            return df

        df_pgrm_dt['Horaire théorique'] = pd.to_datetime(df_pgrm_dt['Horaire théorique'],format='%H:%M:%S')
        # df_pgrm_dt['Horaire théorique'] = df_pgrm_dt["Horaire théorique"].apply(lambda x: x.hour)
        df_pgrm_dt = df_pgrm_dt.drop_duplicates(subset=df_pgrm_dt.columns.difference(['Unnamed: 0']))

        my_bar2.progress(15)

        def DISPATCH_NEW(df):
            """Permet la création d'un DF dispatch qui facilite le tri par batterie de PIF"""
            col = ['Local Date', 'Horaire théorique', 'Prov Dest', 'A/D', 'Libellé terminal', 'Faisceau géographique'] + L_pif

            dispatch_df = pd.DataFrame(columns = col, index = df['Unnamed: 0'])

            

            dispatch_df['Local Date'] = df['Local Date']
            dispatch_df['Horaire théorique'] = df['Horaire théorique']
            dispatch_df['Prov Dest'] = df['Prov Dest']
            dispatch_df['A/D'] = df['A/D']
            dispatch_df['Libellé terminal'] = df['Libellé terminal']
            dispatch_df['Faisceau géographique'] = df['Faisceau géographique']
            dispatch_df['Plage'] = df['Plage']
            dispatch_df['Pax LOC TOT'] = df['Pax LOC TOT']
            dispatch_df['Pax CNT TOT'] = df['Pax CNT TOT']
            dispatch_df['PAX TOT'] = df['PAX TOT']
            dispatch_df['Affectation'] = df['Affectation']
            dispatch_df['TOT_théorique'] = 0

            dispatch_df.loc[(dispatch_df['A/D'] == 'A') & (dispatch_df['Affectation'].isin(['E', 'F', 'G'])), 'TOT_théorique'] = dispatch_df['Pax CNT TOT']
            dispatch_df.loc[(dispatch_df['A/D'] == 'D') & (~dispatch_df['Affectation'].isin(['E', 'F', 'G'])), 'TOT_théorique'] = dispatch_df['PAX TOT']
            dispatch_df.loc[(dispatch_df['A/D'] == 'D') & (dispatch_df['Affectation'].isin(['E', 'F', 'G'])), 'TOT_théorique'] = dispatch_df['Pax LOC TOT']

            my_bar2.progress(20)

            def dispatch_term(terminal, salle_apport, salle_emport, AD):
                hyp_rep = HYP_REP(salle_apport + "_" + salle_emport)
                
                L_df = []
                for i,n in zip(hyp_rep['heure'].values, numpy.roll(hyp_rep['heure'].values,-1)):
                    for j in faisceaux:
                        x = hyp_rep.loc[(hyp_rep['heure'] == i)][j].tolist()[0] 
                        if AD == 'D':
                            x = 1 
                        if x != 0:
                            temp = df.loc[(df['A/D'] == AD) & (df['Libellé terminal'] == terminal)].copy()
                            temp = temp.loc[(df['Faisceau géographique'] == j)]
                            # temp = temp.loc[(temp['Horaire théorique'] >= i) & (temp['Horaire théorique'] < n) ]['Pax CNT TOT']*x
                            temp = temp.loc[temp['Plage']== i]['Pax CNT TOT']*x
                            L_df += [temp]

                return reduce(lambda a, b: a.add(b, fill_value = 0),L_df)

            
            my_bar2.progress(30)

            def dispatch_term_D(terminal, type_pax = 'PAX TOT'):                
                temp = df.loc[(df['A/D'] == 'D') & (df['Libellé terminal'] == terminal)].copy()
                return temp[type_pax]

            
            
            for pif in L_pif:
                dispatch = []
                df_config = get_pif_in_fichier_config(pif)
                
                for index, row in df_config.iterrows():
                    if row['Arr/Dep'] == 'D':
                        dispatch += [dispatch_term_D(row['terminal'], row['type_pax'])]
                    else:
                        dispatch += [dispatch_term(row['terminal'], row['salle_apport'], row['salle_emport'], row['Arr/Dep'])]

                    
                dispatch_df[pif] = reduce(lambda a, b: a.add(b, fill_value = 0),
                                        dispatch)  

            dispatch_df.fillna(0, inplace=True)

            dispatch_df['TOT_calcul'] = dispatch_df[L_pif].sum(axis=1)

            

            for i in L_pif:
                dispatch_df[i] = dispatch_df[i] / dispatch_df['TOT_calcul']*dispatch_df['TOT_théorique']

            dispatch_df.fillna(0, inplace=True)
            return dispatch_df    

        dispatch = DISPATCH_NEW(df_pgrm_dt)

        dispatch.to_excel("dispatch.xlsx", sheet_name="dispatch")
        
        my_bar2.progress(50)
        l_courbe_geo_t = {}
        
        for t in list_terminaux:
            df_courbe = COURBE_PRES(t).copy()
            l_courbe_geo_t[t] = {}
            for i in df_courbe["faisceau_geographique"].unique():
                temp = df_courbe.copy()
                temp = temp[temp["faisceau_geographique"].copy()==i].copy()
                l_courbe_geo_t[t][i] = {}
                for j in ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P0"]:
                    l_courbe_geo_t[t][i][j] = {}
                    # st.write(df_courbe)
                    # st.write(t)
                    # st.write(l_courbe_geo_t[t][i][j])
                    l_courbe_geo_t[t][i][j] = temp[j].tolist()
                    
        my_bar2.progress(55)


        global pb_index
        pb_index = 0


        dispatch_paf = dispatch.copy()

        my_bar2.progress(65)

        dispatch_paf['new_date'] = dispatch_paf['Local Date'].dt.date
        dispatch_paf['new_time'] = dispatch_paf['Horaire théorique'].dt.time
        dispatch_paf['new_datetime'] = pd.to_datetime(dispatch_paf['new_date'].astype(str) + ' ' + dispatch_paf['new_time'].astype(str))

        dispatch_paf.loc[(dispatch_paf['Libellé terminal'].isin(list_terminaux_P4)) & (dispatch_paf['Horaire théorique']>datetime(1900, 1, 1, 0, 00, 00, 0)), 'Plage'] = 'P2'
        dispatch_paf.loc[(dispatch_paf['Libellé terminal'].isin(list_terminaux_P4)) & (dispatch_paf['Horaire théorique']>datetime(1900, 1, 1, 11, 00, 00, 0)), 'Plage'] = 'P4'
        dispatch_paf.loc[(dispatch_paf['Libellé terminal'].isin(list_terminaux_P4)) & (dispatch_paf['Horaire théorique']>datetime(1900, 1, 1, 15, 00, 00, 0)), 'Plage'] = 'P5'
        dispatch_paf.loc[(dispatch_paf['Libellé terminal'].isin(list_terminaux_P4)) & (dispatch_paf['Horaire théorique']>datetime(1900, 1, 1, 17, 00, 00, 0)), 'Plage'] = 'P6'
        dispatch_paf.loc[(dispatch_paf['Libellé terminal'].isin(list_terminaux_P4)) & (dispatch_paf['Horaire théorique']>datetime(1900, 1, 1, 19, 00, 00, 0)), 'Plage'] = 'P7'

        my_bar2.progress(75)

        dispatch_paf_D = dispatch_paf.copy()
        dispatch_paf_D = dispatch_paf_D[dispatch_paf_D["A/D"] == "D"]
        dispatch_paf_A = dispatch_paf.copy()
        dispatch_paf_A = dispatch_paf_A[dispatch_paf_A["A/D"] == "A"]
        
        my_bar2.progress(85)

        n_D = 24
        n_A = 4 #len(L_A)

        # Create a list to store the duplicated rows
        rows = []
        L_A = [0, 0, 0.5, 0.5]

        my_bar2.progress(90)

        # DEPART
        # Loop through each row in the dataframe
        for index, row in dispatch_paf_D.iterrows():
            # Loop n times to duplicate the row and subtract 10 minutes from the datetime column each time
            for i in range(n_D):
                # Create a copy of the original row
                new_row = row.copy()
                if new_row['Faisceau géographique'] == 0:
                    x = "Extrême Orient"
                else:
                    x = new_row['Faisceau géographique']    
                L = l_courbe_geo_t[new_row['Libellé terminal']][x][new_row['Plage']]               
                # Subtract 10 minutes from the datetime column
                new_row['new_datetime'] -= timedelta(minutes=10*i)
                for pif in L_pif:
                    new_row[pif] = L[i]*new_row[pif]
                
                # Append the modified row to the list
                rows.append(new_row)
                
        my_bar2.progress(95)

        # Create a new dataframe from the list of duplicated rows
        new_df = pd.DataFrame(rows)


        # ARRIVER

        for index, row1 in dispatch_paf_A.iterrows():
            # Loop n times to duplicate the row and subtract 10 minutes from the datetime column each time
            for i in range(n_A):
                # Create a copy of the original row
                new_row = row1.copy()
                # Subtract 10 minutes from the datetime column
                new_row['new_datetime'] += timedelta(minutes=10*i)
                for pif in L_pif:
                    new_row[pif] = L_A[i]*new_row[pif]
                
                # Append the modified row to the list
                rows.append(new_row)
                
        my_bar2.progress(96)
        my_bar2.progress(97)

        # Create a new dataframe from the list of duplicated rows
        new_df_A = pd.DataFrame(rows)

        

        new_df_A['Local Date'] = new_df_A['new_datetime'].dt.date
        new_df_A['Horaire théorique'] = new_df_A['new_datetime'].dt.time

        df_final = pd.melt(new_df_A, id_vars=['new_datetime'], value_vars=L_pif)

        def ceil_dt(x):
            return x + (datetime.min - x) % timedelta(minutes=10)

        df_final['new_datetime'] = df_final['new_datetime'].apply(lambda x: ceil_dt(x))        
        df_final['Horaire théorique'] = df_final['new_datetime'].dt.time
        df_final['new_datetime'] = df_final['new_datetime'].dt.date

        df_final = df_final.groupby(['new_datetime', 'Horaire théorique', 'variable']).sum().reset_index()
        
        df_final.rename(columns={"new_datetime":"jour",
                        'Horaire théorique':'heure',
                        'variable':'site',
                        'value':'charge'}, inplace=True)
        
        my_bar2.progress(98)

        if "REPLAY" in uploaded_file.name.upper():
                    directory_exp = "export_pif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + "_REPLAY.xlsx"
        else:
                    directory_exp = "export_pif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"

        #directory_exp = "export_pif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"
            




     
        
        
        my_bar2.progress(100)

        processed_data = download_excel(df_final)
        st.download_button(
        label="Télécharger fichier Export pif",
        data=processed_data,
        file_name=directory_exp,
        mime="application/vnd.ms-excel"
        )
                        

        st.info("Export PIF créé avec succès !" + "\n\nPour lancer une nouvelle étude, lancer uniquement 'CHOISIR LES DATES'")
        


