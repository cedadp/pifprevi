import pandas as pd
import numpy as np
import datetime
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import streamlit as st

###V2 - Dans cette version le traitement des heures au format homogène entre AF et SARIA est réalisé dans Concat et non plus dans Prévis

#uploaded_file = "C:/Users/demanet/Downloads/Prévisions d'activité Sem 14-22 du 02.04.2024 (1).xlsx" # utilisé pour débogage en local
#uploaded_file2 = "C:/Users/demanet/Downloads/Prévis_PIF_J-7_2024-04-02.xlsx" # utilisé pour débogage en local


           
st.set_page_config(page_title="Concat", page_icon="📦", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

download = False # stop le rerun de l'app

######### Input #########

#   Noms des feuilles, peut changer dans le temps si qqn le modifie
st.title("📦 Concat")
name_sheet_cies = "pgrm_cies"
name_sheet_af = "Programme brut"
st.divider()
st.subheader("Prévision activité AF :")

def get_dispatch_sat_T1(sat):
    df = pd.read_excel("fichier_config_PIF.xlsx", sheet_name="dispatch_sat")
    df = df.fillna("XXXXX")
    return list(df[sat])


uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None and download is False:
    @st.cache_data(ttl=60)
    def previ_af():
        with st.spinner('Chargemement prévision AF 1 ...'):
            df_af_1 = pd.read_excel(uploaded_file,name_sheet_af,usecols=['A/D', 'Cie Ope', 'Num Vol', 'Porteur', 'Type Avion', 'Prov Dest', 'Affectation',
                        'Service emb/deb', 'Local Date', 'Semaine', 
                        'Jour', 'Scheduled Local Time 2', 'Plage',  
                        'Pax LOC TOT', 'Pax CNT TOT', 'PAX TOT'])
            df_af_1.rename(columns = {'Type Avion':'Sous-type avion'}, inplace = True)
            df_af_1['Service emb/deb'] = np.where((df_af_1["A/D"]=="D") & (df_af_1["Affectation"]=="F"), 'F', df_af_1['Service emb/deb'])
            df_af_1 = df_af_1.rename(columns={"Jour":"Jour (nb)",
                                    "Service emb/deb":"Libellé terminal",
                                    "Scheduled Local Time 2":"Horaire théorique"})
        return df_af_1
    
    df_af_1 = previ_af()
    st.success("Prévision AF 1 chargée !")

    st.subheader("Prévision activité ADP :")
    uploaded_file2 = st.file_uploader("Choisir un fichier :", key=3)
    if uploaded_file2 is not None:
        @st.cache_data(ttl=90)
        def previ_adp():
            with st.spinner('Chargemement prévision ADP ...'):
                df_cies_1 = pd.read_excel(uploaded_file2)
                df_cies_1.rename(columns={"sens":"A/D",
                                "Jour":"Local Date",
                                "Nombre de passagers prévisionnels":"PAX TOT",
                                "Terminal_format_saria":"Terminal_corrigé",
                                "Numéro de vol":"Num Vol",
                                "Code IATA compagnie":"Cie Ope",
                                "Code aéroport IATA proche":"Prov Dest"},
                                inplace = True)
                df_cies_1["Pax LOC TOT"] = 0
                df_cies_1["Pax CNT TOT"] = 0
                df_cies_1 = df_cies_1.rename(columns={"Terminal_corrigé":"Libellé terminal"})
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2B","Terminal 2B")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2D","Terminal 2D")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2A","Terminal 2A")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2C","Terminal 2C")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2E","Terminal 2E")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2F","Terminal 2F")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2G","Terminal 2G")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C1","Terminal 1")
                # df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C1","T1_Inter")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("CT","Terminal 3")
            
            return df_cies_1

        df_cies_1 = previ_adp()
        st.success("Prévisions ADP chargées !")


        st.markdown('--------------------')
        ######### Traitement #########



        ######### Gestion des dates #########

        min_date_previ = min(df_af_1['Local Date']) 
        max_date_previ = max(df_af_1['Local Date']) 
        min_date_adp = min(df_cies_1['Local Date'])
        max_date_adp = max(df_cies_1['Local Date'])

        st.warning("Plage des programmes AF/Skyteam : du " + str(min_date_previ.date()) + " au " + str(max_date_previ.date()))
        st.warning("Plage du programme ADP : du " + str(min_date_adp.date()) + " au " + str(max_date_adp.date()))

        if min_date_adp <= min_date_previ and max_date_adp >= max_date_previ:
            st.warning("Prévision d'activité est limitant")

            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] >= min_date_previ) & (df_cies_1['Local Date'] <= max_date_previ)]
            
        elif min_date_adp >= min_date_previ and max_date_adp <= max_date_previ:
            st.warning("Réalisé d'activité est limitant")
            
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] >= min_date_adp)]
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] <= max_date_adp)]
            
        elif min_date_adp >= min_date_previ and max_date_adp >= max_date_previ and max_date_previ >= min_date_adp:
            st.warning("Programme ADP et AF 2 limitant")
            
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] >= min_date_adp)]
            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] <= max_date_previ)]

        elif min_date_adp <= min_date_previ and max_date_adp <= max_date_previ and max_date_adp >= min_date_previ:
            st.warning("Programme AF 1 et ADP limitant")
            
            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] >= min_date_previ)]
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] <= max_date_adp)]
            
        else:
            st.warning("Les programmes AF/ADP ne se recouvrent pas, impossible de continuer"
                                    + "\n Veuillez sélectionner des programmes d'activités compatibles")

        placeholder = st.empty()

        #######################################################################
        term_adp = ["Terminal 2E", "Terminal 2G", "Terminal 2F"]   #liste les terminaux périmètre Air France
        

        ######### Traitement #########

        df_cies_1 = df_cies_1[~(df_cies_1["Libellé terminal"].isin(term_adp) == True)] # supprime les terminaux Air France des prévis périmètre ADP

        ######### Def #########

        placeholder.success("Mise en forme des prévisions faite !")
        placeholder.info("Préparation à la concaténation des prévisions ...")
        placeholder.info("Récupération des champs vides ...")
        df_concat = pd.concat([df_af_1, df_cies_1])
        df_concat.reset_index(drop=True, inplace=True)
        df_pgrm_concat = df_concat.copy() # inutile pour le moment
        df_pgrm_concat['Plage'] = df_pgrm_concat['Plage'].fillna(value = "P4")

        #   A automatiser car ne prend pas toutes les cies en compte, ex ici c'est RC
        df_pgrm_concat = df_pgrm_concat.dropna(subset=['Pax LOC TOT']) # On supprime les vols où PAX LOC est vide car vol type maintenance avion sans px ?

        df_pgrm_concat['Libellé terminal'].loc[(df_pgrm_concat['Cie Ope'] == 'RC')] = 'Terminal 2D'
        #df_nan['Plage'] = df_nan['Plage'].fillna(value = "P4")

        #         36% est le nomre moyen de corres pour prévision activité AF
        #df_pgrm_concat['Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int')
        # si LOC TOT égal 0 ou vide on calcul LOC TOT  = PAX TOT * 0,64
        #df_pgrm_concat.loc[(df_pgrm_concat['Pax LOC TOT'].isna()) , 'Pax LOC TOT'] = (df_pgrm_concat['PAX TOT']*(1-0.36)).astype('int')
        #df_pgrm_concat.loc[(df_pgrm_concat['Pax LOC TOT'] == 0) , 'Pax LOC TOT'] = (df_pgrm_concat['PAX TOT']*(1-0.36)).astype('int')
        #df_pgrm_concat.loc[(df_pgrm_concat['Pax LOC TOT'].isna()) , 'Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int') # Si PAX LOC est vide on fait qqchose (le problème c'est qu'on a supprimé ces lignes juste avant)
        df_pgrm_concat['Pax CNT TOT'] = 0 # On met toute la colonne CNT à 0 car on va la recalculer dessous
       

        df_pgrm_concat.loc[df_pgrm_concat['Num Vol'] == 'MNE', 'Cie Ope'] = 'ZQ'
       
        df_pgrm_concat.loc[df_pgrm_concat['Pax LOC TOT'] != 0, 'Pax CNT TOT'] = df_pgrm_concat['PAX TOT'] - df_pgrm_concat['Pax LOC TOT'] # On recalcule (lorsque LOC TOT différent de 0 ) CNT = PAX TOT - LOC TOT pour avoir l'égalité PAX totaux = LOC + CNT même si erreur dans fichier
        
        sat5 = get_dispatch_sat_T1("sat5")
        sat6 = get_dispatch_sat_T1("sat6")
        df_pgrm_concat.loc[df_pgrm_concat['Cie Ope'].isin(sat6), 'Libellé terminal'] = 'Terminal 1_6'
        df_pgrm_concat.loc[df_pgrm_concat['Cie Ope'].isin(sat5), 'Libellé terminal'] = 'Terminal 1_5'
        
       
        def check_datetime_validity(cell):
            cell_str = str(cell)
            cell_split_space = cell_str.split(' ')
            cell_split_heure = cell_split_space[-1].split(':')
            try:
                return datetime.time(hour=int(cell_split_heure[0]), minute=int(cell_split_heure[1]), second=0)
            except ValueError:
                return np.nan
                
        from copy import deepcopy 
        
        # df_test = deepcopy(df_pgrm_concat)
        df_pgrm_concat['Horaire théorique'] = df_pgrm_concat['Horaire théorique'].apply(lambda x: check_datetime_validity(x))
        df_pgrm_concat = df_pgrm_concat.dropna(subset=['Horaire théorique'], axis=0).reset_index(drop=True)  #supprime les lignes avec une heure nan, défnit par la fonction ci dessus lorsque heure = 25:40 par exemple
 
        df_pgrm_concat['Horaire théorique'] = pd.to_datetime(df_pgrm_concat['Horaire théorique'], format='%H:%M:%S').dt.time 
       
       # df_pgrm_concat.to_excel("C:/Users/demanet/Downloads/programme_concat_compare6.xlsx", index= False)
        
        # à ajouter : df_pgrm_concat.dropna(inplace=True)
        placeholder.success("Concaténation des prévisions réussie !")

        CIE = "Cie Ope"
        TERMINAL = "Libellé terminal"   

        
               # Initialisation : une seule fois par session ->
        if "pgrm_source" not in st.session_state: st.session_state["pgrm_source"] = df_pgrm_concat.copy(deep=True)
        if "pgrm_resultat" not in st.session_state: st.session_state["pgrm_resultat"] = df_pgrm_concat.copy(deep=True) 
        source = st.session_state["pgrm_source"]
        st.subheader("Modificaiton des affectations")
        
               #Récapitulatif des combinaisons d'affectation ->
        recap = ( source.groupby( [CIE, TERMINAL], dropna=False, observed=True, sort=False, ) .size() .reset_index(name="Nombre de lignes") )
        st.markdown("Combinaisons présentes dans les données d’origine") 
        st.dataframe(recap, hide_index=True)
        
               
        nb_terminaux = ( source.loc[source[CIE].notna()] .groupby(CIE, observed=True, sort=False)[TERMINAL] .nunique(dropna=False) )
        cies_modifiables = nb_terminaux[nb_terminaux == 1].index
        table_edition = ( source.loc[ source[CIE].isin(cies_modifiables), [CIE, TERMINAL], ] .drop_duplicates(subset=[CIE]) .reset_index(drop=True) )

        table_edition[TERMINAL] = table_edition[TERMINAL].astype("string")
        compagnies = source[CIE].dropna().drop_duplicates().tolist()


        with st.form("form_gestion_pgrm"): 
             st.markdown("Modifier les terminaux")
             st.caption( "Seules les compagnies ayant un unique terminal distinct " "dans les données d'origine apparaissent ici." )
             edition = st.data_editor( table_edition,hide_index=True,disabled=[CIE],num_rows="fixed",column_config={TERMINAL: st.column_config.TextColumn( "Libellé terminal", help="Saisissez le nouveau libellé du terminal.",  ) },key="pgrm_editeur_terminaux",)

             st.markdown("**Exclure des compagnies**")

             tout_exclure = st.checkbox("Tout exclure — le dataset résultant sera vide",key="pgrm_tout_exclure",)

             cies_exclues = st.multiselect(    "Compagnies à exclure",    options=compagnies,    help="Cette sélection est ignorée si « Tout exclure » est coché.",    key="pgrm_cies_exclues",)

             appliquer = st.form_submit_button("Appliquer les modifications")


        if appliquer: resultat = source.copy(deep=True)
        # Repérer uniquement les libellés réellement modifiés.
        ancien = table_edition[TERMINAL].astype("string")
        nouveau = edition[TERMINAL].astype("string")

        identiques = (    ancien.eq(nouveau).fillna(False)    | (ancien.isna() & nouveau.isna()))

        modifications = edition.loc[~identiques, [CIE, TERMINAL]].copy()

        # Refuser un nouveau libellé vide ->
        valeurs_invalides = (    modifications[TERMINAL].astype("string")    .fillna("")    .str.strip()    .eq(""))

        if valeurs_invalides.any():
          st.error(          "Un nouveau libellé ne peut pas être vide. "        "Les changements n'ont pas été appliqués."    )
        else:
         if not modifications.empty:
        # Permet aussi de modifier une colonne initialement catégorielle.
          resultat[TERMINAL] = resultat[TERMINAL].astype(object)

        for compagnie, terminal in modifications.itertuples(
            index=False, name=None
        ):
            resultat.loc[
                resultat[CIE].eq(compagnie).fillna(False),
                TERMINAL,
            ] = terminal

       if tout_exclure:
        resultat = resultat.iloc[0:0].copy()
       else:
        resultat = resultat.loc[
            ~resultat[CIE].isin(cies_exclues)
        ].copy()

       st.session_state["pgrm_resultat"] = resultat
       st.success("Modifications appliquées.")

       df_pgrm_concat = st.session_state["pgrm_resultat"].copy(deep=True)
       st.markdown("Dataset après application des modifications") st.write(f"{len(df_pgrm_concat):,} lignes conservées.") st.dataframe(df_pgrm_concat.head(100), hide_index=True)













               
        ######### Export PGRM CONCAT ########      
        from datetime import datetime
        placeholder.info("Préparation à l'export du programme complet ...")
        directory_concat = "pgrm_complet_" + str(datetime.now())[:10] + ".xlsx"
        df_pgrm_concat.to_excel(directory_concat, sheet_name = "pgrm_complet")
        placeholder.success("Programme complet exporté !")
        placeholder.info("Fin du traitement")
        
        import io
        from pyxlsb import open_workbook as open_xlsb

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write each dataframe to a different worksheet.
            df_pgrm_concat.to_excel(writer, sheet_name= "pgrm_complet")
            # Close the Pandas Excel writer and output the Excel file to the buffer
            writer.close()

            st.download_button(
            label="Télécharger fichier Programme complet",
            data=buffer,
            file_name=directory_concat,
            mime="application/vnd.ms-excel"
            )
        

            download = True
        # st.markdown('<a href="/" target="_self">Revenir à l\'Accueil</a>', unsafe_allow_html=True)
        # st.markdown('<a href="/Pif_Previ_" target="_self">Aller directement à l\'outils Pif prévi</a>', unsafe_allow_html=True)

