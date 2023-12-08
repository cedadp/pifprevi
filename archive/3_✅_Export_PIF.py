import pandas as pd
import streamlit as st
import os
import time
import pandas as pd
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter
import datetime
import calendar
import locale
from openpyxl.styles import Font
import itertools
from datetime import datetime
locale.setlocale(locale.LC_ALL, "fr_FR")

st.title("✅ Macro final")
st.write("Macro du fichier Export_pif final")

def findDay(date):
    born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
    return (calendar.day_name[born])   
 
uploaded_file = st.file_uploader("Choose a file")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df1 = pd.DataFrame(columns=df.columns)
    wb= Workbook()
    writer = pd.ExcelWriter('multiple3.xlsx', engine='xlsxwriter')

    def clean(df,i):
        df['SOMME PAX LOCAUX DE LA JOURNEE'] = df.iloc[:, 1:145].sum(axis=1)
        df['Numéro de Jour'] = df['jour'].dt.day
        df['Date complète'] = df['jour'].dt.strftime('%d/%m/%Y')
        df['Jour de la semaine'] = df['jour'].dt.day_name(locale="fr_FR")     
        g = str(i).replace(" ", "_")
        df[str(i).replace(" ", "_")] = df['jour'].dt.month_name(locale="fr_FR")
        df["Jour férié ?"] = ""
        first_column = df.pop('Jour férié ?')
        df.insert(1, '"Jour férié ?', first_column)
        first_column = df.pop('Numéro de Jour')
        df.insert(1, 'Numéro de Jour', first_column)
        first_column = df.pop('Date complète')
        df.insert(3, 'Date complète', first_column)
        first_column = df.pop('Jour de la semaine')
        df.insert(3, 'Jour de la semaine', first_column)
        first_column = df.pop(str(i).replace(" ", "_"))
        df.insert(0, str(i).replace(" ", "_"), first_column)
        df.pop('jour')
        df[str(i).replace(" ", "_")] = list(itertools.chain.from_iterable([key] + [float('nan')]*(len(list(val))-1) 
                            for key, val in itertools.groupby(df[str(i).replace(" ", "_")].tolist())))

    
    def findDay(date):
        born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
        return (calendar.day_name[born])    


    import io
    from pyxlsb import open_workbook as open_xlsb
  
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        site = []
        for i in df.site.unique():
            name = str(i).replace(" ", "_")
            site += [name]
            name = df.copy()
            name = name[name['site'] == i]
            name = name.pivot_table(values='charge', index='jour', columns=['heure'], aggfunc='first')
            name.reset_index(inplace=True)
            name.fillna(0, inplace=True)
            clean(name,i)
            name.to_excel(writer, sheet_name=str(i).replace(" ", "_"), index=False)
        writer.close()

        st.download_button(
        label="Télécharger fichier Export pif",
        data=buffer,
        file_name="export_pif.xlsx",
        mime="application/vnd.ms-excel"
        )