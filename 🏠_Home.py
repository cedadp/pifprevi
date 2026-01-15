import pandas as pd
import streamlit as st

st.set_page_config(page_title="PIF Prévis", page_icon="🛫", layout="centered", initial_sidebar_state="auto")


st.title('PIF Prévis') 


with st.sidebar.expander("Version"):
    st.sidebar.info("Version : 2.0")
    



hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
