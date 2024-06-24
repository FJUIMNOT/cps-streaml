#homepage

import streamlit as st

st.set_page_config(page_title= "Streamlit ANA", page_icon=None, layout="centered", initial_sidebar_state="auto", menu_items=None)
st.sidebar.page_link("pages/CPS.py")
st.sidebar.page_link("pages/Sedigraph.py")
#st.page_link("your_app.py", label="Home", icon="🏠")
