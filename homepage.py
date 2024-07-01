import streamlit as st

# Set page configuration
st.set_page_config(
    page_title="Streamlit ANA",
    page_icon=None,
    layout="centered",
    initial_sidebar_state="auto"
)

# Create the homepage content
st.title("Welcome to Streamlit ANA")
st.write("""
    This is the homepage of the Streamlit ANA application. 
    Use the sidebar to navigate to different pages.
""")
