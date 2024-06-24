import streamlit as st
import pandas as pd
import numpy as np
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import xlsxwriter
from io import BytesIO
from scipy.interpolate import interp1d

from operator import itemgetter
from itertools import groupby
import linecache

@st._cache_data

def save_data_excel(data): 
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    worksheet_1 = workbook.add_worksheet('Results')
    worksheet_1.write(0,0, 'YO')

    mydict = {}
    for sample in data : 
        name = sample['name']
        data_sample = sample['data']
        mydict[f"worksheet_{name}"] =  workbook.add_worksheet(name)
        mydict[f"worksheet_{name}"].write(0,0,data_sample)

    workbook.close()


    return output



st.set_page_config(layout="wide", page_title="Du coté de l'analyse")
st.sidebar.title('Sedigraph')
st.sidebar.write("Preparation des **données Sedigraph** et calcul des parametres. Le programme accepte en entrée les fichiers xlsx.")
uploaded_files = st.sidebar.file_uploader("Choose Sedigraph files",  accept_multiple_files=True)

index =[]
Data_all = list()
for file in uploaded_files : 
    sample_name = file.name
    index.append(sample_name)

    data = pd.read_excel(file,sheet_name=0, skiprows=35)
    

    Data_all.append({'name' : sample_name, 'data' : data})

st.download_button( label="Download Excel workbook", data=save_data_excel(Data_all).getvalue(), file_name="workbook.xlsx", mime="application/vnd.ms-excel")

