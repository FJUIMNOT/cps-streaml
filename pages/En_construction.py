import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO

@st.cache_data

def write_multiple_dfs_to_xlsx(data_frames):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})
    worksheet1 = workbook.add_worksheet("Results")

    tab_format = workbook.add_format({
        'bold': False,
        'font_size': 10,
        'font_name': 'Arial',
        'num_format': '#,##0.0',
        'center_across': True,
        'border': 1
    })

    tab_title = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'center_across': True,
        'border': 1,
        'bg_color' : '#fcd5b4',
    })

    tab_title2 = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'center_across': True,
        'border': 1,
        'bg_color' : '#ccc0da',
    })

    tab_title3 = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'center_across': True,
        'border': 1,
        'bg_color' : '#daeef3',
    })

    tab_title4 = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'center_across': True,
        'border': 1,
        'bg_color' : '#d8e4bc',
    })

    Parameters = ["Vpt (mL/g)",	"S (m²/g)",	"","", "VD<1µm (mL/g)","", "Ldp","", "Méthode 1996","" , "Vg/Vs", "V2/V1 (%)","", "V3","", "   ","Vpt (mL/g)", "S (m²/g)", "Mode (nm)", "Ldp", "Ldp log"]
    Parameters_2 = ["    ","    ","1ère intr",	"2ème intr",	"1ère intr",	"2ème intr",	"1ère intr",	"2ème intr",	"If",	"Is", " ",  "130°",	"    140°   ",	"1ère intr",	"2ème intr","   "," "," "," "," "," "]              
    
    worksheet1.write_row(3,2, Parameters,tab_format)
    worksheet1.write_row(4,2,Parameters_2,tab_format)
    worksheet1.merge_range(3,0,4,0," Echantillons",tab_format)
    worksheet1.merge_range(3,1,4,1," Lot ",tab_format)
    worksheet1.merge_range(2,0,2,12,"Solvay (140° - 0,485 N/cm)", tab_title)
    worksheet1.merge_range(2,13,2,14,"V2/V1 (0,484 N/cm)", tab_title2)
    worksheet1.merge_range(2,15,2,16,"V3",tab_title3)
    worksheet1.merge_range(2,17,2,22,"Conditions GY (141,3° - 0,480 N/cm)", tab_title4)
    worksheet1.merge_range(3,4,3,5,"Mode (nm)",tab_format)
    worksheet1.autofit()
    
    #rowi = 5
   # for df, file_name in  data_frames :
    #    sheet_name  = df.iloc[3, 8][:31]
     #   worksheet1.write(rowi,0, file_name)
      #  worksheet = workbook.add_worksheet(sheet_name)
#
 #       col= 0
  #      for i in range(6) : 
   #         worksheet.write_column(5, col, df.iloc[22, 7 - col])
    #        col =+1
                    
    workbook.close()
    output.seek(0)
    return output


#######################
# Streamlit app layout
#######################


st.title("Mercury porosimetry")

uploaded_files = st.file_uploader("Upload files : ", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    data_frames = []

    for uploaded_file in uploaded_files:

        df = pd.read_excel(uploaded_file, engine = 'xlrd')
        data_frames.append((df, uploaded_file.name))

        if data_frames :
            # Write DataFrames to a new XLSX file
            output = write_multiple_dfs_to_xlsx(data_frames)

            # Provide a download link for the new XLSX file
            st.download_button(label="Download XLSX file",
                            data=output,
                            file_name="PoroHg.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

