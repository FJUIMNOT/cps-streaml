import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO

@st.cache_data

# Function to read the first sheet of an XLS or XLSX file and convert it to a DataFrame
def read_first_sheet(file):
    return pd.read_excel(file, sheet_name=0)

# Function to write multiple DataFrames to an XLSX file with each DataFrame in a separate sheet
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

    worksheet1.autofit()
    chart_all = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

    row1 = 2
    col1 = 2
    rowi = 3
    coli = 3
    parameters = ['Echantillons', ' %<75µm ', ' %<50µm ', ' %<25µm ', ' %<10µm ', ' %<5µm ', ' %<2µm ', ' %<1,5µm ', ' %<1µm ', ' %<0,5µm ', ' %<0,3µm ']
    parameters_2nd = ['Echantillons', ' %<1µm ', ' %<0,3µm ']

    worksheet1.write_row(row1, col1, parameters, tab_format)
    worksheet1.write_row(rowi + 11, coli+11, parameters_2nd, tab_format)

    for df, filename in data_frames:
        sheet_name = df.iloc[4, 1][:31]
        worksheet = workbook.add_worksheet(sheet_name)
        chart_sheet = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

        # Write Result Data in 1st Tab
        worksheet1.write(rowi, col1, sheet_name, tab_format)
        worksheet1.write_row(rowi, coli, df.iloc[133:143, 1], tab_format)
        rowi += 1
        
        # Write Data in 2nd Tab
        worksheet1.write(rowi + 11, coli + 11, sheet_name, tab_format)
        worksheet1.write(rowi + 11, coli + 12, df.iloc[140, 1], tab_format) # Value in B142
        worksheet1.write(rowi + 11, coli + 13, df.iloc[142, 1], tab_format) # Value in B144

        # Write raw data
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header)

        for row_num, row_data in enumerate(df.values, 1):
            for col_num, cell_data in enumerate(row_data):
                if pd.isna(cell_data):
                    worksheet.write(row_num, col_num, '')  # Keep blank cells as blank
                else:
                    worksheet.write(row_num, col_num, cell_data)

        worksheet1.autofit()

        # Add series to charts
        Abs_range = [sheet_name, 151, 0, 249, 0]  # The range of x values (A152:A250)
        Ord_range = [sheet_name, 151, 1, 249, 1]  # The range of y values (B152:B250)
        chart_all.add_series({
            'name': sheet_name,
            'categories': Abs_range,
            'values': Ord_range,
            'line': {'width': 1.75}
        })

        chart_sheet.add_series({
            'name': sheet_name,
            'categories': Abs_range,
            'values': Ord_range,
            'line': {'width': 1.75}
        })
        chart_sheet.set_title ({
            'name': sheet_name, 
            'name_font':  {'name': 'calibri (corps)', 'size': 13}
        })
        chart_sheet.set_x_axis({
            'name': 'Diameter  (µm)',
            'name_font': {'name': 'Calibri', 'size': 10, 'bold': False},
            'num_font': {'name': 'Calibri', 'size': 9},
            'log_base': 10,
            'min': 0.1,
            'major_gridlines': {'visible': True, 'line': {'width': 0.6}},
            'minor_gridlines': {'visible': True, 'line': {'width': 0.03}},
            'crossing' : 0.1,
            'reverse' : True
        })
        chart_sheet.set_y_axis({
            'min': 0, 
            'max': 100,
            'minor_unit' : 10, 
            'major_unit': 20, 
            'num_font': {'name': 'Calibri', 'size': 9}
        })
        chart_sheet.set_legend({
            'position': 'bottom', 
            'font': {'size': 9}
        })
        chart_sheet.set_size({'width': 600, 'height': 376})

        worksheet.insert_chart(10, 10, chart_sheet)


    chart_all.set_x_axis({
        'name': 'Diameter  (µm)',
        'name_font': {'name': 'Calibri', 'size': 10, 'bold': False},
        'num_font': {'name': 'Calibri', 'size': 9},
        'log_base': 10,
        'min': 0.1,
        'major_gridlines': {'visible': True, 'line': {'width': 0.6}},
        'minor_gridlines': {'visible': True, 'line': {'width': 0.03}},
        'crossing' : 0.1,
        'reverse' : True
    })
    chart_all.set_y_axis({
        'min': 0, 
        'max': 100,
        'minor_unit' : 10, 
        'major_unit': 20, 
        'num_font': {'name': 'Calibri', 'size': 9}
    })
    chart_all.set_legend({
        'position': 'bottom', 
        'font': {'size': 9}
    })
    chart_all.set_size({'width': 600, 'height': 376})

    worksheet1.insert_chart(rowi + 2, coli, chart_all)
    workbook.close()
    output.seek(0)
    return output

#######################
# Streamlit app layout #
#######################

st.title("Sedigraph")

uploaded_files = st.file_uploader("Upload files : ", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    data_frames = []

    for uploaded_file in uploaded_files:
        # Read the first sheet of the uploaded file
        try:
            df = read_first_sheet(uploaded_file)
            data_frames.append((df, uploaded_file.name))
        except Exception as e:
            st.error(f"Error reading file {uploaded_file.name}: {e}")

    if data_frames:
        # Write DataFrames to a new XLSX file
        output = write_multiple_dfs_to_xlsx(data_frames)

        # Provide a download link for the new XLSX file
        st.download_button(label="Download XLSX file",
                           data=output,
                           file_name="Sedigraph.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
