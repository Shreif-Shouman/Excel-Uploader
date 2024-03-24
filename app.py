import streamlit as st
import pandas as pd
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')



# Page setup
st.set_page_config(page_title="Excel Processor", page_icon="📊", layout="wide")

# Sidebar content
with st.sidebar:
    st.title("Excel Processor")
    st.write("""
        This tool allows you to upload multiple Excel files, processes each file according to specified rules, and download each processed file separately. It's designed to ignore certain sheets and collect data from specified columns only.

        ## Instructions
        - Click the 'Upload Excel Files' button below to select your Excel files (.xlsx or .xlsm).
        - You can select multiple files at once for processing.
        - Once files are uploaded, each file's processed data will be displayed on the page with a corresponding download button.
        - Click the 'Download' button next to each file to save the processed data to your device.

        ## Designed by: Shreif Shouman, Bsc.
    """)

# Main content
st.title("Upload & Process Excel Files")
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True, type=['xlsm', 'xlsx'])

sheets_to_ignore = ['Übersicht', 'Vorlage_Seefracht', 'Vorlage_Luftfracht', 'Vorlage_Strasse', 'Legende', 'Frächter', 'Status']
columns_to_collect = ['Status', 'Einteildatum', 'Ladedatum', 'Kundenname', 'PO-Nummer', 'Auftrag']

if uploaded_files:
    for uploaded_file in uploaded_files:
        xls = pd.ExcelFile(uploaded_file)
        adjusted_collected_data = pd.DataFrame()

        for sheet_name in xls.sheet_names:
            if sheet_name not in sheets_to_ignore:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=1)
                if set(columns_to_collect).issubset(df.columns):
                    adjusted_collected_data = pd.concat([adjusted_collected_data, df[columns_to_collect]], ignore_index=True)

        adjusted_collected_data.dropna(how='all', inplace=True)

        for col in ['Einteildatum', 'Ladedatum']:
            if col in adjusted_collected_data.columns:
                adjusted_collected_data[col] = pd.to_datetime(adjusted_collected_data[col], errors='coerce').dt.date

        if not adjusted_collected_data.empty:
            st.subheader(f"Processed Data for {uploaded_file.name}")
            st.dataframe(adjusted_collected_data)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                adjusted_collected_data.to_excel(writer, index=False, sheet_name='Processed Data')

                workbook = writer.book
                worksheet = writer.sheets['Processed Data']

                date_format = workbook.add_format({'num_format': 'dd-mm-yyyy'})
                for col_name, col_idx in zip(adjusted_collected_data.columns, range(len(adjusted_collected_data.columns))):
                    if col_name in ['Einteildatum', 'Ladedatum']:
                        worksheet.set_column(col_idx, col_idx, 20, date_format)

                writer.save()
            processed_data = output.getvalue()

            st.download_button(label=f'Download Processed Data for {uploaded_file.name}',
                               data=processed_data,
                               file_name=f'processed_{uploaded_file.name}.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.write(f"No data collected or processed for {uploaded_file.name}.")
else:
    st.write("Upload Excel files to begin processing.")
