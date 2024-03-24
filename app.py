import streamlit as st
import pandas as pd
from io import BytesIO

# Page configuration
st.set_page_config(page_title="Excel Processor", page_icon="📊", layout="wide")

# Sidebar for general information or settings
with st.sidebar:
    st.title("Excel Processor")
    st.write("This tool allows you to upload multiple Excel files, processes each file according to specified rules, and download each processed file separately. It's designed to ignore certain sheets and collect data from specified columns only.")

    st.write("## Instructions")
    st.write("""
    - Click the 'Upload Excel Files' button below to select your Excel files (.xlsx or .xlsm).
    - You can select multiple files at once for processing.
    - Once files are uploaded, each file's processed data will be displayed on the page with a corresponding download button.
    - Click the 'Download' button next to each file to save the processed data to your device.
    """)

    st.write("## Designed by: Shreif Shouman, Bsc.")

# Main page layout
st.title("Upload & Process Excel Files")

# File uploader allows user to add multiple files
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True, type=['xlsm', 'xlsx'], help="Select one or more Excel files to be processed.")

sheets_to_ignore = ['Übersicht', 'Vorlage_Seefracht', 'Vorlage_Luftfracht', 'Vorlage_Strasse', 'Legende', 'Frächter', 'Status']
columns_to_collect = ['Status', 'Einteildatum', 'Ladedatum', 'Kundenname', 'PO-Nummer', 'Auftrag']

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Using BytesIO to read the uploaded file
        xls = pd.ExcelFile(uploaded_file)
        adjusted_collected_data = pd.DataFrame()

        for sheet_name in xls.sheet_names:
            if sheet_name not in sheets_to_ignore:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=1)
                if set(columns_to_collect).issubset(df.columns):
                    adjusted_collected_data = pd.concat([adjusted_collected_data, df[columns_to_collect]], ignore_index=True)

        # Remove rows where all specified columns are NaN
        adjusted_collected_data.dropna(how='all', inplace=True)

        # Display and allow download of processed data for each file
        if not adjusted_collected_data.empty:
            st.subheader(f"Processed Data for {uploaded_file.name}")
            st.dataframe(adjusted_collected_data)

            # Convert DataFrame to Excel for downloading, applying date format
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                adjusted_collected_data.to_excel(writer, index=False, sheet_name='Processed Data')
                
                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Processed Data']
                
                # Define your date format
                date_format = workbook.add_format({'num_format': 'dd-mm-yyyy'})
                
                # Apply the date format to the date columns
                # Assuming your date columns are 'Einteildatum' and 'Ladedatum', adjust as necessary
                date_columns = ['Einteildatum', 'Ladedatum']  # Add or adjust your date column names
                for column in date_columns:
                    col_idx = adjusted_collected_data.columns.get_loc(column)  # Get the Excel column index
                    worksheet.set_column(col_idx + 1, col_idx + 1, None, date_format)  # Apply format, +1 due to Excel indexing

                writer.save()
            processed_data = output.getvalue()

            # Download button for each processed file
            st.download_button(label=f'Download Processed Data for {uploaded_file.name}',
                               data=processed_data,
                               file_name=f'processed_{uploaded_file.name}',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.write(f"No data collected or processed for {uploaded_file.name}.")
else:
    st.write("Upload Excel files to begin processing.")
