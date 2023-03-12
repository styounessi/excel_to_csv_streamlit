import streamlit as st
import pandas as pd

import io
import os

from PIL import Image

#-----------------------------------------#

st.markdown('''
            <style> .font {
            font-size: 45px;
            font: "monospace";
            color: #008000;
            text-align: left;}
            </style>
            ''', unsafe_allow_html=True)

st.sidebar.markdown('<p class="font">Excel to CSV File Converter</p>', unsafe_allow_html=True)

logo_image = Image.open('imgs/excel_img.jpg')
st.sidebar.image(logo_image, use_column_width=True)

with st.sidebar.expander('**How it Works**', expanded=True):
    st.write('''
            This converter takes Excel files and converts them directly to CSV files
            without changing the columns or rows in any way. There is a limit to the size 
            and quantity of files that can be processed so use cases will vary but simple
            operations and files will benefit from the elimination of manual conversion.
            
            Please ensure that input files are single sheet Excel files, not multi-sheet, since
            CSVs do not have multiple sheets.

            Please ensure that only .xlsx or .xls files are used, using any other file type
            will result in an error.
            ''')

st.markdown('<p class="font">Upload Excel Files</p>', unsafe_allow_html=True)

file_upload = st.file_uploader('', type=['xlsx', 'xls'], accept_multiple_files=True)

if file_upload is not None:
    for file in file_upload:
        # Check if the uploaded file is an accepted excel file
        if file.type not in ['application/vnd.ms-excel',
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
            st.error(f'This file type is invalid: {file.name} must be an .xlsx or .xls file to proceed.')
        else:
            # Split the file name from the file extension to make the download button distinctive
            file_name = file.name
            file_name_no_ext = os.path.splitext(file_name)[0]
            df_excel = pd.read_excel(file)
            csv = df_excel.to_csv(index=False)
            df_csv = pd.read_csv(io.StringIO(csv))
            
            # Checks the shape of the original file and csv file to ensure it has the same
            # number of rows and columns
            if df_excel.shape != df_csv.shape:
                st.error(f'''
                         The CSV file generated from {file_name} is not identical to the original file. 
                         Inspect the Excel file for merged/formatted cells, formulas, special characters, etc
                         and try again.
                         ''')
            else:
                st.download_button(
                    label=f'Download \'{file_name_no_ext}\' as CSV',
                    data=csv,
                    file_name=file_name.replace('.xlsx', '.csv'),
                    mime='text/csv')
