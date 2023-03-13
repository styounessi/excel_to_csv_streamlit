import streamlit as st
import pandas as pd

import zipfile
import uuid
import io
import os

from PIL import Image

#----------------------------------------------------------#

st.markdown('''
            <style> .font {
            font-size: 45px;
            font: "monospace";
            color: #008000;
            text-align: left;}
            </style>
            ''', unsafe_allow_html=True)

st.sidebar.markdown(
    '<p class="font">Excel to CSV File Converter</p>', unsafe_allow_html=True)

logo_image = Image.open('imgs/excel_img.jpg')
st.sidebar.image(logo_image, use_column_width=True)

with st.sidebar.expander('**How it Works**', expanded=True):
    st.write('''
            This converter takes Excel files and converts them directly to CSV files
            without changing the columns or rows in any way. There is a limit to the size 
            and quantity of files that can be processed so use cases will vary but simple
            operations and files will benefit from the elimination of manual conversion.
            
            Please ensure that input files are single sheet Excel files, not multi-sheet, since
            CSVs do not support multiple sheets.

            Please ensure that only .xlsx or .xls files are used, using any other file type
            will result in an error.
            ''')

st.markdown('<p class="font">Upload Excel Files</p>', unsafe_allow_html=True)

# The `label` argument cannot be an empty value. This is discouraged for accessibility reasons
# and may be disallowed in the future by raising an exception. A string is provided but it is hidden
# using `label_visibility` instead to prevent any future issues.
file_upload = st.file_uploader('Upload File(s)', type=['xlsx', 'xls'], accept_multiple_files=True,
                               label_visibility='hidden')

if file_upload is not None:
    file_list = []  # create a list to store file data and name
    for file in file_upload:
        # Check if the uploaded file is an accepted excel file
        if file.type not in ['application/vnd.ms-excel',
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
            st.error(
                f'This file type is invalid: {file.name} must be an .xlsx or .xls file to proceed.')
        else:
            # Split the file name from the file extension to make the download button distinctive
            df_excel = pd.read_excel(file)
            csv = df_excel.to_csv(index=False)
            df_csv = pd.read_csv(io.StringIO(csv))
            file_name_no_ext = os.path.splitext(file.name)[0]

            # Check the shape of the original file and csv file to ensure it has the same
            # number of rows and columns
            if df_excel.shape != df_csv.shape:
                st.error(f'''
                         The CSV file generated from {file.name} is not identical to the original file. 
                         Inspect the Excel file for merged/formatted cells, formulas, special characters, etc
                         and try again.
                         ''')
            else:
                file_list.append((csv, file_name_no_ext + '.csv'))
                # If there's only one file, show the regular download button
                st.download_button(
                    label=f'Download \'{file_name_no_ext}\' as CSV',
                    data=csv,
                    file_name=file_name_no_ext + '.csv',
                    mime='text/csv',
                    key=str(uuid.uuid4())) # Generate random key to avoid 'DuplicateWidgetID' error

    # When more than one file is uploaded, add an option to download them all as zip
    if len(file_list) > 1:
        zipped_csvs = io.BytesIO()
        with zipfile.ZipFile(zipped_csvs, mode='w') as zip_file:
            for file_data, file_name in file_list:
                zip_file.writestr(file_name, file_data)
        st.download_button(
            label='Download All Files as Zip',
            data=zipped_csvs.getvalue(),
            file_name='Combined_CSVs.zip',
            mime='application/zip',
            key=str(uuid.uuid4()))
