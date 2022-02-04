#!/usr/bin/env python
# -*- coding: utf-8 -*-

import io

import pandas as pd
import streamlit as st

# SETUP BUFFER THAT WILL HOLD EXCEL FILE FOR DOWNLOAD
buffer = io.BytesIO()

# CREATE APP
# Add file_uploader
uploaded_file = st.file_uploader(
    label='Upload a file', type=['xls', 'xlsx', 'xlsm'],
    key='file-uploader',
    help='''
        Upload an Excel file. The file must be
        closed in order for you to upload it.
    '''
)


# Define function that converts df to CSV
@st.cache
def convert_df(df):     # Cache the conversion to prevent computation on every rerun        # noqa: E501
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

        writer.save()       # Output the Excel file to the buffer
    return


# Add selectbox
if 'selectbox_sheet' not in st.session_state:       # Initialise variable
    st.session_state['selectbox_sheet'] = '--'

if uploaded_file is not None:
    sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)      # This creates a dictionary of dataframes. sheet_name=None needs explicitly including - it isn't the default        # noqa: E501

    default_option = {'--': ''}

    selectbox_options = dict(**default_option, **sheets_dict)     # Join dictionaries        # noqa: E501

    st.selectbox(
        key='selectbox_sheet',
        label='Select worksheet',
        options=selectbox_options.keys()
    )

# Load sheet preview
if st.session_state.selectbox_sheet != '--':
    st.dataframe(sheets_dict[st.session_state.selectbox_sheet])

    convert_df(sheets_dict[st.session_state.selectbox_sheet])      # This accepts a dictionary of dataframes as input        # noqa: E501

    st.download_button(
        label='Download file', data=buffer,
        file_name='test.xlsx', key='download_button_file',
        mime='application/vnd.ms-excel'
    )
