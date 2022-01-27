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


# Define function that loads sheet preview
def load_sheet_preview(df):
    st.dataframe(df[st.session_state.selectbox_sheet])

    convert_df(df[st.session_state.selectbox_sheet])

    st.download_button(
        label='Download file', data=buffer,
        file_name='test.xlsx', key='download_button_file',
        mime='application/vnd.ms-excel'
    )
    return


# Define function that converts dataframe to CSV
@st.cache
def convert_df(df):     # Cache the conversion to prevent computation on every rerun        # noqa: E501
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

        writer.save()       # Output the Excel file to the buffer
    return


# Add selectbox
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=None)      # sheet_name=None needs explicitly including - it isn't the default        # noqa: E501

    st.selectbox(
        key='selectbox_sheet',
        label='Select worksheet',
        options=df.keys(),
        on_change=load_sheet_preview,
        args=(df, )
    )
