import streamlit as st
import pandas as pd
import pyodbc
import time
from datetime import datetime
from fuzzywuzzy import fuzz
import numpy as np


st.set_page_config(
    page_title="Plinth Level Disbursement.py",
    page_icon="🔗",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'batchfile' not in st.session_state:
    st.session_state['batchfile']=None

if 'sheetbatch' not in st.session_state:
    st.session_state['sheetbatch']=None

if 'Tcol' not in st.session_state:
    st.session_state['Tcol']=None
if 'Bcol' not in st.session_state:
    st.session_state['Bcol']=None

batchfile = st.file_uploader('Upload Batch Data', type=['xlsx','xls','xlsb'])
st.session_state['batchfile']=batchfile

if st.session_state['batchfile']!=None:
    batchdata=pd.ExcelFile(batchfile)

    sheetbatch = st.selectbox(
    'Select Batch sheet to proceed further !',
    (batchdata.sheet_names))

    if st.button("Load Batch Sheet"):
        st.session_state['sheetbatch']=sheetbatch
        
    if st.session_state['sheetbatch']:
        batchdata=batchdata.parse(sheetbatch)
        st.dataframe(batchdata)

        c1,c2=st.columns(2)
        tehsil = c1.selectbox(
        'Select Tehsil column to proceed further !',
        (batchdata.columns))
        st.session_state['Tcol']=tehsil

        bank = c2.selectbox(
        'Select Bank column to proceed further !',
        (batchdata.columns))
        st.session_state['Bcol']=bank


if st.session_state['sheetbatch']!=None and st.button("Execute"):
    bat=batchdata.copy()

    progress_text ="Select Columns To make Key"
    my_bar = st.progress(0, text=progress_text)

    bat["KEYstop"]=bat[tehsil]+"_"+bat[bank].astype(str)

    progress_text= "!!!!! Aur ye mai Asman ki Uchaiyon mai !!!!!"
    my_bar.progress(10, text=progress_text)
    time.sleep(1)

    limited_df = pd.DataFrame()
    remaining_df = pd.DataFrame()

    for key in bat['KEYstop'].unique():
        key_data = bat[bat['KEYstop'] == key]

        if len(key_data) > 700:
            limited_data = key_data.sample(n=700)  # Randomly select 700 rows
            remaining_data = key_data.drop(limited_data.index)  # Remove selected rows
            remaining_df = pd.concat([remaining_df, remaining_data])  # Append remaining data
        
        else:
            limited_data = key_data
        
        limited_df = pd.concat([limited_df, limited_data])
        
   
    progress_text= "!!!!! Aur Mai Motherboard hu jo idhr aya !!!!"
    my_bar.progress(75, text=progress_text)
    time.sleep(1)


    filename='Output files/Strategy File '+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+st.session_state['batchfile'].name
    with pd.ExcelWriter(filename) as writer:
        limited_df.to_excel(writer, sheet_name="700 Per Key Cleared", index=False)
        remaining_df.to_excel(writer, sheet_name="Remaining to be send later", index=False)



    progress_text="!!!!! Chalo kaam hogaya check krlo !!!!"
    my_bar.progress(100, text=progress_text)
    time.sleep(1)

    st.success("Processing Completed Click below to download output")
    with open(filename, 'rb') as my_file:
        st.download_button(label = 'Download output data', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')