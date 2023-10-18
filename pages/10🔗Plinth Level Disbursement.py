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


c1,c2,c3=st.columns(3)

if 'masterfile' not in st.session_state:
    st.session_state['masterfile']=None
if 'batchfile' not in st.session_state:
    st.session_state['batchfile']=None
if 'blockfile' not in st.session_state:
    st.session_state['blockfile']=None
if 'sheetmaster' not in st.session_state:
    st.session_state['sheetmaster']=None
if 'sheetbatch' not in st.session_state:
    st.session_state['sheetbatch']=None
if 'sheetblock' not in st.session_state:
    st.session_state['sheetblock']=None

masterfile = c1.file_uploader('Upload Master Data', type=['xlsx','xls','xlsb','csv'])
st.session_state['masterfile']=masterfile

batchfile = c2.file_uploader('Upload Batch Data', type=['xlsx','xls','xlsb'])
st.session_state['batchfile']=batchfile

blockfile = c3.file_uploader('Upload Block Data', type=['xlsx','xls','xlsb','csv'])
st.session_state['blockfile']=blockfile

if st.session_state['masterfile']!=None:
    if masterfile.type!="text/csv":
        masterdata=pd.ExcelFile(masterfile)

        sheetmaster = c1.selectbox(
        'Select Master sheet to proceed further !',
        (masterdata.sheet_names))
        if c1.button("Load Master Sheet"):
            st.session_state['sheetmaster']=sheetmaster
            
        if st.session_state['sheetmaster']:
            masterdata=masterdata.parse(sheetmaster)

            c1.dataframe(masterdata)
    else:
        masterdata=pd.read_csv(masterfile,low_memory=False)
        st.session_state['sheetmaster']="Sheet1"
        c1.dataframe(masterdata)


if st.session_state['batchfile']!=None:
    batchdata=pd.ExcelFile(batchfile)

    sheetbatch = c2.selectbox(
    'Select Batch sheet to proceed further !',
    (batchdata.sheet_names))

    if c2.button("Load Batch Sheet"):
        st.session_state['sheetbatch']=sheetbatch
        
    if st.session_state['sheetbatch']:
        batchdata=batchdata.parse(sheetbatch)

        c2.dataframe(batchdata)

if st.session_state['blockfile']!=None:
    if masterfile.type!="text/csv":
        blockdata=pd.ExcelFile(blockfile)

        sheetblock = c3.selectbox(
        'Select Block sheet to proceed further !',
        (blockdata.sheet_names))

        if c3.button("Load Block Sheet"):
            st.session_state['sheetblock']=sheetblock
            
        if st.session_state['sheetblock']:
            blockdata=blockdata.parse(sheetblock)

            c3.dataframe(blockdata)
    else:
        blockdata=pd.read_csv(blockfile,low_memory=False)
        st.session_state['sheetblock']="Sheet1"
        c3.dataframe(blockdata)


if st.session_state['sheetmaster']!=None and st.session_state['sheetbatch']!=None and st.session_state['sheetblock']!=None:
    bat=batchdata.copy()
    mast=masterdata.copy()
    blck=blockdata.copy()

    progress_text ="Verifying data with Blocklist..............."
    my_bar = st.progress(0, text=progress_text)

    bat['CNIC_Duplicate'] = np.where(bat['CNIC'].duplicated(), 'Yes', 'No')     

    progress_text= "!!!!! Firingup the Engines !!!!!"
    my_bar.progress(10, text=progress_text)
    time.sleep(1)

    submast=mast[["CNIC",'Urban Unit #',"Batch#", 'Processed / Rejected','2d Inst batch #', 'Instalment 2 Processed','3rd Inst batch #','Instalment 3 Processed']]
    
    progress_text="!!!!! Proceding to Runway !!!!!"
    my_bar.progress(30, text=progress_text)
    time.sleep(1)
    
    mapped=pd.merge(bat,submast,left_on="CNIC",right_on="CNIC",how='left')

    progress_text="!!!!! Taking OFF !!!!!"
    my_bar.progress(50, text=progress_text)
    time.sleep(1)

    Fmap=pd.merge(mapped,blck[["CNIC","Block List","Reasons"]],left_on="CNIC",right_on="CNIC",how='left')
    
    progress_text= "!!!!!Preparing to Land on Jupiter!!!!"
    my_bar.progress(75, text=progress_text)
    time.sleep(1)


    filename='Output files/Mapped File '+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+st.session_state['batchfile'].name
    with pd.ExcelWriter(filename) as writer:
        Fmap.to_excel(writer, sheet_name="Mapped Data", index=False)


    progress_text="!!!!! Successfully Landed on Jupiter !!!!"
    my_bar.progress(100, text=progress_text)
    time.sleep(1)

    st.success("Processing Completed Click below to download output")
    with open(filename, 'rb') as my_file:
        st.download_button(label = 'Download output data', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')