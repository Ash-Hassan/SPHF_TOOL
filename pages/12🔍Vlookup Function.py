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

if 'File1' not in st.session_state:
    st.session_state['File1']=None
if 'File2' not in st.session_state:
    st.session_state['File2']=None

if 'Sfile1' not in st.session_state:
    st.session_state['Sfile1']=None
if 'Sfile2' not in st.session_state:
    st.session_state['Sfile2']=None

if 'f1col' not in st.session_state:
    st.session_state['f1col']=None
if 'f2col' not in st.session_state:
    st.session_state['f2col']=None

if 'how' not in st.session_state:
    st.session_state['how']=None

c1,c2=st.columns(2)

file1 = c1.file_uploader('Upload First File', type=['xlsx','xls','xlsb'])
st.session_state['File1']=file1

file2 = c2.file_uploader('Upload Second File', type=['xlsx','xls','xlsb'])
st.session_state['File2']=file2

if st.session_state['File1']!=None:
    fileOne=pd.ExcelFile(file1)

    sheetfileOne = c1.selectbox(
    'Select Batch sheet to proceed further !',
    (fileOne.sheet_names))

    if c1.button("Load 1stFile Sheet"):
        st.session_state['Sfile1']=sheetfileOne
        
    if st.session_state['Sfile1']:
        DataOne=fileOne.parse(sheetfileOne)
        c1.dataframe(DataOne)

        f1col = c1.selectbox(
        'Select 1st File column for mapping !!',
        (list(DataOne.columns)))
        st.session_state['f1col']=f1col

if st.session_state['File2']!=None:
    fileTwo=pd.ExcelFile(file2)

    sheetfileTwo = c2.selectbox(
    'Select Batch sheet to proceed further !!',
    (fileTwo.sheet_names))

    if c2.button("Load 2ndFile Sheet"):
        st.session_state['Sfile2']=sheetfileTwo
        
    if st.session_state['Sfile2']:
        DataTwo=fileTwo.parse(sheetfileTwo)
        c2.dataframe(DataTwo)

        f2col = c2.selectbox(
        'Select 2nd File column for mapping !!',
        (list(DataTwo.columns)))
        st.session_state['f2col']=f2col

if st.session_state['f1col']!=None and st.session_state['f2col']!=None:
    data_one = pd.DataFrame(
        {
            "File 1 Columns": ["Select All"]+list(DataOne.columns),
            "File 1 Checkbox": [False]+[ False for f in list(DataOne.columns)],
        }
    )
    data_two = pd.DataFrame(
        {
            "File 2 Columns": ["Select All"]+list(DataTwo.columns),
            "File 2 Checkbox": [False]+[ False for f in list(DataTwo.columns)],
        }
    )

    d1=c1.data_editor(
        data_one,
        column_config={
            "File 1 Checkbox": st.column_config.CheckboxColumn(
                "Select Columns to include in New file",
                help="Select **Columns** to include from 1st file",
                default=False,
            )
        },
        disabled=["File 1 Columns"],
        hide_index=True,
    )
    d2=c2.data_editor(
        data_two,
        column_config={
            "File 2 Checkbox": st.column_config.CheckboxColumn(
                "Select Columns to include in New file",
                help="Select **Columns** to include from 1st file",
                default=False,
            )
        },
        disabled=["File 2 Columns"],
        hide_index=True,
    )
    how = st.selectbox(
        'Select how to map?',
        (["left","right","inner","outer"]))
    st.session_state['how']=how



if st.session_state['f1col']!=None and st.session_state['f2col']!=None and st.button("Execute"):
    f1dict=dict(zip(d1["File 1 Columns"], d1["File 1 Checkbox"]))
    f2dict=dict(zip(d2["File 2 Columns"], d2["File 2 Checkbox"]))

    if f1dict["Select All"] and f2dict["Select All"]:
        fnf=pd.merge(DataOne,DataTwo,left_on=st.session_state['f1col'],right_on=st.session_state['f2col'],how=st.session_state['how'])
    elif f1dict["Select All"]:
        fnf=pd.merge(DataOne,DataTwo[[col for col in f2dict.keys() if f2dict[col]]],
                     left_on=st.session_state['f1col'],right_on=st.session_state['f2col'],how=st.session_state['how'])
    elif f2dict["Select All"]:
        fnf=pd.merge(DataOne[[col for col in f1dict.keys() if f1dict[col]]],DataTwo,
                     left_on=st.session_state['f1col'],right_on=st.session_state['f2col'],how=st.session_state['how'])
    else:
        fnf=pd.merge(DataOne[[col for col in f1dict.keys() if f1dict[col]]],DataTwo[[col for col in f2dict.keys() if f2dict[col]]]
                     ,left_on=st.session_state['f1col'],right_on=st.session_state['f2col'],how=st.session_state['how'])


    st.dataframe(fnf)

    # bat=batchdata.copy()

    # progress_text ="Select Columns To make Key"
    # my_bar = st.progress(0, text=progress_text)

    # bat["KEYstop"]=bat[tehsil]+"_"+bat[bank].astype(str)

    # progress_text= "!!!!! Aur ye mai Asman ki Uchaiyon mai !!!!!"
    # my_bar.progress(10, text=progress_text)
    # time.sleep(1)

    # limited_df = pd.DataFrame()
    # remaining_df = pd.DataFrame()

    # for key in bat['KEYstop'].unique():
    #     key_data = bat[bat['KEYstop'] == key]

    #     if len(key_data) > 700:
    #         limited_data = key_data.sample(n=700)  # Randomly select 700 rows
    #         remaining_data = key_data.drop(limited_data.index)  # Remove selected rows
    #         remaining_df = pd.concat([remaining_df, remaining_data])  # Append remaining data
        
    #     else:
    #         limited_data = key_data
        
    #     limited_df = pd.concat([limited_df, limited_data])
        
   
    # progress_text= "!!!!! OO bhaiiiii land karadyyy plz Land karady !!!!"
    # my_bar.progress(75, text=progress_text)
    # time.sleep(1)


    # filename='Strategy File '+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+st.session_state['batchfile'].name
    # with pd.ExcelWriter(filename) as writer:
    #     limited_df.to_excel(writer, sheet_name="700 Per Key Cleared", index=False)
    #     remaining_df.to_excel(writer, sheet_name="Remaining to be send later", index=False)



    # progress_text="!!!!! Mai Motherboard hu jo idhr aya !!!!"
    # my_bar.progress(100, text=progress_text)
    # time.sleep(1)

    # st.success("Processing Completed Click below to download output")
    # with open(filename, 'rb') as my_file:
    #     st.download_button(label = 'Download output data', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')