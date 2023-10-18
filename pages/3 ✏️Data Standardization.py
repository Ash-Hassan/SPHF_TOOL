import streamlit as st
import pandas as pd
import time
import json
from datetime import datetime

st.set_page_config(
    page_title="Data Standardization",
    page_icon="✏️",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'file' not in st.session_state:
    st.session_state['file']=None

tab11,tab22=st.tabs(["For Excel DOB - DOI","For CSV DOB - DOI"])


with tab11:
    excelfile = st.file_uploader("Upload an Excel File", type=["xlsx", "xls"])
    st.session_state['file']=excelfile
    if st.session_state['file']!=None:
        data = pd.ExcelFile(st.session_state['file'])
        c1,c2=st.columns(2)

        sheetname = c1.selectbox(
        'Select sheet to proceed further !',
        (data.sheet_names))

        colname = c2.selectbox(
        'Select Column Name to proceed further !',
        (data.parse(sheetname).columns))

        if st.session_state['file']!=None:
            if st.button("Execute Process"):
                with st.spinner('Wait for it...'):
                    progress_text ="!!!!! Horaha hai kaam sbr rakho !!!!!"
                    my_bar = st.progress(0, text=progress_text)
                    time.sleep(2)

                    df=data.parse(sheetname)
                    st.dataframe(df,use_container_width=True)
                    
                    with open('ToolData/stdDict.json', 'r') as fp:
                        mainDICT = json.load(fp)
                        
                    df[colname]=df[colname].str.upper()
                    df["Mapped Banks"]=df[colname].map(mainDICT)

                    my_bar.progress(50, text="!!! Aur jaldi nhe hosktaa, sabrooka  !!!")
                    time.sleep(2)
                    
                    outputFileName='Output files/Mapped File '+st.session_state['file'].name+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S "))+'.xlsx'
                    with pd.ExcelWriter(outputFileName, engine='xlsxwriter') as writer:
                        df.to_excel(writer,sheet_name=sheetname,index=False)

                    my_bar.progress(100, text="!!! hogaya kaam Enjoy kro  !!!")
                    time.sleep(2)

                    st.success("Processing Completed Click below to download output")

                    with open(outputFileName, 'rb') as my_file:
                        st.download_button(label = 'Download output data', data = my_file, file_name = outputFileName, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


with tab22:
    excelfile = st.file_uploader("Upload an Excel File", type=["xlsx", "xls","csv"])
    st.session_state['file']=excelfile
    if st.session_state['file']!=None:
        data = pd.read_csv(st.session_state['file'],low_memory=False)

        colname = st.selectbox(
        'Select Column Name to proceed further !',
        (data.columns))

        if st.session_state['file']!=None:
            if st.button("Execute Process"):
                with st.spinner('Wait for it...'):

                    progress_text ="!!!!! Horaha hai kaam sbr rakho !!!!!"
                    my_bar = st.progress(0, text=progress_text)
                    time.sleep(2)

                    df=data
                    st.dataframe(df,use_container_width=True)
                    
                    with open('ToolData/stdDict.json', 'r') as fp:
                        mainDICT = json.load(fp)
                        
                    df[colname]=df[colname].str.upper()
                    df["Mapped Banks"]=df[colname].map(mainDICT)

                    my_bar.progress(50, text="!!! Aur jaldi nhe hosktaa, sabrooka  !!!")
                    time.sleep(2)
                    
                    outputFileName='Output files/Mapped File '+st.session_state['file'].name+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S "))+'.csv'
                    df.to_csv(outputFileName,index=False)

                    my_bar.progress(100, text="!!! hogaya kaam Enjoy kro  !!!")
                    time.sleep(2)

                    st.success("Processing Completed Click below to download output")

                    with open(outputFileName, 'rb') as my_file:
                        st.download_button(label = 'Download output data', data = my_file, file_name = outputFileName, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

