import streamlit as st
import pandas as pd
import pyodbc
from datetime import datetime


st.set_page_config(
    page_title="DC Land Status Mapping.py",
    page_icon="🗺️",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'file' not in st.session_state:
    st.session_state['file']=None

if "sheet" not in st.session_state:
    st.session_state['sheet']=None
    
uploaded_files = st.file_uploader('Upload Input Data', type=['xlsx','xls','xlsb','csv'])
st.session_state['file']=uploaded_files


if st.session_state['file']!=None:
    progress_text ="Loading files..............."
    my_bar = st.progress(0, text=progress_text)
    
    if uploaded_files.type=="text/csv":
        data=pd.read_csv(uploaded_files)
    else:
        datamain=pd.ExcelFile(uploaded_files)

        sheetname = st.selectbox(
        'Select sheet to proceed further !',
        (datamain.sheet_names))
        if st.button("Load Sheet"):
            st.session_state['sheet']=sheetname
        
    if st.session_state['sheet'] or st.session_state['file'].type=="text/csv":
        progress_text= "!!!!! Reading Sheet from File !!!!!"
        my_bar.progress(15, text=progress_text)

        if uploaded_files.type=="text/csv":
            "CSV File submitted"
        else:
            data=datamain.parse(sheetname)

        CNICcol = st.selectbox('Select CNIC Column for Mapping !',(data.columns))

        if st.button("Execute"):

            progress_text= "!!!!! Establishing Connection !!!!!"
            my_bar.progress(30, text=progress_text)
           
            server = 'DESKTOP-KL8DNKI\HASSANEYSQL'
            database = 'PDMA + UU Consolidated + 12.06.2023 + 2.06M'

            # Establish the connection
            conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database)
            # Create a cursor object
            cursor = conn.cursor()

            # Execute a SELECT query
            cursor.execute('select id,[cnic],[Status of Land] from [FinalTable-WithDehs]')

            progress_text= "!!!!! Fetching Data From SQL !!!!!"
            my_bar.progress(60, text=progress_text)

            rows = cursor.fetchall()

            # Get the column names
            columns = [column[0] for column in cursor.description]

            # # Close the cursor and connection
            cursor.close()
            conn.close()

            # Create a DataFrame from the rows and columns
            df = pd.DataFrame.from_records(rows, columns=columns)
            df['cnic']=df.cnic.astype('int64')

            progress_text= "!!!!! Closing Connection !!!!!"
            my_bar.progress(75, text=progress_text)

            mappedDC=pd.merge(data,df,right_on=["cnic"],left_on=[CNICcol],how="inner")

            progress_text= "!!!!! Mapped DC land Titles !!!!!"
            my_bar.progress(90, text=progress_text)

            filename="Output files/Mapped DC Titles "+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+uploaded_files.name

            progress_text= "!!!!! Saving File !!!!!"
            my_bar.progress(95, text=progress_text)

            if uploaded_files.type=="text/csv":
                mappedDC.to_csv(filename,index=False)
            else:
                with pd.ExcelWriter(filename) as writer:
                    mappedDC.to_excel(writer, sheet_name='Mapped',index=False)
        
            st.success("Processing Completed Click below to download output")
            with open(filename, 'rb') as my_file:
                st.download_button(label = 'Download output data', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

