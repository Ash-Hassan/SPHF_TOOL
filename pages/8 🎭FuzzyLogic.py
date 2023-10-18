import streamlit as st
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz

st.set_page_config(
    page_title="Files Consolidation",
    page_icon="📖",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'file' not in st.session_state:
    st.session_state['file']=None

if "sheet" not in st.session_state:
    st.session_state['sheet']=None
    
uploaded_files = st.file_uploader('Upload Input Data', type=['xlsx','xls','xlsb'])
st.session_state['file']=uploaded_files

print(uploaded_files)

tab1,tab2=st.tabs(["Simple Name Checking","Proper fuzzylogic multi columns"])

with tab1:
    if st.session_state['file']!=None:
        datamain=pd.ExcelFile(uploaded_files)

        sheetname = st.selectbox(
        'Select sheet name to proceed further !',
        (datamain.sheet_names))
        if st.button("Read Sheet"):
            st.session_state['sheet']=sheetname
            
        if st.session_state['sheet']:
            data=datamain.parse(sheetname)

            c1,c2=st.columns(2)

            col1st = c1.selectbox('Select First Column for Comparision !',(data.columns))
            col2nd = c2.selectbox('Select Second Column for Comparision !',(data.columns))

            if st.button("Verify"):
                data[col1st]=data[col1st].str.replace("ASAAN ACCOUNT","").str.replace("(","").str.replace(")","")
                data[col2nd]=data[col2nd].str.replace("ASAAN ACCOUNT","").str.replace("(","").str.replace(")","")

                data[col1st]=data[col1st].fillna("N/A")
                data[col1st]=data[col1st].str.upper()
                
                data[col2nd]=data[col2nd].fillna("N/A")
                data[col2nd]=data[col2nd].str.upper()

                def part_ratio(df,columns,f_col,thresh=70):
                    df[f_col]=df[f_col].str.upper()
                    df[f_col]=df[f_col].fillna("-")
                    col_name=[]
                    for i in columns:
                        df[i]=df[i].str.upper()
                        print(i)
                        col_name.append(i+' partial_ratio')
                        df[i+' partial_ratio'] = df.apply(lambda x: fuzz.partial_ratio(x[f_col], x[i]), axis=1)
                        
                    condition = df[col_name].gt(70).any(axis=1)
                    df[f_col+' Not Matching'] = condition.map({True: 'Matched', False: 'Not Matched'})
                    df[f_col+' Not Matching'] = np.where((df[f_col]=="-"),'Not Matched',df[f_col+' Not Matching'])
                    return df
                
                verify=part_ratio(data,[col1st],col2nd)
                
                st.write(verify[list(verify.columns)[-1]].value_counts())

                with pd.ExcelWriter("Output files/FuzzyLogic "+uploaded_files.name) as writer:
                    verify[verify[list(verify.columns)[-1]]=="Matched"].to_excel(writer, sheet_name='Names Matched',index=False)
                    verify[verify[list(verify.columns)[-1]]=="Not Matched"].to_excel(writer, sheet_name='Names Not Matched',index=False)
            
                st.success("Processing Completed Click below to download output")
                with open("Output files/FuzzyLogic "+uploaded_files.name, 'rb') as my_file:
                    st.download_button(label = 'Download output data', data = my_file, file_name = "FuzzyLogic "+uploaded_files.name, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



with tab2:
    if st.session_state['file']!=None:
        datamain=pd.ExcelFile(uploaded_files)

        sheetname = st.selectbox(
        'Select sheet to proceed further !',
        (datamain.sheet_names))
        if st.button("Load Sheet"):
            st.session_state['sheet']=sheetname
            
        if st.session_state['sheet']:
            data=datamain.parse(sheetname)

            c1,c2,c3,c4,c5,c6=st.columns(6)

            CNICcol = c1.selectbox('Select CNIC Column for Analysis !',(data.columns))
            ACCcol = c2.selectbox('Select Bank Account Title Column for Analysis !',(data.columns))
            ACCnumcol = c3.selectbox('Select Bank Account No# Column for Analysis !',(data.columns))
            Ocupaintcol = c4.selectbox('Select Occupant name Column for Analysis !',(data.columns))
            Spousecol = c5.selectbox('Select Spouse name Column for Analysis !',(data.columns))
            Fathercol = c6.selectbox('Select Father name Column for Analysis !',(data.columns))

            if st.button("Execute"):
                cout1,cout2,cout3,cout4=st.columns(4)
                cout1.dataframe(data[CNICcol].value_counts(),use_container_width=True)
                cout2.dataframe(data[CNICcol].astype(str).str.len().value_counts(),use_container_width=True)

                cout3.dataframe(data[ACCnumcol].value_counts(),use_container_width=True)
                cout4.dataframe(data[ACCnumcol].astype(str).str.len().value_counts(),use_container_width=True)

                data[ACCcol]=data[ACCcol].str.replace("ASAAN ACCOUNT","").str.replace("(","").str.replace(")","")

                data[Fathercol]=data[Fathercol].fillna("N/A")
                data[Fathercol]=data[Fathercol].str.upper()
                
                data[Spousecol]=data[Spousecol].fillna("N/A")
                data[Spousecol]=data[Spousecol].str.upper()

                data[Ocupaintcol]=data[Ocupaintcol].fillna("N/A")
                data[Ocupaintcol]=data[Ocupaintcol].str.upper()

                data[ACCcol]=data[ACCcol].fillna("N/A")
                data[ACCcol]=data[ACCcol].str.upper()

                def part_ratio(df,columns,f_col,thresh=75):
                    col_name=[]
                    for i in columns:
                        print(i)
                        col_name.append(i+' partial_ratio')
                        df[i+' partial_ratio'] = df.apply(lambda x: fuzz.partial_ratio(x[f_col], x[i]), axis=1)    
                        print(i+' partial_ratio')
                #         df[f_col+' Not Matching'] = np.where((df[i+' partial_ratio']>thresh),'Matched','Not Matched')
                        
                    condition = df[col_name].gt(70).any(axis=1)
                    df['Verification Status'] = condition.map({True: 'Matched', False: 'Not Matched'})
                    return df
                
                verify=part_ratio(data,[Ocupaintcol,Fathercol,Spousecol],ACCcol)

                st.write(verify["Verification Status"].value_counts())

                with pd.ExcelWriter("Output files/FuzzyLogic "+uploaded_files.name) as writer:
                    verify[verify["Verification Status"]=="Matched"].to_excel(writer, sheet_name='Names Matched',index=False)
                    verify[verify["Verification Status"]=="Not Matched"].to_excel(writer, sheet_name='Names Not Matched',index=False)
            
                st.success("Processing Completed Click below to download output")
                with open("Output files/FuzzyLogic "+uploaded_files.name, 'rb') as my_file:
                    st.download_button(label = 'Download output data', data = my_file, file_name = "FuzzyLogic "+uploaded_files.name, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

