import streamlit as st
import pandas as pd
from zipfile import ZipFile
from io import BytesIO
import base64
import datetime

st.set_page_config(
    page_title="File Segmentation",
    page_icon="📚",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'file' not in st.session_state:
    st.session_state['file']=None

excelfile = st.file_uploader("Upload an Excel File", type=["xlsx", "xls"])
st.session_state['file']=excelfile

tab1,tab2=st.tabs["Normal Segmentation","W/O Date Segmentation"]

if st.session_state['file']!=None:
    data = pd.ExcelFile(st.session_state['file'])
    sheetname = st.selectbox(
    'Select sheet to proceed further !',
    (data.sheet_names))
    df=data.parse(sheetname,dtype={'From Account': object})
    st.dataframe(df,use_container_width=True)

    with tab1:
        if st.button("Execute Process"):
            with st.spinner('Wait for it...'):
                files={}
                zipObj = ZipFile("Output files/Output.zip", "w")
                print(excelfile.name)

                num_files = -(-len(df) // 200)

                smaller_dfs = [df[i*200:(i+1)*200] for i in range(num_files)]

                batchcount=0
                for i, smaller_df in enumerate(smaller_dfs):

                    print(i)
                    if i==num_files-1:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +smaller_df.shape[0]}.xlsx'
                    else:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +200}.xlsx'
                    
                    batchcount+=200
                    print(file)
            #             smaller_df.to_excel(file)

                    writer = pd.ExcelWriter(file, engine='xlsxwriter')
                    smaller_df["DateTime"]=smaller_df["DateTime"].dt.strftime("%d-%B-%Y").astype(str)
                    smaller_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    number_format = workbook.add_format({'num_format': '0'})  # Whole number format

                    # Apply formatting to specific columns
                    worksheet.set_column('B:B', None, number_format)  # Formatting column A as date
                    writer.close()

                    files.update({file:smaller_dfs})

                    zipObj.write(file)

                # close the Zip File
                zipObj.close()

                ZipfileDotZip = "Output files/Output.zip"

                with open(ZipfileDotZip, "rb") as f:
                    bytes = f.read()
                    b64 = base64.b64encode(bytes).decode()
                    href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
                        Click here to download the output\
                    </a>"

                st.markdown(href, unsafe_allow_html=True)

                [datetime.datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),"Username",excelfile.name,"Output Files : "+str(num_files),files]

    with tab2:
        if st.button("Execute Process"):
            with st.spinner('Wait for it...'):
                files={}
                zipObj = ZipFile("Output files/Output.zip", "w")
                print(excelfile.name)

                num_files = -(-len(df) // 200)

                smaller_dfs = [df[i*200:(i+1)*200] for i in range(num_files)]

                batchcount=0
                for i, smaller_df in enumerate(smaller_dfs):

                    print(i)
                    if i==num_files-1:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +smaller_df.shape[0]}.xlsx'
                    else:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +200}.xlsx'
                    
                    batchcount+=200
                    print(file)
            #             smaller_df.to_excel(file)

                    writer = pd.ExcelWriter(file, engine='xlsxwriter')
                    smaller_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    number_format = workbook.add_format({'num_format': '0'})  # Whole number format

                    # Apply formatting to specific columns
                    worksheet.set_column('B:B', None, number_format)  # Formatting column A as date
                    writer.close()

                    files.update({file:smaller_dfs})

                    zipObj.write(file)

                # close the Zip File
                zipObj.close()

                ZipfileDotZip = "Output files/Output.zip"

                with open(ZipfileDotZip, "rb") as f:
                    bytes = f.read()
                    b64 = base64.b64encode(bytes).decode()
                    href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
                        Click here to download the output\
                    </a>"

                st.markdown(href, unsafe_allow_html=True)

                [datetime.datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),"Username",excelfile.name,"Output Files : "+str(num_files),files]


else:
    st.warning('Upload input data to proceed ! ', icon="⚠️")