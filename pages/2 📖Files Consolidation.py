import streamlit as st
import pandas as pd
import json
from time import sleep
import numpy as np
import re
from datetime import datetime

st.set_page_config(
    page_title="Files Consolidation",
    page_icon="📖",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'file' not in st.session_state:
    st.session_state['file']=None
    
uploaded_files = st.file_uploader('Upload Input Data', type=['xlsx','xls','xlsb'],accept_multiple_files=True)
st.session_state['file']=uploaded_files
tab1, tab2, = st.tabs(["Simple Consolidation","Dictionary Based Consolidation"])
print(uploaded_files)
with tab1:
    if st.session_state['file']!=None:
        if st.button("Execute Process"):
            with st.spinner('Wait for it...'):
                concatenated_data = pd.DataFrame()
                for file_name in uploaded_files:
                    if file_name.name.endswith('.xlsx') or file_name.name.endswith('.xls'):
                        print(file_name.name)

                        # Read the Excel file into a DataFrame
                        df = pd.read_excel(file_name)
                        df["File Source"]=file_name.name

                        # Concatenate the DataFrame to the existing data
                        concatenated_data = pd.concat([concatenated_data, df], ignore_index=True)

                # Save the concatenated data to a new Excel file
                output_file = 'Output files/'+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S "))+'_concatenated_data.xlsx'
                concatenated_data.to_excel(output_file, index=False)
            st.success("Processing Completed Click below to download output")

            with open(output_file, 'rb') as my_file:
                st.download_button(label = 'Download output data', data = my_file, file_name = output_file, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    else:
        st.warning('Upload input data to proceed ! ', icon="⚠️")


with tab2:
    with open('ToolData/data.json', 'r') as fp:
        possible_columns_dict = json.load(fp)

    with open('ToolData/BankMap.json', 'r') as fp:
        bankMAP = json.load(fp)

    with open('ToolData/mapping.json', 'r') as fp:
        mapping = json.load(fp)

    with open('ToolData/Initials.json', 'r') as fp:
        initals = json.load(fp)

    column_mapping = {}

    for main_column, possibilities in possible_columns_dict.items():
        for possible_column in possibilities:
            column_mapping[possible_column] = main_column
    if st.session_state['file']!=None:
        if st.button("Execute"):
            file_list = uploaded_files#[x.name for x in uploaded_files]

            possible_columns = list(column_mapping.keys())

            total=0

            summary=pd.DataFrame(columns=["File Name","Sheet_Name","Total_Count","CNIC_Column","CNIC_Count","Name_Column","Name_Count","ACCOUNT NO_Column","ACCOUNT NO_Count","IBAN Numbers_Column","IBAN Numbers_Count","Branch Name_Column","Branch Name_Count","Branch Code_Column","Branch Code_Count","Bank name_Column","Bank name_Count","Available Columns","Selected Columns"])

            concatenated_data = pd.DataFrame()

            progress_text ="!!!!! Firingup the Engines !!!!!"
            my_bar = st.progress(0, text=progress_text)

            for i, file_name in enumerate(file_list, 1):
                    progress = int((i / len(file_list)) * 100)

                    if progress < 20:
                        progress_text="!!!!! Firingup the Engines !!!!!"
                        sleep(1)
                    elif progress < 30:
                        progress_text="!!!!! Proceding to Runway !!!!!"
                        sleep(1)
                    elif progress < 50:
                        progress_text="!!!!! Taking OFF !!!!!"
                        sleep(1)
                    elif progress < 75:
                        progress_text="!!!!! Preparing to Land on Mars !!!!"
                        sleep(1)
                    elif progress == 100:
                        progress_text="!!!!! Successfully Landed on Mars !!!!"
                        sleep(1)

                    my_bar.progress(progress, text=progress_text)

                    print("File Loaded: ",file_name.name)

                    excel_data = pd.ExcelFile(file_name)
                    for sheet_name in excel_data.sheet_names:
                        print(f"Sheet: {sheet_name}")
                        data = excel_data.parse(sheet_name)
                        if data.empty:
                            continue

                        #display(data)
                        if data.columns.astype(str).str.contains('Unnamed').sum() > 5:
                            print("Mapping Columns")
                            data = data.dropna(thresh=(len(data.columns) * 0.9))
                            if data.empty:
                                continue
                            data.columns = data.iloc[0]
                            data = data[1:]
                            data.reset_index(drop=True, inplace=True)

                        print("Original Columns: ",list(data.columns))

                        subset_columns = [col for col in possible_columns if col in data.columns]

                        selected_data = data[subset_columns]

                        print("Selected Columns: ",list(subset_columns))

                        total+=len(data)

                        original_columns = selected_data.columns.tolist()

                        selected_data=selected_data.rename(columns=column_mapping)

                        if "CNIC" not in list(selected_data.columns):
                            continue   

                        selected_data.loc[0.5]=original_columns

                        selected_data=selected_data.loc[:, ~selected_data.columns.duplicated()]

                        selected_data=selected_data[selected_data["CNIC"].notna()]

                        selected_data["CNIC"]=selected_data["CNIC"].astype(str).str.replace('-','')

                        missing_columns = set(column_mapping.values()) - set(selected_data.columns)
                        if missing_columns:
                            for col in missing_columns:
                                selected_data[col] = "" 

                        selected_data['ACCOUNT NO'] = selected_data["ACCOUNT NO"].astype(str).str.replace(" ","").str.lower()

                        selected_data["IBAN Numbers"]=np.where(selected_data["ACCOUNT NO"].str.contains("pk"),selected_data["ACCOUNT NO"],selected_data["IBAN Numbers"])
                        selected_data["ACCOUNT NO"]=np.where(selected_data["ACCOUNT NO"].str.contains("pk"),"",selected_data["ACCOUNT NO"])

                        selected_data['IBAN Numbers'] = selected_data["IBAN Numbers"].astype(str).str.replace(" ","").str.lower()



                        selected_data=selected_data.reset_index(drop=True)
                        selected_data=selected_data[["CNIC","Name","ACCOUNT NO","IBAN Numbers","Branch Name","Branch Code","Bank name"]]

            #             display(selected_data)

                        newSumary=pd.DataFrame([[file_name,sheet_name,len(data)+1,
                        list(selected_data["CNIC"])[-1],len(set(selected_data["CNIC"])),
                        list(selected_data["Name"])[-1],len(set(selected_data["Name"])),
                        list(selected_data["ACCOUNT NO"])[-1],len(set(selected_data["ACCOUNT NO"])),
                        list(selected_data["IBAN Numbers"])[-1],len(set(selected_data["IBAN Numbers"])),
                        list(selected_data["Branch Name"])[-1],len(set(selected_data["Branch Name"])),
                        list(selected_data["Branch Code"])[-1],len(set(selected_data["Branch Code"])),
                        list(selected_data["Bank name"])[-1],len(set(selected_data["Bank name"])),list(data.columns),list(subset_columns)]],columns=summary.columns)

                        summary=pd.concat([summary,newSumary],ignore_index=True)

                        concatenated_data = pd.concat([concatenated_data, selected_data.iloc[:-1]],ignore_index=True)

            #****************************************Verification*****************************************    
            duplicates_CNIC = concatenated_data[concatenated_data.duplicated(subset='CNIC')]
            unique_CNICdata = concatenated_data.drop_duplicates(subset='CNIC')
            nan_IBAN = unique_CNICdata[unique_CNICdata["IBAN Numbers"]==""]
            noNAN_IBAN = unique_CNICdata[unique_CNICdata["IBAN Numbers"]!=""]
            duplicates_IBAN = noNAN_IBAN[noNAN_IBAN.duplicated(subset='IBAN Numbers')]
            unique_IBANdata = noNAN_IBAN.drop_duplicates(subset='IBAN Numbers')

            #****************************************Correct Ibans*****************************************    
            unique_IBANdata['IBAN Numbers']=unique_IBANdata['IBAN Numbers'].str.replace(",","")
            unique_IBANdata["IBAN BANK"]=unique_IBANdata['IBAN Numbers'].str.extract(r"\d+([a-zA-Z]+)\d+")
            unique_IBANdata["IBAN BANK"]=unique_IBANdata['IBAN BANK'].replace(mapping)
            unique_IBANdata["IBAN prefix"]=unique_IBANdata['IBAN Numbers'].str.extract(r'(pk\d+)', flags=re.IGNORECASE)
            unique_IBANdata['IBAN AccNo']=unique_IBANdata['IBAN Numbers'].str.extract(r'[A-Za-z]+\d+[A-Za-z]+(\d+)')
            unique_IBANdata["Mapped IBAN Numbers"]=unique_IBANdata["IBAN prefix"]+unique_IBANdata["IBAN BANK"]+unique_IBANdata['IBAN AccNo']
            unique_IBANdata["Mapped BANK NAME"]=unique_IBANdata['IBAN BANK'].replace(bankMAP)
            unique_IBANdata["Bank initails"]=unique_IBANdata["Mapped BANK NAME"].map(initals)
            incorrect_IBAN=unique_IBANdata[unique_IBANdata["Mapped IBAN Numbers"].isna()]
            unique_IBANdata=unique_IBANdata[unique_IBANdata["Mapped IBAN Numbers"].notna()].copy()
            

            #****************************************IBAN Len Check*****************************************
            unique_IBANdata["IBAN Len"]=unique_IBANdata["Mapped IBAN Numbers"].str.len()
            correctData=unique_IBANdata[unique_IBANdata["IBAN Len"]==24]
            wrongiban=unique_IBANdata[unique_IBANdata["IBAN Len"]!=24]

            #****************************************Data Summary*****************************************    
            DataSummary=pd.DataFrame(columns=["Verification","Duplicates","Unique"])
            DataSummary.loc[len(DataSummary)] = ["CNIC",duplicates_CNIC.shape[0],unique_CNICdata.shape[0]]
            DataSummary.loc[len(DataSummary)] = ["IBAN Null",nan_IBAN.shape[0],noNAN_IBAN.shape[0]]
            DataSummary.loc[len(DataSummary)] = ["IBAN",duplicates_IBAN.shape[0],unique_IBANdata.shape[0]]
            DataSummary.loc[len(DataSummary)] = ["IBAN Length",wrongiban.shape[0],correctData.shape[0]]


            duplicates=pd.DataFrame()
            duplicates=pd.concat([duplicates,duplicates_CNIC])
            duplicates=pd.concat([duplicates,duplicates_IBAN])
            incorrect=pd.DataFrame()
            incorrect=pd.concat([incorrect,nan_IBAN])
            incorrect=pd.concat([incorrect,incorrect_IBAN])
            incorrect=pd.concat([incorrect,wrongiban])


            summary=summary.replace("","Not Found")
            summary=summary.where(summary.ne(1), 0)
            def highlight_data(x):
                if x == 'Not Found' or x==0:
                    return 'background-color: red'
                elif type(x)==int:
                    if x < 10:
                        return 'background-color: orange'

            styled_summary = summary.style.applymap(highlight_data)

            outputFileName='Output files/Consolidation File '+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+'.xlsx'
            with pd.ExcelWriter(outputFileName, engine='xlsxwriter') as writer:
                styled_summary.to_excel(writer,sheet_name='Summary',index=False)
                DataSummary.to_excel(writer,sheet_name='Data Summary',index=False)
                correctData.to_excel(writer,sheet_name='Correct Data',index=False)
                duplicates.to_excel(writer,sheet_name='Duplicates Data',index=False)
                incorrect.to_excel(writer,sheet_name='incorrect Data',index=False)

            st.success("Processing Completed Click below to download output")

            with open(outputFileName, 'rb') as my_file:
                st.download_button(label = 'Download output data', data = my_file, file_name = outputFileName, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

