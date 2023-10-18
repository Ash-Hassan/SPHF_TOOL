import streamlit as st
import pandas as pd
import time
import numpy as np
from fuzzywuzzy import fuzz
from datetime import datetime
from zipfile import ZipFile
import base64
import pyodbc

st.set_page_config(
    page_title="Disbursment (Batch Creation)",
    page_icon="📑",
    layout="wide")

st.image("Images/Consolidation Tool.png")

tab1,tab2=st.tabs(["Batch Checking","5000 Breakdown"])

with tab1:
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

    batchfile = c2.file_uploader('Upload Batch Data', type=['xlsx','xls','xlsb','csv'])
    st.session_state['batchfile']=batchfile

    blockfile = c3.file_uploader('Upload Block Data', type=['xlsx','xls','xlsb','csv'])
    st.session_state['blockfile']=blockfile

    if st.session_state['masterfile']!=None:
        if masterfile.type!="text/csv":
            masterdata=pd.ExcelFile(masterfile)
        else:
            masterdata=pd.read_csv(masterfile)

        if masterfile.type!="text/csv":
            sheetmaster = c1.selectbox(
            'Select master sheet to proceed further !',
            (masterdata.sheet_names))
            if c1.button("Load Master Sheet"):
                st.session_state['sheetmaster']=sheetmaster
                
            if st.session_state['sheetmaster']:
                masterdata=masterdata.parse(sheetmaster)
                c1.dataframe(masterdata)

    if st.session_state['batchfile']!=None:
        if batchfile.type!="text/csv":
            batchdata=pd.ExcelFile(batchfile)
        else:
            batchdata=pd.read_csv(batchfile)

        if batchfile.type!="text/csv":
            sheetbatch = c2.selectbox(
            'Select batch sheet to proceed further !',
            (batchdata.sheet_names))

            if c2.button("Load Batch Sheet"):
                st.session_state['sheetbatch']=sheetbatch
                
            if st.session_state['sheetbatch']:
                batchdata=batchdata.parse(sheetbatch)

                if "CNIC" not in batchdata.columns:
                    batchdata["CNIC"]=batchdata["CNIC without"]
                if "bank_account_number" not in batchdata.columns:
                    batchdata["bank_account_number"]=batchdata["Bank Account Number - IBAN"]
                if "uu_assessment_no" not in batchdata.columns:
                    batchdata["uu_assessment_no"]=batchdata["UD_UUID"]
                if "occupant_name" not in batchdata.columns:
                    batchdata["occupant_name"]=batchdata["DA_Occupant_Name"]
                if "spouse_name" not in batchdata.columns:
                    batchdata["spouse_name"]=batchdata["DA_SpouseName"]
                if "father_name" not in batchdata.columns:
                    batchdata["father_name"]=batchdata["DA_FatherName"]
                if "bank_account_title" not in batchdata.columns:
                    batchdata["bank_account_title"]=batchdata["Bank Account Title"]
            

                c2.dataframe(batchdata)

    if st.session_state['blockfile']!=None:
        if blockfile.type!="text/csv":
            blockdata=pd.ExcelFile(blockfile)
        else:
            blockdata=pd.read_csv(blockfile)

        if blockfile.type!="text/csv":
            sheetblock = c3.selectbox(
            'Select block sheet to proceed further !',
            (blockdata.sheet_names))

            if c3.button("Load Block Sheet"):
                st.session_state['sheetblock']=sheetblock
                
            if st.session_state['sheetblock']:
                blockdata=blockdata.parse(sheetblock)
                
                if "UD_UUID" not in blockdata.columns:
                    blockdata["UD_UUID"]=blockdata["UU ID"]

                c3.dataframe(blockdata)


    if st.session_state['masterfile']!=None and st.session_state['sheetbatch']!=None and st.session_state['blockfile']!=None:
        ad=batchdata.copy()
        amd=masterdata.copy()
        blacklist=blockdata.copy()

        progress_text ="Verifying data with Blocklist..............."
        my_bar = st.progress(0, text=progress_text)
                
        blacked=ad[ad["CNIC"].isin(list(blacklist["CNIC"]))]
        ad=ad[~ad["CNIC"].isin(list(blacklist["CNIC"]))] 
        ad.shape

        landGRM=ad[ad["landownership_type"].isin(["multiOwner","tenant","nonServeyKachaLan"])]
        ad=ad[~ad["landownership_type"].isin(["multiOwner","tenant","nonServeyKachaLan"])]
        ad.shape
        
        progress_text= "!!!!! Firingup the Engines !!!!!"
        my_bar.progress(10, text=progress_text)
        time.sleep(1)

        def find_duplicates(data,columns):
            for i in columns:
                print(i)
                data[i+' Duplicates']=np.where(ad.duplicated(subset=[i]),'Duplicate','Not Duplicate')
            return data

        def find_md_duplicates(data,amd,columns):
            for i in columns:
                print(i)
                data[i[0]+' MD_Check']=np.where(data[i[0]].isin(amd[i[1]].tolist()),'Present','Not Present')
            return data

        def part_ratio(df,columns,f_col,thresh=75):
            col_name=[]
            for i in columns:
                print(i)
                col_name.append(i+' partial_ratio')
                df[i+' partial_ratio'] = df.apply(lambda x: fuzz.partial_ratio(x[f_col], x[i]), axis=1)    
                print(i+' partial_ratio')
                #df[f_col+' Not Matching'] = np.where((df[i+' partial_ratio']>thresh),'Matched','Not Matched')

            condition = df[col_name].gt(70).any(axis=1)
            df['Verification Status'] = condition.map({True: 'Matched', False: 'Not Matched'})
            return df
        
        ad['BAN'] = ad['bank_account_number'].str[-10:]
        amd['BAN'] = amd['Bank Account number'].str[-10:]
        ad=find_duplicates(ad,['CNIC','uu_assessment_no','bank_account_number'])
        
        progress_text="!!!!! Proceding to Runway !!!!!"
        my_bar.progress(30, text=progress_text)
        time.sleep(1)
        
        ad=find_md_duplicates(ad,amd,[['CNIC','CNIC'],['uu_assessment_no','Urban Unit #'],['BAN','BAN']])
        
        progress_text="!!!!! Taking OFF !!!!!"
        my_bar.progress(50, text=progress_text)
        time.sleep(1)

        ad["occupant_name"]=ad["occupant_name"].str.upper()
        ad["spouse_name"]=ad["spouse_name"].str.upper()
        ad["father_name"]=ad["father_name"].str.upper()
        ad["bank_account_title"]=ad["bank_account_title"].str.upper()
        
        ad.fillna('-',inplace=True)

        naNames=ad[((ad["occupant_name"]=='-')|(ad["father_name"]=='-')|(ad["spouse_name"]=='-'))&(ad["bank_account_title"]!='-')]
        naBankName=ad[((ad["occupant_name"]!='-')|(ad["father_name"]!='-')|(ad["spouse_name"]!='-'))&(ad["bank_account_title"]=='-')]

        ad=ad[~((ad["bank_account_title"]=='-')|(ad["occupant_name"]=='-')|(ad["father_name"]=='-')|(ad["spouse_name"]=='-'))]
        ad.shape

        ad=part_ratio(ad,['occupant_name','father_name','spouse_name'],'bank_account_title')
        naNames=part_ratio(naNames,['occupant_name'],'bank_account_title')

        progress_text= "!!!!!Preparing to Land on Jupiter!!!!"
        my_bar.progress(75, text=progress_text)
        time.sleep(1)

        allclear=ad[ad["Verification Status"]=="Matched"]
        notclear=ad[ad["Verification Status"]=="Not Matched"]
        nonameclear=naNames[naNames["Verification Status"]=="Matched"]
        nonamehold=naNames[naNames["Verification Status"]=="Not Matched"]

        allclear["AI Status"]="All - Clear"
        notclear["AI Status"]="All - HOLD"
        nonameclear["AI Status"]="No Name - Clear"
        nonamehold["AI Status"]="No Name - HOLD"

        cleared=allclear._append(nonameclear)
        notcleared=notclear._append(nonamehold)

        sh1 =cleared[~((cleared['CNIC Duplicates']=='Duplicate' )| (cleared['uu_assessment_no Duplicates']=='Duplicate')|(cleared['bank_account_number Duplicates']=='Duplicate' )| (cleared['CNIC MD_Check']=='Present')| (cleared['uu_assessment_no MD_Check']=='Present')| (cleared['BAN MD_Check']=='Present')|(cleared['bank_account_title']=='Not Matched'))]
        sh2 = cleared[(cleared['CNIC Duplicates']=='Duplicate' )| (cleared['uu_assessment_no Duplicates']=='Duplicate')|(cleared['bank_account_number Duplicates']=='Duplicate' )| (cleared['CNIC MD_Check']=='Present')| (cleared['uu_assessment_no MD_Check']=='Present')| (cleared['BAN MD_Check']=='Present')|(cleared['bank_account_title']=='Not Matched')]
        cleared.shape
        notcleared.shape

        server = 'DESKTOP-KL8DNKI\HASSANEYSQL'
        database = 'PDMA + UU Consolidated + 12.06.2023 + 2.06M'

        # Establish the connection
        conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database)
        # Create a cursor object
        cursor = conn.cursor()

        # Execute a SELECT query
        cursor.execute('select id,[cnic],[Status of Land] from [FinalTable-WithDehs]')

        rows = cursor.fetchall()

        # Get the column names
        columns = [column[0] for column in cursor.description]

        # # Close the cursor and connection
        cursor.close()
        conn.close()

        # Create a DataFrame from the rows and columns
        df = pd.DataFrame.from_records(rows, columns=columns)
        df['cnic']=df.cnic.astype('int64')
        
        clearMap=pd.merge(sh1,df,right_on=['cnic'],left_on=["CNIC"],how="left")

        filename='Output files/Disbursment File '+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+'.xlsx'
        clearedFinal=clearMap[~((clearMap["landownership_type"]=="pvtLand")&((clearMap["Status of Land"].isin(["Uncertain Ownership",
                                                                                        "Under Litigation in a Court of Law",
                                                                                        "Government Department Land",
                                                                                        "Joint Ownership Private Land",
                                                                                        "Other Private Land",
                                                                                        "Owner willing to donate or not"])) | (clearMap["Status of Land"].isnull())))]
        
        holdFinal=clearMap[(clearMap["landownership_type"]=="pvtLand")&((clearMap["Status of Land"].isin(["Uncertain Ownership",
                                                                                        "Under Litigation in a Court of Law",
                                                                                        "Government Department Land",
                                                                                        "Joint Ownership Private Land",
                                                                                        "Other Private Land",
                                                                                        "Owner willing to donate or not"])) | (clearMap["Status of Land"].isnull()))]
        
        blackedTAGed=pd.merge(blacked,df,left_on=["CNIC"],right_on=["cnic"],how="inner")
        with pd.ExcelWriter(filename) as writer:
            clearedFinal.to_excel(writer, sheet_name="Cleaned Data - Matched", index=False)
            holdFinal.to_excel(writer, sheet_name="DC Land - Hold", index=False)
            notcleared.to_excel(writer, sheet_name="Clear filters but Not Matched", index=False)
            naBankName.to_excel(writer, sheet_name="No Data in Bank Title", index=False)
            landGRM.to_excel(writer, sheet_name="Land Title Issue", index=False)
            sh2.to_excel(writer, sheet_name="Data that needs to be verified", index=False)
            blackedTAGed.to_excel(writer, sheet_name="Blocklisted Data in batch", index=False)
            # amd.to_excel(writer, sheet_name="Master Data", index=False)
            # ad.to_excel(writer, sheet_name="Batch Data", index=False)
            # blacklist.to_excel(writer, sheet_name="Blocklist Data", index=False)

        progress_text="!!!!! Successfully Landed on Jupiter !!!!"
        my_bar.progress(100, text=progress_text)
        time.sleep(1)
    
        st.success("Processing Completed Click below to download output")
        with open(filename, 'rb') as my_file:
            st.download_button(label = 'Download output data', data = my_file, file_name = filename, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

with tab2:
    
    if 'file' not in st.session_state:
        st.session_state['file']=None

    excelfile = st.file_uploader("Upload an Excel File", type=["xlsx", "xls"])
    st.session_state['file']=excelfile

    if st.session_state['file']!=None:
        data = pd.ExcelFile(st.session_state['file'])
        sheetname = st.selectbox(
        'Select sheet to proceed further !',
        (data.sheet_names))
        df=data.parse(sheetname)
        st.dataframe(df,use_container_width=True)

        if st.button("Execute Process"):
            zipfilename="Output files/Output"+str(datetime.now().strftime("%Y-%m-%d %H-%M-%S"))+".zip"
            with st.spinner('Wait for it...'):
                files={}
                zipObj = ZipFile(zipfilename, "w")
                print(excelfile.name)

                num_files = -(-len(df) // 5000)

                smaller_dfs = [df[i*5000:(i+1)*5000] for i in range(num_files)]

                batchcount=0
                for i, smaller_df in enumerate(smaller_dfs):

                    print(i)
                    if i==num_files-1:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +smaller_df.shape[0]}.xlsx'
                    else:
                        file = f'Output files/{excelfile.name.split(".")[0]} {batchcount+1}-{batchcount +5000}.xlsx'
                    
                    batchcount+=5000
                    print(file)
            #             smaller_df.to_excel(file)

                    writer = pd.ExcelWriter(file, engine='xlsxwriter')
                    smaller_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    writer.close()

                    files.update({file:smaller_dfs})

                    zipObj.write(file)

                # close the Zip File
                zipObj.close()

                ZipfileDotZip = zipfilename

                with open(ZipfileDotZip, "rb") as f:
                    bytes = f.read()
                    b64 = base64.b64encode(bytes).decode()
                    href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
                        Click here to download the output\
                    </a>"

                st.markdown(href, unsafe_allow_html=True)

                [datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),"Username",excelfile.name,"Output Files : "+str(num_files),files]

    else:
        st.warning('Upload input data to proceed ! ', icon="⚠️")