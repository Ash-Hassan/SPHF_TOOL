import streamlit as st
import pandas as pd
import json
from time import sleep
import numpy as np
import re
from datetime import datetime
from fuzzywuzzy import fuzz
from qvd import qvd_reader

st.set_page_config(
    page_title="Report Analysis",
    page_icon="📈",
    layout="wide")

st.image("Images/Consolidation Tool.png")

if 'qlikfile' not in st.session_state:
    st.session_state['qlikfile']=None

qlikData=st.file_uploader('Upload Report Only', type=['xlsx','xls','xlsb','csv','qvd'])
st.session_state['qlikfile']=qlikData

if st.session_state['qlikfile']!=None and st.button("Execute Process"):

    progress_text ="**** Report chalaye ga mera beta? ****"
    my_bar = st.progress(0, text=progress_text)
    sleep(2)

    IPdistricts={
    "Dadu":"TRDP","Jamshoro":"TRDP","Tharparkar":"TRDP",
    "Naushahro Feroze":"SAFCO","Shaheed Benazirabad":"SAFCO",        
    "Kashmore":"HANDS","Thatta":"HANDS","Umerkot":"HANDS","Sukkur":"HANDS",
    "Karachi":"HANDS","Hyderabad":"HANDS","Ghotki (At Mirpur Mathelo)":"HANDS",
    "Badin":"NRSP","Matiari":"NRSP","Sujawal":"NRSP","Tando Allahyar":"NRSP",
    "Tando Mohammad Khan":"NRSP","Mirpur Khas":"NRSP","Sanghar":"NRSP",
    "Jacobabad":"SRSO","Qambar Shahdadkot":"SRSO","Larkana":"SRSO","Shikarpur":"SRSO","Khairpur":"SRSO"
}
    
    try:
        if st.session_state['qlikfile']!=None:
            if qlikData.type!="text/csv":
                maindata=pd.read_excel(st.session_state['qlikfile'])
            else:
                maindata=pd.read_csv(st.session_state['qlikfile'])
    except Exception as e:
        print("File excel nhe",e)
        with open(qlikData.name,'wb') as f:
            f.write(qlikData.read())
        maindata=qvd_reader.read(qlikData.name)
    try:
        maindata["DistrictMAP"]=maindata["district"].map(IPdistricts)
    except:
        maindata["DistrictMAP"]=maindata["District"].map(IPdistricts)


    maindataOG=maindata.copy()

    st.markdown("Actual Data: "+str(maindataOG.shape[0]))

    my_bar.progress(10, text="!!! dekho ab mai kya krta hu  !!!")
    sleep(2)

    # CNIC Invalids
    st.markdown("Data Supplied for cnic verification: "+str(maindata.shape[0]))

    maindata["verify_DA_CNIC"]=maindata["DA_CNIC"].str.replace("-","")
    validCNIC1=maindata[maindata["verify_DA_CNIC"].str.len()==13]
    print(validCNIC1.shape[0])
    invalidCNIC1=maindata[maindata["verify_DA_CNIC"].str.len()!=13]
    print(invalidCNIC1.shape[0])

    validCNIC2=validCNIC1[validCNIC1["DA_CNIC"].str[0].isin(["1","2","3","4","5","6","7","8","9"])]
    print(validCNIC2.shape[0])
    invalidCNIC2=validCNIC1[~validCNIC1["DA_CNIC"].str[0].isin(["1","2","3","4","5","6","7","8","9"])]
    print(invalidCNIC2.shape[0])

    # validCNIC2
    invalidCNIC=pd.concat([invalidCNIC1,invalidCNIC2])

    # Deviation in CNIC
    clearedCNIC=validCNIC2[validCNIC2["CNIC Deviation"]=="No Deviation"]
    print(clearedCNIC.shape)
    holdCNIC=validCNIC2[validCNIC2["CNIC Deviation"]=="Deviation"]
    print(holdCNIC.shape)

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
    holdCNIC
    if holdCNIC.empty:
        holdCNIC_clearedName=pd.DataFrame()
        holdCNIC_holdName=pd.DataFrame()
    else:
        holdCNIC=part_ratio(holdCNIC,['U_Occupant Name'],'DA Occupant Name')
        holdCNIC_clearedName=holdCNIC[holdCNIC["DA Occupant Name Not Matching"]=="Matched"]
        print(holdCNIC_clearedName.shape)
        holdCNIC_holdName=holdCNIC[holdCNIC["DA Occupant Name Not Matching"]=="Not Matched"]
        print(holdCNIC_holdName.shape)

    if holdCNIC_holdName.empty:
        holdCNIC_holdName_clearedSFname=pd.DataFrame()
        holdCNIC_holdName_holdSFname=pd.DataFrame()
    else:
        holdCNIC_holdName=part_ratio(holdCNIC_holdName,['U_Occupant Name'],'DA Spouse Name')

        holdCNIC_holdName_clearedSFname=holdCNIC_holdName[holdCNIC_holdName["DA Spouse Name Not Matching"]=="Matched"]
        print(holdCNIC_holdName_clearedSFname.shape)
        holdCNIC_holdName_holdSFname=holdCNIC_holdName[holdCNIC_holdName["DA Spouse Name Not Matching"]=="Not Matched"]
        print(holdCNIC_holdName_holdSFname.shape)

    cnicGRM=holdCNIC_holdName_holdSFname.copy()

    # clearedCNIC | holdCNIC_clearedName | holdCNIC_holdName_clearedSFname
    # holdCNIC -> holdCNIC_holdName -> holdCNIC_holdName_holdSFname

    cnicCLEARED=pd.concat([clearedCNIC,holdCNIC_clearedName,holdCNIC_holdName_clearedSFname])

    st.markdown("Data Cleared cnic verification: "+str(cnicCLEARED.shape[0]))
    st.markdown("Data Failed cnic verification: "+str(cnicGRM.shape[0]+invalidCNIC.shape[0]))

    my_bar.progress(25, text="!!! CNIC Verified aur kya chaye?  !!!")
    sleep(2)

    if validCNIC2.shape[0]==cnicCLEARED.shape[0]+cnicGRM.shape[0]:
        print("No data missing GO to step 2")

        print(cnicGRM.shape)
        print(cnicCLEARED.shape)

    ## Land Title
    Holdland=cnicCLEARED[cnicCLEARED["landownership_type"].isin(["nonServeyKachaLan","multiOwner","tenant"])]
    print(Holdland.shape)
    clearedLAND=cnicCLEARED[~cnicCLEARED["landownership_type"].isin(["nonServeyKachaLan","multiOwner","tenant"])]
    print(clearedLAND.shape)

    st.markdown("Data Cleared Land Titles: "+str(clearedLAND.shape[0]))
    st.markdown("Data Failed Land Titles: "+str(Holdland.shape[0]))

    # landGRM 
    # clearedLAND

    my_bar.progress(35, text="!!! Chalo Land Titles bhi hogaye...  !!!")
    sleep(2)

    ## Reconstrution
    hold3rd=clearedLAND[clearedLAND["Has_Reconstruction_Started"]!="no"]
    print(hold3rd.shape)
    cleared3rd=clearedLAND[clearedLAND["Has_Reconstruction_Started"]=="no"]
    print(cleared3rd.shape)

    st.markdown("Data Cleared Reconstrution Status: "+str(cleared3rd.shape[0]))
    st.markdown("Data Failed Reconstrution Status: "+str(hold3rd.shape[0]))

    # hold3rd
    # cleared3rd

    my_bar.progress(45, text="!!! Reconstruction bhi dekh lya bs kro ab...  !!!")
    sleep(2)

        # Damage assesment
    cleared3rd["UD_Damage CategorySTD"]=cleared3rd["UD_Damage Category"].str.replace("visible","Partially").replace("partially","Partially").replace("Kacha","Fully").replace("Pakka","Partially")
    cleared3rd["DA_Damage CategorySTD"]=cleared3rd["UD_Damage Category"].str.replace("collapsed","Fully").replace("washedAway","Fully").replace("visible","Partially").replace("partially","Partially")
    cleared3rd["DA TypeSTD"]=cleared3rd["DA Type"].str.replace("hybrid","Kacha").replace("other","Kacha").replace("kacha","Kacha").replace("pucca","Pakka")
    cleared3rd["UU TypeSTD"]=cleared3rd["UU Type"].str.replace("Thatch","Kacha").replace("Plinth (Concrete-Super structure \\n Brick in mud mortar)","Kacha").replace("Plinth (Concrete-Super structure mud)","Kacha").replace("partially","Pakka").replace("partially}]","Pakka")

    cleared3rd["Combined"]=cleared3rd["DA TypeSTD"]+"_"+cleared3rd["UU TypeSTD"]+"_"+cleared3rd["DA_Damage CategorySTD"]+"_"+cleared3rd["UD_Damage CategorySTD"]


    CatMap={
        "Kacha_Kacha_Fully_Fully":"Category A", 
        "Kacha_Kacha_Fully_Partially":"Category B",
        "Kacha_Pakka_Fully_Fully":"Category C",
        "Pakka_Kacha_Fully_Fully":"Category D",
        "Kacha_Kacha_Partially_Fully":"Category E",
        "Pakka_Pakka_Fully_Fully":"Category F",
        "Kacha_Kacha_Partially_Partially":"Category G",
        "Pakka_Kacha_Fully_Partially":"Category H",
        "Kacha_Pakka_Partially_Fully":"Category I",

        "Kacha_Pakka_Fully_Partially":"GRM Category1",
        "Pakka_Pakka_Fully_Partially":"GRM Category2",
        "Pakka_Pakka_Partially_Partially":"GRM Category3",
        "Kacha_Pakka_Partially_Partially":"GRM Category4",
        "Pakka_Kacha_Partially_Fully":"GRM Category5",
        "Pakka_Pakka_Partially_Fully":"GRM Category6",
        "Pakka_Kacha_Partially_Partially":"GRM Category7",
        np.nan:"GRM Category8",

        "Kacha_Kacha_intact_Fully":"GRM Category9",
        "Kacha_Kacha_intact_Partially":"GRM Category10",
        "Kacha_Pakka_intact_Fully":"GRM Category11",
        "Kacha_Pakka_intact_Partially":"GRM Category12",
        "Pakka_Kacha_intact_Fully":"GRM Category13",
        "Pakka_Pakka_intact_Partially":"GRM Category14",
        "Pakka_Pakka_intact_Fully":"GRM Category15",
        "Pakka_Kacha_intact_Partially":"GRM Category16",


    }

    cleared3rd["Category"]=cleared3rd["Combined"].map(CatMap)
    cleared3rd
    DamagedCleared=cleared3rd[~(cleared3rd["Category"].isin(["GRM Category1","GRM Category9",
                                                    "GRM Category2","GRM Category10",
                                                    "GRM Category3","GRM Category11",
                                                    "GRM Category4","GRM Category12",
                                                    "GRM Category5","GRM Category13",
                                                    "GRM Category6","GRM Category14",
                                                    "GRM Category7","GRM Category15",
                                                    "GRM Category8","GRM Category16"
                                                        ]))]
    print(DamagedCleared.shape)
    DamagedGRM=cleared3rd[cleared3rd["Category"].isin(["GRM Category1","GRM Category9",
                                                    "GRM Category2","GRM Category10",
                                                    "GRM Category3","GRM Category11",
                                                    "GRM Category4","GRM Category12",
                                                    "GRM Category5","GRM Category13",
                                                    "GRM Category6","GRM Category14",
                                                    "GRM Category7","GRM Category15",
                                                    "GRM Category8","GRM Category16"
                                                        ])]
    print(DamagedGRM.shape)

    st.markdown("Data Cleared Damaged Assesment: "+str(DamagedCleared.shape[0]))
    st.markdown("Data Failed Damaged Assesment: "+str(DamagedGRM.shape[0]))

    my_bar.progress(70, text="!!! Ye kacha pucca kya hai? Sans phool gayi kacha pucca krty krty  !!!")
    sleep(2)

    DamagedClearedNoDup=DamagedCleared.sort_values("is_hazardous_location")
    DamagedClearedNoDup=DamagedClearedNoDup.drop_duplicates(subset="DA_CNIC",keep='first')
    print(DamagedClearedNoDup.shape)

    clearedHazard=DamagedCleared[DamagedCleared["is_hazardous_location"]=="no"]
    print(clearedHazard.shape)
    holdHazard=DamagedCleared[DamagedCleared["is_hazardous_location"]!="no"]
    print(holdHazard.shape)

    st.markdown("Data Not in Hazaderous (with - Duplicates): "+str(DamagedCleared.shape[0]))
    st.markdown("Data in Hazaderous (with - Duplicates): "+str(DamagedGRM.shape[0]))

    clearedHazardNoDup=DamagedClearedNoDup[DamagedClearedNoDup["is_hazardous_location"]=="no"]
    print(clearedHazardNoDup.shape)
    holdHazardNoDup=DamagedClearedNoDup[DamagedClearedNoDup["is_hazardous_location"]!="no"]
    print(holdHazardNoDup.shape)

    st.markdown("Data Not in Hazaderous (No - Duplicates): "+str(DamagedCleared.shape[0]))
    st.markdown("Data in Hazaderous (No - Duplicates): "+str(DamagedGRM.shape[0]))

    my_bar.progress(80, text="!!! Bs kro AI ki jaan logy ab? sb done file banaing pakro aur niklo..  !!!")
    sleep(2)

    def mapSummary(df1,df2):
        districtt=df1[["DistrictMAP","District","DA_CNIC"]].groupby(["DistrictMAP","District"]).agg('count').reset_index()
        districtt=districtt.merge(df2[["District","DA_CNIC"]].groupby(["District"]).agg('count').reset_index(),on="District")
        districtt.columns=["IpName","District","Total Case","Cleared"]
        districtt["% cleared"]=districtt["Cleared"]/districtt["Total Case"]
        districtt["IpName"]=""
        districtt=districtt[["IpName","District","Total Case","Cleared","% cleared"]]

        return districtt
    
    df=pd.DataFrame()
    df['Data Type'] =['Total Cases','Cleared without deviation', 'Invalid CNIC','Cleared on Occupant name as Beneficiary','Deviation in CNIC No# & NAME','Cleared on Spouse as Beneficiary','Residual Cases','GRM CNIC Cases']
    df['Data'] = [maindata.shape[0],clearedCNIC.shape[0],invalidCNIC.shape[0],holdCNIC_clearedName.shape[0],holdCNIC_holdName.shape[0],holdCNIC_holdName_clearedSFname.shape[0],cnicGRM.shape[0],cnicGRM.shape[0]+invalidCNIC.shape[0]]

    df.loc[6.5]=["",""]
    df.loc[9]=["Total Cleared",df.loc[1.0]["Data"]+df.loc[3.0]["Data"]+df.loc[5.0]["Data"]]
    df=df.sort_index()
    
    df=df._append(mapSummary(maindata,cnicCLEARED))

    landALL=pd.DataFrame(cnicCLEARED["landownership_type"].value_counts()).reset_index()
    landALL=landALL._append(landALL.sum(numeric_only=True), ignore_index=True)
    landALL=landALL.fillna("Total")
    landGRM=Holdland["landownership_type"].value_counts().reset_index()
    landGRM=landGRM._append(landGRM.sum(numeric_only=True), ignore_index=True)
    landGRM=landGRM.fillna("Total")
    landClear=clearedLAND["landownership_type"].value_counts().reset_index()
    landClear=landClear._append(landClear.sum(numeric_only=True), ignore_index=True)
    landClear=landClear.fillna("Total")
    landClear.columns=["landownership_type","Total"]
    
    landClear=landClear._append(mapSummary(cnicCLEARED,clearedLAND))

    df3rd=pd.DataFrame()
    df3rd["Data Type"]=["Reconstruction started - Yes","Reconstruction started - No"]
    df3rd["Data"]=[hold3rd.shape[0],cleared3rd.shape[0]]
    df3rd=df3rd._append(df3rd.sum(numeric_only=True), ignore_index=True)
    df3rd=df3rd.fillna("Total")
    
    df3rd=df3rd._append(mapSummary(clearedLAND,cleared3rd))
    
    Damaged=pd.DataFrame(cleared3rd["Category"].value_counts()).reset_index()
    Damaged
    Damaged=Damaged._append(Damaged.sum(numeric_only=True), ignore_index=True)
    Damaged=Damaged.fillna("Total")
    Damaged=Damaged.sort_values("Category")
    Damaged["MAP"]=Damaged["Category"].map({value: key for key, value in CatMap.items()})
    Damaged=Damaged.join(Damaged['MAP'].str.split('_',expand=True).add_prefix('Data_'))
    Damaged=Damaged[["Data_0","Data_2","Data_1","Data_3","Category","count"]]
    Damaged.columns=["MIS Structure","MIS Damaged Category","JV Structure","JV Damaged Category","Marked as","Total"]
    
    Damaged=Damaged._append(mapSummary(cleared3rd,DamagedCleared))
    Level1Summary=pd.DataFrame.from_dict({
        'Total Cases':maindata.shape[0],
        'Total Cleared':DamagedCleared.shape[0],
        'Total GRM':DamagedGRM.shape[0]
                                        +
        df3rd[df3rd["Data Type"]=="Reconstruction started - Yes"]["Data"].item()
                                        +
        landGRM[landGRM["landownership_type"]=="Total"]["count"].item()
                                        +
        df[df["Data Type"]=="GRM CNIC Cases"]["Data"].item()
    }.items())
    
    Level1Summary=Level1Summary._append(mapSummary(cleared3rd,DamagedCleared))

    hazardSum=pd.DataFrame()
    hazardSum["Data Type"]=["Hazardeous Location - No","Hazardeous Location - Yes"]
    hazardSum["Data"]=[clearedHazard.shape[0],holdHazard.shape[0]]
    hazardSum=hazardSum._append(hazardSum.sum(numeric_only=True), ignore_index=True)
    hazardSum=hazardSum.fillna("Total")

    hazardSum=hazardSum._append(mapSummary(DamagedCleared,clearedHazard))

    hazardSumNoDup=pd.DataFrame()
    hazardSumNoDup["Data Type"]=["Total Duplicates","Duplicates Removed","Hazardeous Location - No","Hazardeous Location - Yes"]
    hazardSumNoDup["Data"]=[
        DamagedCleared[DamagedCleared.duplicated(subset="DA_CNIC",keep=False)].shape[0],
        DamagedCleared[DamagedCleared.duplicated(subset="DA_CNIC",keep='first')].shape[0],
        clearedHazardNoDup.shape[0],
        holdHazardNoDup.shape[0]
    ]
    hazardSumNoDup=hazardSumNoDup._append(hazardSumNoDup.sum(numeric_only=True), ignore_index=True)
    hazardSumNoDup=hazardSumNoDup.fillna("Total")
    
    hazardSumNoDup=hazardSumNoDup._append(mapSummary(DamagedCleared,clearedHazardNoDup))

    Level2Summary=pd.DataFrame.from_dict({
        'Total Cases':maindata.shape[0],
        'Total Cases Cleared':clearedHazardNoDup.shape[0],#Hazardeous Location - No
        'Total Cases in Hazardeous':holdHazardNoDup.shape[0],
        'Total Duplicates Removed':DamagedCleared[DamagedCleared.duplicated(subset="DA_CNIC",keep='first')].shape[0],
        'Total GRM':DamagedGRM.shape[0]
                                        +
        df3rd[df3rd["Data Type"]=="Reconstruction started - Yes"]["Data"].item()
                                        +
        landGRM[landGRM["landownership_type"]=="Total"]["count"].item()
                                        +
        df[df["Data Type"]=="GRM CNIC Cases"]["Data"].item()
    }.items())
    
    Level2Summary.columns=["Data Type","Data"]
    Level2Summary=Level2Summary._append(mapSummary(DamagedCleared,clearedHazardNoDup))
    Level2Summary
    my_bar.progress(90, text="!!! Ruko zara sbr kro  !!!")
    sleep(2)

    st.markdown(Damaged[Damaged["Marked as"].isin(["GRM Category1","GRM Category9"
                                                    "GRM Category2","GRM Category10"
                                                    "GRM Category3","GRM Category11"
                                                    "GRM Category4","GRM Category12"
                                                    "GRM Category5","GRM Category13"
                                                    "GRM Category6","GRM Category14"
                                                    "GRM Category7","GRM Category15"
                                                    "GRM Category8","GRM Category16"
                                                        ])]["Total"].sum())
    st.markdown(df3rd[df3rd["Data Type"]=="Reconstruction started - Yes"]["Data"].item())
    st.markdown(landGRM[landGRM["landownership_type"]=="Total"]["count"].item())
    st.markdown(df[df["Data Type"]=="GRM CNIC Cases"]["Data"].item())

    summaryDatafile='Output files/FinalSummaryALL '+qlikData.name[:-5]+datetime.today().strftime('%d-%m-%Y')+'.xlsx'
    with pd.ExcelWriter(summaryDatafile) as writer:
        df.to_excel(writer, sheet_name='CNIC Summary',index=False)
        landALL.to_excel(writer, sheet_name='Land Titles Summary',index=False)
        landClear.to_excel(writer, sheet_name='Land Titles Cleared Summary',index=False)
        landGRM.to_excel(writer, sheet_name='Land Titles GRM Summary',index=False)
        df3rd.to_excel(writer, sheet_name='Reconstruction Summary',index=False)
        Damaged.to_excel(writer, sheet_name='Damaged Summary',index=False)
        Level1Summary.to_excel(writer, sheet_name='Level 1 Summary',index=False)
        hazardSum.to_excel(writer, sheet_name='Hazardeous with Duplicates',index=False)
        hazardSumNoDup.to_excel(writer, sheet_name='Hazardeous without Duplicates',index=False)
        Level2Summary.to_excel(writer, sheet_name='Level 2 Summary',index=False)
        
    invalidCNIC["AI Status"]="Invalid CNIC - GRM"
    holdCNIC_holdName_holdSFname["AI Status"]="CNIC - GRM"
    Holdland["AI Status"]="Land Title - GRM"
    hold3rd["AI Status"]="Reconstruction - GRM"
    DamagedGRM["AI Status"]="Damaged - GRM"
    holdHazardNoDup["AI Status"]="Hazardeous - GRM"
    clearedHazardNoDup["AI Status"]="All - Cleared"
    GRM_DATA=invalidCNIC._append([holdCNIC_holdName_holdSFname,Holdland,hold3rd,DamagedGRM])

    my_bar.progress(95, text="!!! Sabrookaaaa...  !!!")
    sleep(2)
    
    clearedHazardNoDup=clearedHazardNoDup[[
    "Rows Count","DA_UUID","UD_UUID","UUID Deviation","DA_CNIC","UD_CNIC","CNIC Deviation","DA Occupant Name","U_Occupant Name",
    "Occupant Name Deviation","DA Father Name","DA Spouse Name","U Father Spouse","Father Name Deviation","DA Type","UU Type","Type  Deviation",
    "DA_Damage Category","UD_Damage Category","Category Deviation","District","Tehsil","Deh","is_hazardous_location","landownership_type",
    "is_located_in_flood_plain","Has_Reconstruction_Started","have_bank_account","bank_account_number","bank_account_title","bank_name","bank_address",
    "financial_service_bank_nearest_branch","financial_service_bank_mother_maiden_name","financial_service_bank_name","Additional Remarks By STAT","Long",
    "Lat","IP Name",'village',"DistrictMAP","verify_DA_CNIC","AI Status","U_Occupant Name partial_ratio","DA Occupant Name Not Matching","DA Spouse Name Not Matching",
    "UD_Damage CategorySTD","DA_Damage CategorySTD","DA TypeSTD","UU TypeSTD","Combined","Category"]]

    holdHazardNoDup=holdHazardNoDup[[
    "Rows Count","DA_UUID","UD_UUID","UUID Deviation","DA_CNIC","UD_CNIC","CNIC Deviation","DA Occupant Name","U_Occupant Name",
    "Occupant Name Deviation","DA Father Name","DA Spouse Name","U Father Spouse","Father Name Deviation","DA Type","UU Type","Type  Deviation",
    "DA_Damage Category","UD_Damage Category","Category Deviation","District","Tehsil","Deh","is_hazardous_location","landownership_type",
    "is_located_in_flood_plain","Has_Reconstruction_Started","have_bank_account","bank_account_number","bank_account_title","bank_name","bank_address",
    "financial_service_bank_nearest_branch","financial_service_bank_mother_maiden_name","financial_service_bank_name","Additional Remarks By STAT","Long",
    "Lat","IP Name",'village',"DistrictMAP","verify_DA_CNIC","AI Status","U_Occupant Name partial_ratio","DA Occupant Name Not Matching","DA Spouse Name Not Matching",
    "UD_Damage CategorySTD","DA_Damage CategorySTD","DA TypeSTD","UU TypeSTD","Combined","Category"]]

    allData=GRM_DATA._append([holdHazardNoDup,clearedHazardNoDup])

    outputDatafile='Output files/FinalDataALL '+qlikData.name[:-5]+datetime.today().strftime('%d-%m-%Y')+'.xlsx'
    with pd.ExcelWriter(outputDatafile) as writer:
        Level2Summary.to_excel(writer, sheet_name='Data Summary',index=False)
        DamagedCleared[DamagedCleared.duplicated(subset="DA_CNIC",keep=False)].to_excel(writer, sheet_name='Data CNIC Duplicates',index=False)
        allData.to_excel(writer, sheet_name='ALL Data with AI status',index=False)

    # GRM_DATA.to_csv(outputDatafile+' GRM-DATA.csv',index=False)
    # holdHazardNoDup.to_csv(outputDatafile+'Data in Hazardeous NO-DUP.csv',index=False)
    # DamagedCleared[DamagedCleared.duplicated(subset="DA_CNIC",keep=False)].to_csv(outputDatafile+'Data CNIC Duplicates.csv',index=False)
    # clearedHazardNoDup.to_csv(outputDatafile+'All - Cleared.csv',index=False)

    my_bar.progress(100, text="!!! Muh kya dekh rahy download kro aur Enjoy  !!!")
    sleep(2)

    st.success("Processing Completed Click below to download output")
    with open(summaryDatafile, 'rb') as my_file:
        st.download_button(label = 'Download Summary data', data = my_file, file_name = summaryDatafile, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with open(outputDatafile, 'rb') as my_file:
        st.download_button(label = 'Download Full data', data = my_file, file_name = outputDatafile, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.warning('Upload input data to proceed ! ', icon="⚠️")