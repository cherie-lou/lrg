import numpy as np
import pandas as pd
import streamlit as st
import openpyxl
import xlsxwriter
import io

st.markdown("[Download RFP TemplateV1.xlsx](https://logrgadmin.sharepoint.com/:x:/t/engineering/Ef67ReTKSZ1JjeIyYrktfpoBmiooobNeLLUqZH8d72cVwg?e=rbIV4u)", unsafe_allow_html=True)

st.title("RFPToolV1")
st.header('File Input')
input_file = st.file_uploader("Upload an Excel File", type=["xls", "xlsx"])

def unique(list1):
    ans = pd.Series(list1).drop_duplicates().to_list()
    return ans

def CarrierFilter(df,selectcarrier):
    result = df[df['Carrier'].isin(selectcarrier)]
    return result

def TempRankFilter(df):
    temp = df.copy()
    temp["TempRank"] = temp.groupby('ShipmentID')['Linehaul'].rank(method = 'first')
    result = temp[temp['TempRank']==1]
    return result

def LaneCalculation(df_lane,merged_df,carrierx,x):
    # carrierx_df_t = merged_df[merged_df['Carrier'].isin(carrierx)]
    
    carrierx_df = TempRankFilter(CarrierFilter(merged_df,carrierx))
    cx_min_df = carrierx_df[carrierx_df['MinFilter']=='Min']
    cx_nonmin_df = carrierx_df[carrierx_df['MinFilter']=='Non-Min'].groupby('Lane').agg({'Linehaul': 'sum', 'Czalite0%': 'sum'})
    df_l = df_lane.copy()
    df_l['Shipment Count'+x]=carrierx_df.groupby("Lane").agg({"ShipmentID":"count"})
    # df_l['LLC Carrier'+x]=carrierx_df.groupby("Lane").agg({"Carrier":"first"})
    df_l['AvgMin'+x]=cx_min_df.groupby('Lane').agg({'Linehaul':'mean'})
    df_l['Disc Non Min'+x]= 1-cx_nonmin_df['Linehaul'] / cx_nonmin_df['Czalite0%']
    df_l['Total Linehaul'+x]=carrierx_df.groupby('Lane').agg({'Linehaul': 'sum'})
    return df_l

def RateReviewPrimary(primary_carrier,pivot_lane):
    col_list = ['Lane','ShipmentID']
    for c in primary_carrier:
        col_list.append(c+'_Disc')
        col_list.append(c+'_Min')
    result = pivot_lane[col_list]
    return result

def RateReviewOther(primary_carrier, pivot_lane, carrier_list):
    col_list = ['Lane']
    for i in carrier_list:
        if i not in primary_carrier:
            col_list.append(i + '_Disc')
            col_list.append(i + '_Min')
    df_other = pivot_lane[col_list].copy()
    result = pivot_lane[['Lane']].copy() 
    disc_col = [col for col in df_other.columns if col.endswith('_Disc')]
    result['LLC-Disc'] = df_other.loc[:, disc_col].max(axis=1)
    min_col = [col for col in df_other.columns if col.endswith('_Min')]
    result['LLC-Disc-Carrier'] = df_other.loc[:,disc_col].idxmax(axis=1)
    result['LLC-Min'] = df_other.loc[:, min_col].min(axis=1)
    result['LLC-Min-Carrier'] = df_other.loc[:,min_col].idxmin(axis=1)
    result.reset_index(drop=True, inplace=True)
    return result

def Ratereview(df_primary,df_other,primary_carrier):
    df_ratereview = df_primary.merge(df_other,on='Lane')
    for p in primary_carrier:
        disc_title = p+'-LLC Disc Diff'
        min_title = p+'-LLC Min Diff'
        df_ratereview[disc_title]=df_ratereview[p+'_Disc']-df_ratereview['LLC-Disc']
        df_ratereview[min_title]=df_ratereview[p+'_Min']-df_ratereview['LLC-Min']
        # df_ratereview[disc_title] = df_ratereview.style.applymap(lambda x:style_disc_cell(x),subset=[disc_title])
    return df_ratereview

if input_file is not None:
    df_loads = pd.read_excel(input_file,sheet_name='RFP-Loads')
    df_base = pd.read_excel(input_file ,sheet_name='BaseRate')
    df_dis = pd.read_excel(input_file ,sheet_name='DiscountMinDatabase')
    # df_exc = pd.read_excel(input_file ,sheet_name='ExcludedCarrier')
    # df_loads = st.dataframe(df_loads)
    # df_base = st.dataframe(df_base)
    # df_dis = st.dataframe(df_dis)
    # df_exc = st.dataframe(df_exc)

    carrier_list = unique(df_dis['Carrier'])
    options = st.multiselect("Carrier excluded",carrier_list)
    exc_list = options
    
    # exc_list = df_exc['ExcludedCarrier']
    
    df_dis = df_dis[~df_dis['Carrier'].isin(exc_list)]
    df_dis['Lane']=df_dis.apply(lambda row: f"{row['StateOrig']} - {row['StateDest']}",axis = 1)
    df_dis.drop_duplicates(subset=['Lane', 'Carrier'], keep='first', inplace=True)
    df_loads['Lane']=df_loads.apply(lambda row: f"{row['StateOrig']} - {row['StateDest']}",axis = 1)

    merged_df_shipment = df_loads.merge(df_dis, on=['StateOrig','StateDest'])
    merged_df = merged_df_shipment.merge(df_base,on="ShipmentID")
    merged_df["Linehaul"] = np.maximum(merged_df['Czalite0%'] * (1-merged_df['Disc']), merged_df['Min'])
    minfilter = (merged_df['Linehaul']==merged_df['Min'])
    merged_df['MinFilter'] = np.where(minfilter,'Min','Non-Min')
    merged_df['Lane']=merged_df.apply(lambda row: f"{row['StateOrig']} - {row['StateDest']}",axis = 1)

    df1_lane_counts = df_loads.groupby(['StateOrig', 'StateDest']).size().reset_index(name='Count')
    df2_lane_counts = df_dis.groupby(['StateOrig', 'StateDest']).size().reset_index(name='Count')
    df_missing = pd.merge(df1_lane_counts, df2_lane_counts, on=['StateOrig', 'StateDest'], how='left',suffixes=('_df1', '_df2'))
    # df_missing['Count'] = df_missing['Count_df1'].fillna(0).astype(int) - df_missing['Count_df2'].fillna(0).astype(int)
    df_missing_filtered = df_missing[df_missing['Count_df2'].isna()].drop(['Count_df2'],axis=1)
    df_missing_filtered.rename(columns = {'Count_df1':"Count of Lanes"})

    merged_df["Rank"] = merged_df.groupby('ShipmentID')['Linehaul'].rank(method = 'first')

    rank1_merged_df = merged_df[merged_df['Rank']==1]
    df_bycarrier = rank1_merged_df.groupby('Carrier').agg({'ShipmentID':'count','Linehaul':'sum','Czalite0%':'sum'})
    df_bycarrier['AvgDisc']=1-df_bycarrier['Linehaul']/df_bycarrier['Czalite0%']
    # df_bycarrier.rename(columns={'ShipmentID':'Awarded Loads','Linehaul':'Awarded Linehaul','Czalite0%':'BaseRateTotal'})
    min_df = rank1_merged_df[rank1_merged_df['MinFilter']=='Min']
    df_bycarrier[['AvgMin','MinCount']]=min_df.groupby('Carrier').agg({'Linehaul':'mean','ShipmentID':'count'})
    nonmin_df = rank1_merged_df[rank1_merged_df['MinFilter']=='Non-Min']
    agg_nonmin=nonmin_df.groupby('Carrier').agg({'Linehaul': 'sum', 'Czalite0%': 'sum'})
    df_bycarrier['AvgDiscNonMin'] = 1-agg_nonmin['Linehaul'] / agg_nonmin['Czalite0%']

    df_bycarrier_od = rank1_merged_df.groupby(['Lane','Carrier']).agg({'ShipmentID':'count','Linehaul':'sum','Czalite0%':'sum'})
    df_bycarrier_od.rename(columns={'ShipmentID':'Awarded Loads','Linehaul':'Awarded Linehaul','Czalite0%':'BaseRateTotal'})
    df_bycarrier_od['AvgDisc']=1-df_bycarrier_od['Linehaul']/df_bycarrier_od['Czalite0%']
    df_bycarrier_od[['AvgMin','MinCount']]=min_df.groupby(['Lane','Carrier']).agg({'Linehaul':'mean','ShipmentID':'count'})
    agg_nonmin_od=nonmin_df.groupby(['Lane','Carrier']).agg({'Linehaul': 'sum', 'Czalite0%': 'sum'})
    df_bycarrier_od['AvgDiscNonMin'] = 1-agg_nonmin_od['Linehaul'] / agg_nonmin_od['Czalite0%']
    df_bycarrier_od=df_bycarrier_od.reset_index()
    

    st.subheader('Scenario Comparison')
    #Select Carriers for Scenario side by side analysis
    carrier_ava= list(set(carrier_list)-set(exc_list))
    cola,colb = st.columns(2)
    with cola:
        s1 = st.text_input("Name of 1st scenario")
    with colb:
        s2 = st.text_input("Name of 2nd scenario")
    col1,col2 = st.columns(2)
    if s1 and s2:
        with col1: 
            selected_carrier1 = list(st.multiselect("Carrier for "+s1,carrier_ava))
        with col2:
            selected_carrier2 = list(st.multiselect("Carrier for "+s2,carrier_ava))
        

        df_lane = merged_df.groupby("Lane").agg({"ShipmentID":"count"})
        df_lane_1 = LaneCalculation(df_lane,merged_df,selected_carrier1,'_'+s1)
        df_lane_2 = LaneCalculation(df_lane,merged_df,selected_carrier2,'_'+s2)
        # # df_lane=df_lane.merge(df_lane_1,on = 'ShipmentID')
        df_compare = df_lane_1.merge(df_lane_2,on="Lane")
        df_compare['Min Difference'] =  df_compare['AvgMin_'+s1]-df_compare['AvgMin_'+s2]
        df_compare['Discount Difference']=df_compare['Disc Non Min_'+s1]-df_compare['Disc Non Min_'+s2]
        df_compare['Awarded Total Linehaul Difference']=df_compare['Total Linehaul_'+s1]-df_compare['Total Linehaul_'+s2]
        df_compare = df_compare.drop(['ShipmentID_x','ShipmentID_y','Shipment Count_'+s2],axis =1)

    df_dis['Lane']=df_dis.apply(lambda row: f"{row['StateOrig']} - {row['StateDest']}",axis = 1)
    df_loads['Lane']=df_loads.apply(lambda row: f"{row['StateOrig']} - {row['StateDest']}",axis = 1)
    count_lane = df_loads.groupby('Lane').agg({"ShipmentID":"count"})
    pivot_lane = df_dis.pivot(index='Lane',columns='Carrier',values = ['Disc','Min'])
    pivot_lane.columns = [f'{col[1]}_{col[0]}' for col in pivot_lane.columns]
    pivot_lane = pivot_lane.merge(count_lane,on='Lane').reset_index()

        
    if st.button("Scenario Comparison"):
        st.dataframe(df_compare)
            
        custome_scenario_name = st.text_input("Enter the custom filename (e.g., MyCustomFile.xlsx):"):
        output_scenario = io.BytesIO()
        with pd.ExcelWriter(output_scenario, engine='xlsxwriter') as writer:
            df_compare.to_excel(writer)
        data_scenario = output_scenario.getvalue()
        st.download_buttone(label = "Download Scenario Comparison",data = data_scenario,file_name = custome_scenario_name,key = 'download')

            
    st.subheader("Rate Review")
    #select primary carrier for Rate Review Analysis
    primary_carrier = st.multiselect("Primary Carrier",carrier_ava)

    df_primary = RateReviewPrimary(primary_carrier,pivot_lane)
    df_other = RateReviewOther(primary_carrier,pivot_lane,carrier_list)
    df_ratereview =Ratereview(df_primary,df_other,primary_carrier)

    if st.button("Rate Review"):
        st.dataframe(df_ratereview)
            
        custome_rr_name = st.text_input("Enter the custom filename (e.g., MyCustomFile.xlsx):"):
        output_rr = io.BytesIO()
        with pd.ExcelWriter(output_rr, engine='xlsxwriter') as writer:
            df_ratereview.to_excel(writer)
        data_rr = output_rr.getvalue()
        st.download_buttone(label = "Download Rate Review",data = data_rr,file_name = custome_rr_name,key = 'download') 

    

    st.header('Output Setting')
    custom_filename = st.text_input("Enter the custom filename (e.g., MyCustomFile.xlsx):")
    st.header('Download output')

    output_final = io.BytesIO()  # Create a bytes buffer to store the Excel file
    with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
        df_bycarrier.to_excel(writer, sheet_name='Selected Summary Table')
        df_missing_filtered.to_excel(writer, sheet_name='Missing Lanes Table', index=False)
        df_bycarrier_od.to_excel(writer, sheet_name='Orig Dest Carrier Summary Table')
        merged_df.to_excel(writer, sheet_name='All Combination Table', index=False)

    # Prepare the Excel file for download
    excel_data_final = output_final.getvalue()

# Offer the download of the Excel file
    st.download_button(label="Download Excel File", data=excel_data_final, file_name=custom_filename, key='download')



        

        
