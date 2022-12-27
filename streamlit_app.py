import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import os
import zipfile
import openpyxl
import xlsxwriter
import shutil

st.title('Fabric & Home Care Creative Bulk Upload from TTD Exports')

#Adgroup_end = st.selectbox(
#    'Which PMP type is named at the end of the adgroups for this upload?',
#    ('OTT~PMP', 'OEPMP'))

c_d_file = st.file_uploader("Choose the TTD Creative_details input csv file", type='csv')
b_u_file = st.file_uploader("Choose the TTD Bulk Uploads xlsx template file", type='xlsx')

lengths = st.selectbox(
    'Which Creative Lengths\Type would you like to include',
    (None,'30s', '15s', 'Both 15s and 30s', 'Display'))

def to_excel(df1, df2):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='Ad Groups', index=False)
    df2.to_excel(writer, sheet_name='Budget Flights', index=False)
    #workbook = writer.book
    #worksheet = writer.sheets['Ad Groups','Budget Flights']
    #format1 = workbook.add_format({'num_format': '0.00'}) 
    #worksheet.set_column('A:A', None, format1)  
    workbook = writer.book
    for i in range(len(custom_props_df['name'].tolist())):
        name = custom_props_df['name'].tolist()[i]
        val = custom_props_df[custom_props_df.columns.tolist()[3]].tolist()[i]
        workbook.set_custom_property(name, val)
    workbook.close()
    writer.save()
    processed_data = output.getvalue()
    return processed_data 

if ((c_d_file is not None) & (b_u_file is not None) & (lengths is not None)) :
    targetdir = './/unzipexcel'
    if os.path.exists(targetdir):
            shutil.rmtree(targetdir)
    os.mkdir(targetdir)
    with zipfile.ZipFile(b_u_file,"r") as zip_ref:
        zip_ref.extractall(targetdir)
    custom_props_df = pd.read_xml(targetdir+'//docProps/custom.xml')
    cr_dtls = pd.read_csv(c_d_file)
    crtv_info = cr_dtls[['CreativeName','CreativeId']].drop_duplicates()

    bu_adgroups = pd.read_excel(b_u_file, sheet_name = 'Ad Groups')
    bu_flights = pd.read_excel(b_u_file, sheet_name = 'Budget Flights')
    
    Adgroup_end = bu_adgroups.apply(lambda row :row['Ad Group Name'].rsplit('~', 2)[-2]+'~'+row['Ad Group Name'].rsplit('~', 2)[-1],
                                            axis=1).unique()[0]
    adgroup_sheet_columns = bu_adgroups.columns.tolist()

    del bu_adgroups['Creatives']

    crtv_info['AdGroupName']= crtv_info.apply(lambda row :'2223~'+
                                            row['CreativeName'].split('2223~')[1].split(Adgroup_end)[0]+Adgroup_end,
                                                axis=1)
    crtv_info['AdGrpName_Upper'] = crtv_info['AdGroupName'].str.upper()
    crtv_info['AdGrpName_Upper'] = crtv_info['AdGrpName_Upper'].str.replace('~','-')
    crtv_info['Export Creative Input'] = crtv_info['CreativeName']+'$id:'+crtv_info['CreativeId']+';'
    if lengths != 'Display':
        crtv_info['Creative Length']= crtv_info.apply(lambda row :'1 x 1.'+
                                                    row['CreativeName'].split('1 x 1.')[1].split('_FRQNOVIEW')[0],
                                                    axis=1)
        crtv_info_exp = crtv_info.groupby(['AdGroupName','AdGrpName_Upper','Creative Length']).agg({'Export Creative Input':' '.join}).reset_index()
    else:
        crtv_info_exp = crtv_info.groupby(['AdGroupName','AdGrpName_Upper']).agg({'Export Creative Input':' '.join}).reset_index()
    bu_adgroups['AdGrpName_Upper'] = bu_adgroups['Ad Group Name'].str.upper()
    bu_adgroups['AdGrpName_Upper'] = bu_adgroups['AdGrpName_Upper'].str.replace('~','-')
    if lengths != 'Display':
        adgroup_30s = bu_adgroups[['AdGrpName_Upper', 'Ad Group Name']].merge(crtv_info_exp[crtv_info_exp['Creative Length'] == '1 x 1.(:30)'][['AdGrpName_Upper','Export Creative Input']],
                      on= 'AdGrpName_Upper',
                      how='left')[['Ad Group Name','Export Creative Input']]
        adgroup_30s.columns = ['Ad Group Name','Export Creative Input :30s Only']
        adgroup_15s = bu_adgroups[['AdGrpName_Upper', 'Ad Group Name']].merge(crtv_info_exp[crtv_info_exp['Creative Length'] == '1 x 1.(:15)'][['AdGrpName_Upper','Export Creative Input']],
                      on= 'AdGrpName_Upper', 
                      how='left')[['Ad Group Name','Export Creative Input']]
        adgroup_15s.columns = ['Ad Group Name','Export Creative Input :15s Only']
        adgroup_output = adgroup_30s.merge(adgroup_15s, on='Ad Group Name',how='outer')
        adgroup_output.fillna('',inplace=True)
        adgroup_output['Export Creative Input :15s & :30s'] = adgroup_output['Export Creative Input :30s Only']+adgroup_output['Export Creative Input :15s Only']
        if lengths == '15s':
            creative_cols = ['Ad Group Name','Export Creative Input :15s Only']
        if lengths == '30s':
            creative_cols = ['Ad Group Name','Export Creative Input :30s Only']
        if lengths == 'Both 15s and 30s':
            creative_cols = ['Ad Group Name','Export Creative Input :15s & :30s']
    if lengths == 'Display':
        adgroup_output = bu_adgroups[['AdGrpName_Upper', 'Ad Group Name']].merge(crtv_info_exp[['AdGrpName_Upper','Export Creative Input']],
                      on= 'AdGrpName_Upper', 
                      how='left')[['Ad Group Name','Export Creative Input']]
        creative_cols = ['Ad Group Name','Export Creative Input']    
    final_adgroup_output = adgroup_output[creative_cols].copy()
    final_adgroup_output.columns = ['Ad Group Name','Creatives']

    bu_adgroups_final = bu_adgroups.merge(final_adgroup_output, on='Ad Group Name')[adgroup_sheet_columns]

    file_out_name = bu_adgroups['Campaign [Read Only]'].values[0].split('_')[-1]+'_'+lengths+'_bulk_upload.xlsx'

    file_out_name = file_out_name.replace(':','-')

    df_xlsx = to_excel(bu_adgroups_final, bu_flights)

    st.download_button(label='📥 Download Bulk Upload Result',
                                data=df_xlsx ,
                                file_name= file_out_name)
