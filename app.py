# Import Libraries
import pandas as pd
import numpy as np
import xml.etree.cElementTree as et
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

# Configure Page Title and Icon
st.set_page_config(page_title='Tableau Documentor',page_icon=':smile:')

# Removing Streamlit's Hamburger and Footer
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            a {text-decoration: none;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Declaring custom functions
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.header('Tableau Documentation Made Easy !!')
st.info('It takes atleast 2 minute to open a Tableau File and copy one calculation from Tableau to Excel.  \nIf you have even 10 calulations, that will take 10*2=20 minutes minimum.')
st.success('You can do that in seconds using this website!!')
st.markdown("---")
st.subheader('Upload your TWB file')
uploaded_file=st.file_uploader('',type=['twb'],)

if uploaded_file is not None:
    tree=et.parse(uploaded_file)
    root=tree.getroot()

    # Calculated Fields
    name = []
    formula = []

    # Getting Names of calculated fields
    for x in root.findall('datasources'):
        for y in x.findall('datasource'):
            for z in y.findall('column'):
                try:
                    name.append(z.attrib['caption'])
                except:
                    print('not allowed')
                    #name.append(z.attrib['caption'])


    # Getting formulas
    for x in root.findall('datasources'):
        for y in x.findall('datasource'):
            for z in y.findall('column'):
                for al in z.findall('calculation'):
                    formula.append(al.attrib['formula'])


    print(len(formula),len(name))

    # Creating Dataframe
    df=pd.DataFrame(list(zip(name,formula)),columns=['Calculated Field','Formulae'])

    # Showing Dataframe
    st.write(df)

    # Showing Multiple Download Options
    st.subheader("Select the format of Output file : ")
    op=st.radio('Options', ["CSV File","EXCEL File"])

    # Download Functionality
    if op=="CSV":
        st.download_button("Download",df.to_csv(),file_name="Documentation-csv-output",mime="text/csv")
    else:
        st.download_button("Download",to_excel(df),file_name="Documentation-excel-output")
st.markdown('---')
st.markdown('Made with :heart: by [Sahil Choudhary](https://www.sahilchoudhary.ml/)')
