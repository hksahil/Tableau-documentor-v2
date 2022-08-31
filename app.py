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
st.info('It takes atleast 2 minutes to open a Tableau File and copy one calculation from Tableau to Excel.  \nIf you have even 10 calulations, that will take 10*2=20 minutes minimum.')
st.success('You can do that in seconds with just 1 click using this website!!')
st.markdown("---")
st.subheader('Upload your TWB file')
uploaded_file=st.file_uploader('',type=['twb'],)

if uploaded_file is not None:
    tree=et.parse(uploaded_file)
    root=tree.getroot()

    # create a dictionary of name and tableau generated name

    calcDict = {}

    for item in root.findall('.//column[@caption]'):
        if item.find(".//calculation") is None:
            continue
        else:
            calcDict[item.attrib['name']] = '[' + item.attrib['caption'] + ']'

    # list of calc's name, tableau generated name, and calculation/formula
    calcList = []

    for item in root.findall('.//column[@caption]'):
        if item.find(".//calculation") is None:
            continue
        else:
            if item.find(".//calculation[@formula]") is None:
                continue
            else:
                calc_caption = '[' + item.attrib['caption'] + ']'
                calc_name = item.attrib['name']
                calc_raw_formula = item.find(".//calculation").attrib['formula']
                calc_comment = ''
                calc_formula = ''
                for line in calc_raw_formula.split('\r\n'):
                    if line.startswith('//'):
                        calc_comment = calc_comment + line + ' '
                    else:
                        calc_formula = calc_formula + line + ' '
                for name, caption in calcDict.items():
                    calc_formula = calc_formula.replace(name, caption)

                calc_row = (calc_caption, calc_name, calc_formula, calc_comment)
                calcList.append(list(calc_row))

    # convert the list of calcs into a data frame
    data = calcList

    data = pd.DataFrame(data, columns=['Name', 'Remote Name', 'Formula', 'Comment'])

    # remove duplicate rows from data frame
    data = data.drop_duplicates(subset=None, keep='first', inplace=False)

    df=data[['Name','Formula']]
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
