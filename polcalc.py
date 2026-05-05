import streamlit as st
import pandas as pd
import msoffcrypto
import io

# --- 1. એક્સેલમાંથી ડેટા લોડ કરવાનું ફંક્શન (પાસવર્ડ સાથે) ---
@st.cache_data
def load_data():
    file_path = "Polish Calc_Updated_05-05-2026 - Copy.xlsm" 
    file_password = "Laxmi#1"
    
    try:
        decrypted_workbook = io.BytesIO()
        
        # પાસવર્ડ નાખીને ફાઈલને મેમરીમાં ઓપન કરવી
        with open(file_path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=file_password)
            office_file.decrypt(decrypted_workbook)
            
        # હવે ડિક્રિપ્ટ થયેલી ફાઈલમાંથી ડેટા વાંચવો
        df_tables = pd.read_excel(decrypted_workbook, sheet_name="Tables")
        df_list = pd.read_excel(decrypted_workbook, sheet_name="List")
        
        return df_tables, df_list
        
    except Exception as e:
        st.error(f"ફાઈલ લોડ કરવામાં એરર આવી છે: {e}")
        return None, None

df_tables, df_list = load_data()
