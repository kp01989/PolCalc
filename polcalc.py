import streamlit as st
import pandas as pd
import msoffcrypto
import io

# --- 1. પેજનું સેટિંગ ---
st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.title("💎 Diamond Polish Calculator")
st.markdown("તમારા ડાયમંડની ડિટેલ્સ નીચે સિલેક્ટ કરો એટલે ઓટોમેટિક ગણતરી થઈ જશે.")

# --- 2. એક્સેલમાંથી ડેટા લોડ કરવાનું ફંક્શન ---
@st.cache_data
def load_data():
    file_path = "Polish Calc_Updated_05-05-2026 - Copy.xlsm" 
    file_password = "Laxmi#1"
    
    try:
        with open(file_path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            
            # પાસવર્ડ ચેકિંગ
            if office_file.is_encrypted():
                decrypted_workbook = io.BytesIO()
                office_file.load_key(password=file_password)
                office_file.decrypt(decrypted_workbook)
                df_tables = pd.read_excel(decrypted_workbook, sheet_name="Tables")
                df_list = pd.read_excel(decrypted_workbook, sheet_name="List")
            else:
                df_tables = pd.read_excel(file_path, sheet_name="Tables")
                df_list = pd.read_excel(file_path, sheet_name="List")
                
        return df_tables, df_list
        
    except Exception as e:
        st.error(f"ફાઈલ લોડ કરવામાં એરર આવી છે: {e}")
        return None, None

df_tables, df_list = load_data()

# --- 3. UI અને ડ્રોપડાઉન (જો ડેટા મળે તો જ બતાવશે) ---
if df_tables is not None:
    
    st.subheader("હીરાની વિગતો સિલેક્ટ કરો")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        shape = st.selectbox("SHAPE (શેપ)", ["ROUND", "PRINCESS", "OVAL", "MARQUISE", "PEAR", "EMERALD"])
        color = st.selectbox("Color (કલર)", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"])
        clarity = st.selectbox("Clarity (ક્લેરિટી)", ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"])
        
    with col2:
        cut = st.selectbox("Cut Group (કટ)", ["3EX", "EX", "VG", "GD", "FAIR"])
        fluorescence = st.selectbox("Fluorescence (ફ્લોરોસન્સ)", ["NON", "FNT", "MED", "STG", "VSTG"])
        size_range = st.selectbox("Size (સાઈઝ)", ["0.18 - 0.22", "0.23 - 0.29", "0.30 - 0.34", "0.35 - 0.39", "0.40 - 0.49"])
        
    with col3:
        polish_weight = st.number_input("Polish Weight (વજન)", min_value=0.01, value=0.30, step=0.01)
        rap_price = st.number_input("Rap Price ($)", min_value=0.0, value=1700.0, step=100.0)

    st.divider()

    # --- 4. કેલ્ક્યુલેશન (ગણતરી) ---
    if st.button("Calculate Discount & Amount", type="primary"):
        try:
            # સાઈઝમાંથી પહેલો આંકડો લેવા (જેમ કે 0.30)
            base_size = float(size_range.split(" ")[0]) 
            
            # Tables માંથી ડેટા મેચ કરવો
            condition = (df_tables['Size'] == base_size) & (df_tables['Color'] == color)
            match_data = df_tables[condition]
            
            discount_percent = 0.0
            
            if not match_data.empty:
                discount_percent = float(match_data[clarity].values[0])
            else:
                st.warning("આ કોમ્બિનેશન માટે Tables શીટમાં કોઈ ડિસ્કાઉન્ટ મળ્યું નથી! (ડેમો માટે -30% ગણીએ છીએ)")
                discount_percent = -30.0 

            # રેટ અને એમાઉન્ટની ગણતરી
            rate_per_cts = rap_price * (1 + (discount_percent / 100))
            pol_amt = rate_per_cts * polish_weight
            
            # --- 5. પરિણામ (Output) ---
            st.subheader("📊 ગણતરીનું પરિણામ")
            
            res_col1, res_col2, res_col3 = st.columns(3)
            res_col1.metric("Discount / Prem. (%)", f"{discount_percent}%")
            res_col2.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
            res_col3.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
            
            st.success("તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
            
        except Exception as e:
            st.error(f"ગણતરીમાં કોઈ ભૂલ આવી છે. એરર: {e}")
