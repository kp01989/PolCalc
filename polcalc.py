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
            
            if office_file.is_encrypted():
                decrypted_workbook = io.BytesIO()
                office_file.load_key(password=file_password)
                office_file.decrypt(decrypted_workbook)
                # અહી મેં 'header=None' કર્યું છે જેથી પાયથોન મર્જ કરેલા સેલમાં અટવાય નહિ
                df_tables = pd.read_excel(decrypted_workbook, sheet_name="Tables", header=None)
            else:
                df_tables = pd.read_excel(file_path, sheet_name="Tables", header=None)
                
        return df_tables
        
    except Exception as e:
        st.error(f"ફાઈલ લોડ કરવામાં એરર આવી છે: {e}")
        return None

df_tables = load_data()

# --- 3. UI અને ડ્રોપડાઉન ---
if df_tables is not None:
    st.subheader("હીરાની વિગતો સિલેક્ટ કરો")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        shape = st.selectbox("SHAPE (શેપ)", ["ROUND", "PRINCESS", "OVAL"])
        color = st.selectbox("Color (કલર)", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"])
        clarity = st.selectbox("Clarity (ક્લેરિટી)", ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"])
        
    with col2:
        cut = st.selectbox("Cut Group (કટ)", ["3EX", "VG", "EX", "GD", "FAIR"])
        fluorescence = st.selectbox("Fluorescence", ["NON", "FNT", "MED", "STG", "VSTG"])
        size_range = st.selectbox("Size (સાઈઝ)", ["0.18 - 0.22", "0.23 - 0.29", "0.30 - 0.34", "0.35 - 0.39", "0.40 - 0.49"])
        
    with col3:
        polish_weight = st.number_input("Polish Weight", min_value=0.01, value=0.30, step=0.01)
        rap_price = st.number_input("Rap Price ($)", min_value=0.0, value=1700.0, step=100.0)

    st.divider()

    # --- 4. એક્સેલ જેવું એક્ઝેક્ટ મેચિંગ લોજિક ---
    if st.button("Calculate Discount & Amount", type="primary"):
        try:
            base_size = float(size_range.split(" ")[0]) 
            
            # કટ અને ફ્લોરોસન્સ જોડીને નામ બનાવ્યું (દા.ત. "3EX-NON")
            cut_fluo = f"{cut}-{fluorescence}" 
            
            # પહેલી જ લાઈનમાંથી (Excel ની Row 1 માંથી) આ નામ ગોતશે
            row_0 = df_tables.iloc[0].fillna('').astype(str).str.strip()
            
            if cut_fluo not in row_0.values:
                st.error(f"એક્સેલની પહેલી લાઈનમાં '{cut_fluo}' નામનું કોઈ કોમ્બિનેશન મળ્યું નહિ!")
            else:
                # 3EX-NON કઈ કોલમમાં છે તેનો નંબર કાઢ્યો
                cut_fluo_idx = row_0[row_0 == cut_fluo].index[0]
                
                # ક્લેરિટીના લિસ્ટમાંથી ગણીને સાચી કોલમ પકડી લેશે (જેમ તમે ફોર્મ્યુલામાં + 1, 2, 3 કરતા હતા)
                clarity_list = ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
                if clarity in clarity_list:
                    clarity_offset = clarity_list.index(clarity)
                    final_col_idx = cut_fluo_idx + clarity_offset
                else:
                    st.error("ક્લેરિટીનું નામ મેચ થતું નથી.")
                    final_col_idx = None
                
                if final_col_idx is not None:
                    # હવે હેડિંગ વાળી લાઈનો છોડીને ડેટામાં સાઈઝ અને કલર મેચ કરશે
                    # (એક્સેલ પ્રમાણે: કોલમ 1 માં સાઈઝ છે અને કોલમ 2 માં કલર છે)
                    data_rows = df_tables.iloc[2:].copy()
                    
                    # સાઈઝ કોલમ નંબર 1 છે અને કલર કોલમ નંબર 2 છે
                    condition = (pd.to_numeric(data_rows[1], errors='coerce') == base_size) & (data_rows[2] == color)
                    match_data = data_rows[condition]
                    
                    if not match_data.empty:
                        # સીધો જ સાચો ડિસ્કાઉન્ટનો આંકડો ખેંચી લેશે
                        discount_percent = float(match_data.iloc[0, final_col_idx])
                        
                        rate_per_cts = rap_price * (1 + (discount_percent / 100))
                        pol_amt = rate_per_cts * polish_weight
                        
                        st.subheader("📊 ગણતરીનું પરિણામ")
                        res_col1, res_col2, res_col3 = st.columns(3)
                        res_col1.metric("Discount / Prem. (%)", f"{discount_percent}%")
                        res_col2.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
                        res_col3.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
                        
                        st.success("તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
                    else:
                        st.warning("આ સાઈઝ અને કલર માટે ડિસ્કાઉન્ટ ટેબલમાં ડેટા મળ્યો નથી.")
                        
        except Exception as e:
            st.error(f"ગણતરીમાં કોઈ અણધારી ભૂલ આવી છે: {e}")
