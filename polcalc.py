import streamlit as st
import pandas as pd
import msoffcrypto
import io

# --- 1. પેજનું સેટિંગ ---
st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.title("💎 Diamond Polish Calculator")
st.markdown("તમારા ડાયમંડની ડિટેલ્સ નીચે સિલેક્ટ કરો એટલે ઓટોમેટિક ગણતરી થઈ જશે.")

# ફંક્શનનું નામ બદલ્યું છે જેથી જૂનો Cache ક્લિયર થઈ જાય
@st.cache_data
def load_diamond_data():
    file_path = "Polish Calc_Updated_05-05-2026 - Copy.xlsm" 
    file_password = "Laxmi#1"
    
    try:
        with open(file_path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            if office_file.is_encrypted():
                decrypted_workbook = io.BytesIO()
                office_file.load_key(password=file_password)
                office_file.decrypt(decrypted_workbook)
                df = pd.read_excel(decrypted_workbook, sheet_name="Tables", header=None)
            else:
                df = pd.read_excel(file_path, sheet_name="Tables", header=None)
        return df
    except Exception as e:
        st.error(f"ફાઈલ ઓપન કરવામાં એરર: {e}")
        return None

df_tables = load_diamond_data()

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

    if st.button("Calculate Discount & Amount", type="primary"):
        try:
            base_size = float(size_range.split(" ")[0]) 
            cut_fluo = f"{cut}-{fluorescence}" 
            
            # બુલેટપ્રૂફ મેચિંગ લોજિક
            row_0 = df_tables.iloc[0].fillna('').astype(str).str.strip().str.upper()
            row_1 = df_tables.iloc[1].fillna('').astype(str).str.strip().str.upper()
            
            if cut_fluo not in row_0.values:
                st.error(f"Excel માં '{cut_fluo}' નામનું કોઈ હેડિંગ મળ્યું નહિ!")
            else:
                # 1. સાઈઝ અને કલર ની કોલમ જાતે જ શોધી લેશે
                size_col_idx = None
                color_col_idx = None
                
                for idx, val in row_1.items():
                    if val == 'SIZE':
                        size_col_idx = idx
                    elif val == 'COLOR':
                        color_col_idx = idx
                        
                # જો એક્સેલમાં નામ અલગ હોય તો ડીફોલ્ટ સેટિંગ
                if size_col_idx is None: size_col_idx = 1
                if color_col_idx is None: color_col_idx = 2

                # 2. કટ અને ફ્લોરોસન્સ ની કોલમ શોધશે
                cut_fluo_idx = row_0[row_0 == cut_fluo].index[0]
                
                # 3. ક્લેરિટી ગોતીને સાચી કોલમ પકડશે
                clarity_list = ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"]
                if clarity in clarity_list:
                    clarity_offset = clarity_list.index(clarity)
                    final_col_idx = cut_fluo_idx + clarity_offset
                    
                    # 4. સાઈઝ અને કલરનો ડેટા મેચ કરશે
                    data_rows = df_tables.iloc[2:].copy()
                    
                    size_match = pd.to_numeric(data_rows[size_col_idx], errors='coerce') == base_size
                    color_match = data_rows[color_col_idx].astype(str).str.strip().str.upper() == color
                    
                    match_data = data_rows[size_match & color_match]
                    
                    if not match_data.empty:
                        discount_percent = float(match_data.iloc[0, final_col_idx])
                        
                        rate_per_cts = rap_price * (1 + (discount_percent / 100))
                        pol_amt = rate_per_cts * polish_weight
                        
                        st.subheader("📊 ગણતરીનું પરિણામ")
                        res_col1, res_col2, res_col3 = st.columns(3)
                        res_col1.metric("Discount / Prem. (%)", f"{discount_percent}%")
                        res_col2.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
                        res_col3.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
                        
                        st.success("🎉 તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
                    else:
                        st.warning(f"આ સાઈઝ ({base_size}) અને કલર ({color}) માટે એક્સેલમાં ડેટા મળ્યો નથી.")
                else:
                    st.error("ક્લેરિટીનું નામ મેચ થતું નથી.")
        except Exception as e:
            st.error(f"ગણતરીમાં અણધારી ભૂલ આવી: {e}")
