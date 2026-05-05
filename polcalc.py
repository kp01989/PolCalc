import streamlit as st
import pandas as pd
import msoffcrypto
import io

# --- 1. પેજનું સેટિંગ ---
st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.title("💎 Diamond Polish Calculator")
st.markdown("તમારા ડાયમંડની ડિટેલ્સ નીચે સિલેક્ટ કરો એટલે ઓટોમેટિક ગણતરી થઈ જશે.")

# --- 2. એક્સેલમાંથી બધી શીટ લોડ કરવાનું ફંક્શન ---
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
                df_tables = pd.read_excel(decrypted_workbook, sheet_name="Tables", header=None)
                df_list = pd.read_excel(decrypted_workbook, sheet_name="List", header=None) 
            else:
                df_tables = pd.read_excel(file_path, sheet_name="Tables", header=None)
                df_list = pd.read_excel(file_path, sheet_name="List", header=None)
        return df_tables, df_list
    except Exception as e:
        st.error(f"ફાઈલ ઓપન કરવામાં એરર: {e}")
        return None, None

data_loaded = load_diamond_data()

if data_loaded[0] is not None:
    df_tables, df_list = data_loaded
    
    st.subheader("હીરાની વિગતો સિલેક્ટ કરો")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        shape = st.selectbox("SHAPE (શેપ)", ["ROUND", "PRINCESS", "OVAL", "MARQUISE", "PEAR", "EMERALD"])
        color = st.selectbox("Color (કલર)", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"])
        clarity = st.selectbox("Clarity (ક્લેરિટી)", ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"])
        
    with col2:
        cut = st.selectbox("Cut Group (કટ)", ["3EX", "VG", "EX", "GD", "FAIR"])
        fluorescence = st.selectbox("Fluorescence", ["NON", "FNT", "MED", "STG", "VSTG"])
        polish_weight = st.number_input("Polish Weight (વજન)", min_value=0.01, value=0.30, step=0.01)

    # ==========================================
    # POINT 1: VLOOKUP (Approximate Match) for SIZE
    # ==========================================
    calc_size = ""
    try:
        size_df = df_list.iloc[1:24, 22:24].copy() 
        size_df.columns = ['Weight', 'SizeLabel']
        size_df['Weight'] = pd.to_numeric(size_df['Weight'], errors='coerce')
        size_df = size_df.dropna(subset=['Weight']).sort_values('Weight')
        
        for idx, row in size_df.iterrows():
            if polish_weight >= row['Weight']:
                calc_size = str(row['SizeLabel']).strip()
    except Exception as e:
        calc_size = "Error"

    # ==========================================
    # POINT 2: VLOOKUP (Exact Match) for RAP PRICE 
    # ==========================================
    calc_rap = 0.0
    joined_str = ""
    if calc_size and calc_size != "Error":
        
        # આપણું સર્ચ કરવાનું નામ (બધી જ સ્પેસ હટાવીને કેપિટલ કર્યું)
        joined_str = f"{shape}{calc_size}{clarity}{color}".replace(" ", "").upper()
        
        try:
            # 15મી કોલમ (Index 14) નો ડેટા લીધો અને એમાંથી પણ બધી જ સ્પેસ કાઢી નાખી
            lookup_range = df_list.iloc[:, 14].astype(str).str.replace(" ", "").str.upper()
            
            # પર્ફેક્ટ VLOOKUP માર્યું
            if joined_str in lookup_range.values:
                match_idx = lookup_range[lookup_range == joined_str].index[0]
                
                # 16મી કોલમ (Index 15) માંથી જવાબ લીધો
                val = df_list.iloc[match_idx, 15]
                calc_rap = float(pd.to_numeric(val, errors='coerce'))
                if pd.isna(calc_rap):
                    calc_rap = 0.0
        except Exception as e:
            calc_rap = 0.0

    with col3:
        st.text_input("Size (ઓટોમેટિક સાઈઝ)", value=calc_size, disabled=True)
        st.text_input("Rap Price ($) (ઓટોમેટિક રૅપ)", value=f"{calc_rap:,.2f}" if calc_rap else "0.00", disabled=True)
        
    st.divider()

    # --- ફાઇનલ ગણતરી ---
    if st.button("Calculate Discount & Amount", type="primary"):
        if not calc_size or calc_size == "Error":
            st.error("સાઈઝ ઓટોમેટિક કેલ્ક્યુલેટ થઈ શકી નથી. મહેરબાની કરીને List શીટમાં W2:X23 ચેક કરો.")
        elif calc_rap == 0.0:
            st.error(f"VLOOKUP ERROR: '{joined_str}' નામ એક્સેલની કોલમ 15 (O) માં મળ્યું નથી અથવા ભાવ 0 છે!")
        else:
            try:
                try:
                    base_size = float(calc_size.split("-")[0].strip())
                except:
                    base_size = float(calc_size.split(" ")[0].strip())
                    
                cut_fluo = f"{cut}-{fluorescence}" 
                
                row_0 = df_tables.iloc[0].fillna('').astype(str).str.strip().str.upper()
                row_1 = df_tables.iloc[1].fillna('').astype(str).str.strip().str.upper()
                
                if cut_fluo not in row_0.values:
                    st.error(f"Excel ની Tables શીટમાં '{cut_fluo}' નામનું કોઈ હેડિંગ મળ્યું નહિ!")
                else:
                    size_col_idx = None
                    color_col_idx = None
                    
                    for idx, val in row_1.items():
                        if val == 'SIZE':
                            size_col_idx = idx
                        elif val == 'COLOR':
                            color_col_idx = idx
                            
                    if size_col_idx is None: size_col_idx = 1
                    if color_col_idx is None: color_col_idx = 2

                    cut_fluo_idx = row_0[row_0 == cut_fluo].index[0]
                    
                    clarity_list = ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"]
                    if clarity in clarity_list:
                        clarity_offset = clarity_list.index(clarity)
                        final_col_idx = cut_fluo_idx + clarity_offset
                        
                        data_rows = df_tables.iloc[2:].copy()
                        
                        size_match = pd.to_numeric(data_rows[size_col_idx], errors='coerce') == base_size
                        color_match = data_rows[color_col_idx].astype(str).str.strip().str.upper() == color
                        
                        match_data = data_rows[size_match & color_match]
                        
                        if not match_data.empty:
                            discount_percent = float(match_data.iloc[0, final_col_idx])
                            
                            rate_per_cts = calc_rap * (1 + (discount_percent / 100))
                            pol_amt = rate_per_cts * polish_weight
                            
                            # ==========================================
                            # POINT 3: 4 COLUMN 
                            # ==========================================
                            st.subheader("📊 ગણતરીનું પરિણામ")
                            
                            res_col1, res_col2, res_col3, res_col4 = st.columns(4)
                            
                            res_col1.metric("Rap Price ($)", f"${calc_rap:,.2f}")
                            res_col2.metric("Discount / Prem. (%)", f"{discount_percent}%")
                            res_col3.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
                            res_col4.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
                            
                            st.success("🎉 તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
                        else:
                            st.warning(f"આ સાઈઝ ({base_size}) અને કલર ({color}) માટે ડિસ્કાઉન્ટ ડેટા મળ્યો નથી.")
                    else:
                        st.error("ક્લેરિટીનું નામ મેચ થતું નથી.")
            except Exception as e:
                st.error(f"ગણતરીમાં અણધારી ભૂલ આવી: {e}")
