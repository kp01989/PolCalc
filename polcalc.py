import streamlit as st
import pandas as pd
import msoffcrypto
import io

# --- 1. પેજનું સેટિંગ અને કોમ્પેક્ટ લેઆઉટ ---
st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
        label { font-size: 0.85rem !important; font-weight: bold !important; }
    </style>
""", unsafe_allow_html=True)

st.title("💎 Diamond Polish Calculator")

# --- Compare માટેનું સ્ટોરેજ ---
if 'compare_list' not in st.session_state:
    st.session_state['compare_list'] = []

# --- 2. એક્સેલમાંથી ડેટા લોડ ---
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
    
    # --- 3. ૮ બોક્સ એક જ લાઈનમાં ---
    c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
    
    with c1: shape = st.selectbox("SHAPE", ["ROUND", "PRINCESS", "OVAL", "MARQUISE", "PEAR", "EMERALD"])
    with c2: color = st.selectbox("Color", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"])
    with c3: clarity = st.selectbox("Clarity", ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"])
    with c4: cut = st.selectbox("Cut", ["3EX", "VG", "EX", "GD", "FAIR"])
    with c5: fluorescence = st.selectbox("Fluo.", ["NON", "FNT", "MED", "STG", "VSTG"])
    with c6: polish_weight = st.number_input("Weight", min_value=0.01, value=0.30, step=0.01)

    calc_size = ""
    try:
        size_df = df_list.iloc[:, [11, 13]].copy() 
        size_df.columns = ['Weight', 'SizeLabel']
        size_df['Weight'] = pd.to_numeric(size_df['Weight'], errors='coerce')
        size_df = size_df.dropna(subset=['Weight', 'SizeLabel']).sort_values('Weight')
        for idx, row in size_df.iterrows():
            if polish_weight >= row['Weight']:
                calc_size = str(row['SizeLabel']).strip()
    except:
        calc_size = "Error"

    calc_rap = 0.0
    search_key = ""
    if calc_size and calc_size != "Error":
        search_key = f"{shape.strip()}{calc_size}{clarity.strip()}{color.strip()}".upper()
        try:
            lookup_col = df_list.iloc[:, 14].astype(str).str.strip().str.upper()
            result_col = df_list.iloc[:, 15]
            if search_key in lookup_col.values:
                match_idx = lookup_col[lookup_col == search_key].index[0]
                val = result_col.iloc[match_idx]
                calc_rap = float(pd.to_numeric(val, errors='coerce'))
                if pd.isna(calc_rap): calc_rap = 0.0
        except:
            calc_rap = 0.0

    with c7: st.text_input("Size", value=calc_size, disabled=True)
    with c8: st.text_input("Rap Price", value=f"{calc_rap:,.2f}" if calc_rap else "0.00", disabled=True)

    st.write("")

    # --- 4. બટન્સ ---
    btn1, btn2, empty_space = st.columns([2, 2, 6])
    with btn1:
        calc_clicked = st.button("Calculate", type="primary", use_container_width=True)
    with btn2:
        comp_clicked = st.button("Add to Compare ⚖️", type="secondary", use_container_width=True)

    st.divider()

    # --- ફાઇનલ ગણતરી અને લોજિક ---
    if calc_clicked or comp_clicked:
        if not calc_size or calc_size == "Error":
            st.error("સાઈઝ ઓટોમેટિક કેલ્ક્યુલેટ થઈ શકી નથી.")
        elif calc_rap == 0.0:
            st.error(f"VLOOKUP FAILED: '{search_key}' નામ એક્સેલની કોલમ 15 (O) માં મળ્યું નથી અથવા ત્યાં ભાવ 0 છે!")
        else:
            try:
                cut_fluo = f"{cut}-{fluorescence}" 
                row_0 = df_tables.iloc[0].fillna('').astype(str).str.strip().str.upper()
                row_1 = df_tables.iloc[1].fillna('').astype(str).str.strip().str.upper()
                
                if cut_fluo not in row_0.values:
                    st.error(f"Excel ની Tables શીટમાં '{cut_fluo}' નામનું કોઈ હેડિંગ મળ્યું નહિ!")
                else:
                    size_col_idx = 1
                    color_col_idx = 2
                    for idx, val in row_1.items():
                        if val == 'SIZE': size_col_idx = idx
                        elif val == 'COLOR': color_col_idx = idx

                    cut_fluo_idx = row_0[row_0 == cut_fluo].index[0]
                    clarity_list = ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"]
                    
                    if clarity in clarity_list:
                        clarity_offset = clarity_list.index(clarity)
                        final_col_idx = cut_fluo_idx + clarity_offset
                        data_rows = df_tables.iloc[2:].copy()
                        
                        size_match = pd.to_numeric(data_rows[size_col_idx], errors='coerce') == polish_weight
                        color_match = data_rows[color_col_idx].astype(str).str.strip().str.upper() == color
                        match_data = data_rows[size_match & color_match]
                        
                        if not match_data.empty:
                            discount_percent = float(match_data.iloc[0, final_col_idx])
                            rate_per_cts = calc_rap * (1 + (discount_percent / 100))
                            pol_amt = rate_per_cts * polish_weight
                            
                            # ૫. Calculate દબાવે તો જ રિઝલ્ટ દેખાડવાનું 
                            if calc_clicked:
                                st.subheader("📊 ગણતરીનું પરિણામ")
                                res_col1, res_col2, res_col3, res_col4 = st.columns(4)
                                res_col1.metric("Rap Price ($)", f"${calc_rap:,.2f}")
                                res_col2.metric("Discount / Prem. (%)", f"{discount_percent}%")
                                res_col3.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
                                res_col4.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
                                st.success("🎉 તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
                            
                            # ૬. Compare દબાવે તો ખાલી લિસ્ટમાં ઉમેરવાનું
                            if comp_clicked:
                                st.session_state['compare_list'].append({
                                    "Shape": shape,
                                    "Weight": polish_weight,
                                    "Color": color,
                                    "Clarity": clarity,
                                    "Rap Price": f"${calc_rap:,.2f}",
                                    "Discount": f"{discount_percent}%",
                                    "Rate/Cts": f"${rate_per_cts:,.2f}",
                                    "Total Amount": f"${pol_amt:,.2f}"
                                })
                                st.success("✅ આ હીરો Compare લિસ્ટમાં ઉમેરાઈ ગયો છે!")
                        else:
                            st.warning(f"આ વજન ({polish_weight}) અને કલર ({color}) માટે Tables શીટમાં ડિસ્કાઉન્ટ ડેટા મળ્યો નથી.")
                    else:
                        st.error("ક્લેરિટીનું નામ મેચ થતું નથી.")
            except Exception as e:
                st.error(f"ગણતરીમાં અણધારી ભૂલ આવી: {e}")

    # --- 7. ઊભું (Vertical) Compare ટેબલ દેખાડવું ---
    if st.session_state['compare_list']:
        st.subheader("⚖️ ડાયમંડ કમ્પેરીઝન ટેબલ (Comparison)")
        
        # ડેટાફ્રેમ બનાવી
        df_compare = pd.DataFrame(st.session_state['compare_list'])
        
        # ડેટાફ્રેમને ઊભી (Transpose) કરી
        df_vertical = df_compare.T
        
        # કોલમના નામ Diamond 1, Diamond 2 એમ આપી દીધા
        df_vertical.columns = [f"Diamond {i+1}" for i in range(len(df_vertical.columns))]
        
        # આનાથી ટેબલ એકદમ મસ્ત ઊભું દેખાશે
        st.dataframe(df_vertical, use_container_width=True)
        
        if st.button("🗑️ લિસ્ટ ક્લિયર કરો"):
            st.session_state['compare_list'] = []
            st.rerun()
