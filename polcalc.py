import streamlit as st
import pandas as pd
import msoffcrypto
import io

st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
        
        label { 
            font-size: 0.85rem !important; 
            font-weight: bold !important; 
            padding-left: 5px !important; 
        }
        
        .custom-compare-table {
            font-size: 0.95rem !important; 
            width: 100% !important; 
            border-collapse: collapse;
            margin-top: 5px; 
            margin-bottom: 20px;
        }
        .custom-compare-table th, .custom-compare-table td {
            padding: 4px 6px !important; 
            border: 1px solid rgba(128, 128, 128, 0.3);
            text-align: center !important; 
            vertical-align: middle !important; 
        }
        .custom-compare-table th {
            background-color: rgba(128, 128, 128, 0.1);
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

st.title("💎 Diamond Polish Calculator")
st.markdown("Select your diamond details below for automatic calculation.")

# --- Session State સ્ટોરેજ ---
if 'compare_list' not in st.session_state:
    st.session_state['compare_list'] = []
if 'calc_result' not in st.session_state:
    st.session_state['calc_result'] = None

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
        st.error(f"Error opening file: {e}")
        return None, None

data_loaded = load_diamond_data()

if data_loaded[0] is not None:
    df_tables, df_list = data_loaded
    
    c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns(9)
    
    with c1: shape = st.text_input("Shape", value="ROUND", disabled=True)
    
    # દરેક બોક્સને 'key' આપી છે જેથી તેને ક્લિયર કરી શકાય
    with c2: polish_weight = st.number_input("Weight", min_value=0.01, value=None, placeholder="0.00", step=0.01, key="weight_input")
    with c3: color = st.selectbox("Color", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], index=None, placeholder="Select", key="color_input")
    with c4: clarity = st.selectbox("Clarity", ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"], index=None, placeholder="Select", key="clarity_input")
    with c5: cut = st.selectbox("Cut Group", ["3EX", "VG", "EX", "GD", "FAIR"], index=None, placeholder="Select", key="cut_input")
    with c6: fluorescence = st.selectbox("Fluorescence", ["NON", "FNT", "MED", "STG", "VSTG"], index=None, placeholder="Select", key="fluo_input")
    with c7: add_disc = st.number_input("Additional Disc. %", value=None, placeholder="0.00", step=0.50, format="%.2f", key="add_disc_input")

    calc_size = ""
    if polish_weight is not None:
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
    if polish_weight is not None and calc_size and calc_size != "Error" and color is not None and clarity is not None:
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

    with c8: st.text_input("Size", value=calc_size if (color is not None and polish_weight is not None) else "", disabled=True)
    with c9: st.text_input("Rap Price", value=f"{calc_rap:,.2f}" if calc_rap else "0.00", disabled=True)

    st.write("")

    btn1, btn2, empty_space = st.columns([2, 2, 6])
    with btn1:
        calc_clicked = st.button("Calculate", type="primary", use_container_width=True)
    with btn2:
        comp_clicked = st.button("Add to Compare ⚖️", type="secondary", use_container_width=True)

    if calc_clicked or comp_clicked:
        if polish_weight is None:
            st.warning("⚠️ Please enter the Weight before calculating.")
        elif None in [color, clarity, cut, fluorescence]:
            st.warning("⚠️ Please select Color, Clarity, Cut Group, and Fluorescence before calculating.")
        elif not calc_size or calc_size == "Error":
            st.error("Size could not be calculated automatically.")
        elif calc_rap == 0.0:
            st.error(f"VLOOKUP FAILED: '{search_key}' not found in Excel Column 15 (O) or price is 0!")
        else:
            try:
                actual_add_disc = add_disc if add_disc is not None else 0.00
                fluo_short = fluorescence
                if fluorescence == "Fluorescence": fluo_short = "NON" 
                
                cut_fluo = f"{cut[:3] if cut == '3EX' else cut}-{fluorescence}" 
                row_0 = df_tables.iloc[0].fillna('').astype(str).str.strip().str.upper()
                row_1 = df_tables.iloc[1].fillna('').astype(str).str.strip().str.upper()
                
                if cut_fluo not in row_0.values:
                    st.error(f"Heading '{cut_fluo}' not found in Excel Tables sheet!")
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
                        
                        try:
                            base_size = float(calc_size.split("-")[0].strip())
                        except:
                            try:
                                base_size = float(calc_size.split(" ")[0].strip())
                            except:
                                base_size = polish_weight
                                
                        size_match = pd.to_numeric(data_rows[size_col_idx], errors='coerce') == base_size
                        color_match = data_rows[color_col_idx].astype(str).str.strip().str.upper() == color
                        match_data = data_rows[size_match & color_match]
                        
                        if not match_data.empty:
                            base_discount = float(match_data.iloc[0, final_col_idx])
                            final_discount = base_discount + actual_add_disc 
                            
                            rate_per_cts = calc_rap * (1 + (final_discount / 100))
                            pol_amt = rate_per_cts * polish_weight
                            
                            diamond_summary = f"{shape}-{polish_weight}-{color}-{clarity}-{cut}-{fluorescence}"
                            
                            if calc_clicked:
                                # રિઝલ્ટને ડેટાબેઝ(Session) માં સેવ કર્યું જેથી ક્લિયર થયા પછી પણ દેખાય
                                st.session_state['calc_result'] = {
                                    "summary": diamond_summary,
                                    "rap": calc_rap,
                                    "disc": final_discount,
                                    "rate": rate_per_cts,
                                    "amt": pol_amt
                                }
                                # બધા બોક્સ ઓટો-ક્લિયર કરી દીધા
                                for key in ["weight_input", "color_input", "clarity_input", "cut_input", "fluo_input", "add_disc_input"]:
                                    st.session_state[key] = None
                                st.rerun() # પેજ તરત રિફ્રેશ મારશે
                            
                            if comp_clicked:
                                if len(st.session_state['compare_list']) >= 10:
                                    st.warning("⚠️ List is full! You can compare a maximum of 10 diamonds. Please download the data below.")
                                else:
                                    st.session_state['compare_list'].append({
                                        "Shape": shape,
                                        "Weight": polish_weight,
                                        "Color": color,
                                        "Clarity": clarity,
                                        "Cut Group": cut,
                                        "Fluorescence": fluorescence,   
                                        "Additional Disc. %": f"{actual_add_disc:.2f}%",     
                                        "Rap Price": f"${calc_rap:,.2f}",
                                        "Total Disc": f"{final_discount:.2f}%", 
                                        "Rate/Cts": f"${rate_per_cts:,.2f}",
                                        "Total Amount": f"${pol_amt:,.2f}"
                                    })
                                    st.session_state['calc_result'] = None # નવું રિઝલ્ટ આવે એટલે જૂનું કાઢી નાખ્યું
                                    # બધા બોક્સ ઓટો-ક્લિયર કરી દીધા
                                    for key in ["weight_input", "color_input", "clarity_input", "cut_input", "fluo_input", "add_disc_input"]:
                                        st.session_state[key] = None
                                    st.rerun()
                        else:
                            st.warning(f"Discount data not found in Tables sheet for Weight ({polish_weight}) and Color ({color}).")
                    else:
                        st.error("Clarity name does not match.")
            except Exception as e:
                st.error(f"Unexpected error in calculation: {e}")

    # --- સેવ થયેલું રિઝલ્ટ દેખાડવા માટે ---
    if st.session_state['calc_result']:
        res = st.session_state['calc_result']
        st.divider()
        st.subheader(f"📊 Calculation Result: {res['summary']}")
        res_col1, res_col2, res_col3, res_col4 = st.columns(4)
        res_col1.metric("Rap Price ($)", f"${res['rap']:,.2f}")
        res_col2.metric("Total Disc (%)", f"{res['disc']:.2f}%")
        res_col3.metric("Rate $ / Cts", f"${res['rate']:,.2f}")
        res_col4.metric("Pol Amt ($)", f"${res['amt']:,.2f}")
        st.success("🎉 Data calculated successfully!")

    if st.session_state['compare_list']:
        st.write("") 
        current_len = len(st.session_state['compare_list'])
        st.subheader(f"⚖️ Diamond Comparison Table (Total: {current_len}/10)")
        
        df_compare = pd.DataFrame(st.session_state['compare_list'])
        df_vertical = df_compare.T
        df_vertical.columns = [f"D {i+1}" for i in range(len(df_vertical.columns))]
        
        st.markdown(df_vertical.to_html(classes="custom-compare-table"), unsafe_allow_html=True)
        
        col_clear, col_download, col_empty = st.columns([2, 3, 5])
        
        with col_clear:
            if st.button("🗑️ Clear List", use_container_width=True):
                st.session_state['compare_list'] = []
                st.rerun()
                
        with col_download:
            if current_len >= 2:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_vertical.to_excel(writer, sheet_name='Comparison')
                output.seek(0)
                
                st.download_button(
                    label=f"📥 Download Data for {current_len} Diamonds",
                    data=output,
                    file_name="Diamond_Comparison.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
