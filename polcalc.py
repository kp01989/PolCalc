import streamlit as st
import pandas as pd

# --- પેજનું સેટિંગ ---
st.set_page_config(page_title="Diamond Polish Calc", layout="wide")

st.title("💎 Diamond Polish Calculator")
st.markdown("તમારા ડાયમંડની ડિટેલ્સ નીચે સિલેક્ટ કરો એટલે ઓટોમેટિક ગણતરી થઈ જશે.")

# --- 1. એક્સેલમાંથી ડેટા લોડ કરવાનું ફંક્શન ---
@st.cache_data
def load_data():
    # અહીં તમારી ફાઈલનું એક્ઝેક્ટ નામ લખેલું છે
    file_path = "Polish Calc_Updated_05-05-2026 - Copy.xlsm" 
    try:
        df_tables = pd.read_excel(file_path, sheet_name="Tables")
        df_list = pd.read_excel(file_path, sheet_name="List")
        return df_tables, df_list
    except Exception as e:
        st.error(f"એક્સેલ ફાઈલ લોડ કરવામાં એરર આવી છે. ખાતરી કરો કે ફાઈલ એ જ ફોલ્ડરમાં છે. એરર: {e}")
        return None, None

df_tables, df_list = load_data()

# જો ફાઈલ બરાબર લોડ થઈ જાય તો જ આગળનું UI બતાવશે
if df_tables is not None:
    
    # --- 2. યુઝર ઇનપુટ (બધા ડ્રોપડાઉન મેનુ) ---
    st.subheader("હીરાની વિગતો સિલેક્ટ કરો")
    
    # સ્ક્રીનને 3 ભાગમાં વહેંચી છે જેથી બધું વ્યવસ્થિત દેખાય
    col1, col2, col3 = st.columns(3)
    
    with col1:
        shape = st.selectbox("SHAPE (શેપ)", ["ROUND", "PRINCESS", "OVAL", "MARQUISE", "PEAR", "EMERALD"])
        color = st.selectbox("Color (કલર)", ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"])
        clarity = st.selectbox("Clarity (ક્લેરિટી)", ["FL", "IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3"])
        
    with col2:
        cut = st.selectbox("Cut Group (કટ)", ["3EX", "EX", "VG", "GD", "FAIR"])
        fluorescence = st.selectbox("Fluorescence (ફ્લોરોસન્સ)", ["NON", "FNT", "MED", "STG", "VSTG"])
        # સાઈઝ રેન્જ માટે ડ્રોપડાઉન 
        size_range = st.selectbox("Size (સાઈઝ)", ["0.18 - 0.22", "0.23 - 0.29", "0.30 - 0.34", "0.35 - 0.39", "0.40 - 0.49"])
        
    with col3:
        polish_weight = st.number_input("Polish Weight (વજન)", min_value=0.01, value=0.30, step=0.01)
        rap_price = st.number_input("Rap Price ($)", min_value=0.0, value=1700.0, step=100.0)

    st.divider()

    # --- 3. ગણતરી (Calculation Logic) ---
    # જ્યારે યુઝર 'Calculate' બટન દબાવશે ત્યારે જ આ ગણતરી થશે
    if st.button("Calculate Discount & Amount", type="primary"):
        try:
            # એક્સેલની List શીટ મુજબ વજનને બેઝ સાઈઝમાં ફેરવવાનું લોજિક (જેમ કે '0.30 - 0.34' માંથી 0.30 લેવું)
            base_size = float(size_range.split(" ")[0]) 
            
            # --- VLOOKUP અને INDEX-MATCH નું પાયથોન વર્ઝન ---
            # નોંધ: અહીં df_tables['Size'] અને df_tables['Color'] એ તમારી એક્સેલ શીટના હેડિંગના નામ છે. 
            # જો એક્સેલમાં નામ અલગ હોય તો અહિયાં બદલવા પડશે.
            condition = (df_tables['Size'] == base_size) & (df_tables['Color'] == color)
            match_data = df_tables[condition]
            
            discount_percent = 0.0
            
            if not match_data.empty:
                # ક્લેરિટી (દા.ત. VS1) વાળી કોલમમાંથી ડેટા ખેંચશે
                discount_percent = float(match_data[clarity].values[0])
            else:
                st.warning("આ કોમ્બિનેશન માટે Tables શીટમાં કોઈ ડિસ્કાઉન્ટ મળ્યું નથી! (ડેમો માટે -30% ગણીએ છીએ)")
                discount_percent = -30.0 # જો ડેટા ના મળે તો અત્યારે ડેમો માટે 30% માઇનસ ગણશે

            # રેટ અને એમાઉન્ટની ગણતરી
            # જો ડિસ્કાઉન્ટ માઇનસમાં હોય (જેમ કે -30), તો + કરવાથી તે આપમેળે બાદ થઈ જશે
            rate_per_cts = rap_price * (1 + (discount_percent / 100))
            pol_amt = rate_per_cts * polish_weight
            
            # --- 4. ફાઇનલ પરિણામ (Output) ---
            st.subheader("📊 ગણતરીનું પરિણામ")
            
            res_col1, res_col2, res_col3 = st.columns(3)
            res_col1.metric("Discount / Prem. (%)", f"{discount_percent}%")
            res_col2.metric("Rate $ / Cts", f"${rate_per_cts:,.2f}")
            res_col3.metric("Pol Amt ($)", f"${pol_amt:,.2f}")
            
            st.success("તમારો ડેટા સફળતાપૂર્વક કેલ્ક્યુલેટ થઈ ગયો છે!")
            
        except Exception as e:
            st.error(f"ગણતરીમાં કોઈ ભૂલ આવી છે. કૃપા કરીને એક્સેલના હેડિંગ ચેક કરો. એરર: {e}")