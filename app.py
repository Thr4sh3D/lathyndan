import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from pandasai import SmartDataframe
from pandasai.llm import OpenAI

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="Jaktprovsstatistik & AI", page_icon="🐕", layout="wide")

st.title("📊 Jaktprovsstatistik & AI-Assistent")
st.markdown("Analysera resultat från **Eftersök** och **Jaktprov/Fält**.")

# --- 1. API-NYCKEL FÖR AI (SIDEBAR) ---
st.sidebar.header("🤖 AI-Inställningar")
api_key = st.sidebar.text_input("OpenAI API Key (valfritt)", type="password", help="Krävs för att använda AI-assistenten i fliken.")
st.sidebar.markdown("[Skaffa API-nyckel här](https://platform.openai.com/api-keys)")

# --- 2. LADDA UPP FILER ---
st.sidebar.header("📂 Ladda upp data")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (.xlsx)", type=["xlsx"])
uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"])
uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"])
uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"])

# --- HJÄLPFUNKTION: TOLKA TEXTFIL ---
def parse_txt_to_df(txt_file):
    content = txt_file.getvalue()
    try:
        text_data = content.decode("latin-1")
    except:
        text_data = content.decode("utf-8", errors="ignore")
    
    lines = text_data.split('\n')
    data = []
    for line in lines:
        if re.match(r'^\d{2}-\d{2}-\d{2}', line):
            parts = line.split('\t')
            if len(parts) > 30:
                try:
                    entry = {
                        "Datum": parts[0],
                        "Klass": parts[1],
                        "Hund": parts[2],
                        "regnr": parts[3],
                        "Ras_Clean": parts[4].strip(),
                        "Vatten_Final": pd.to_numeric(parts[27], errors='coerce'),
                        "Spår_Final": pd.to_numeric(parts[28], errors='coerce'),
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except: continue
    return pd.DataFrame(data)

# --- 3. LADDA OCH SLÅ IHOP DATA ---
@st.cache_data
def load_all_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    
    # Excel
    if excel:
        try:
            xls = pd.ExcelFile(excel)
            sheet_e = next((s for s in xls.sheet_names if "eftersök" in s.lower()), None)
            sheet_j = next((s for s in xls.sheet_names if "jakt" in s.lower() or "fält" in s.lower()), None)
            if sheet_e: df_e_list.append(pd.read_excel(xls, sheet_name=sheet_e))
            if sheet_j: df_j_list.append(pd.read_excel(xls, sheet_name=sheet_j))
        except: pass

    # CSV
    if csv_e:
        try:
            df = pd.read_csv(csv_e, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_e, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_e_list.append(df)
        except: pass
    if csv_j:
        try:
            df = pd.read_csv(csv_j, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_j, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_j_list.append(df)
        except: pass

    # TXT
    if txt:
        df_txt = parse_txt_to_df(txt)
        if not df_txt.empty: df_e_list.append(df_txt)

    df_e_tot = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j_tot = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    return df_e_tot, df_j_tot

df_e_raw, df_j_raw = load_all_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)

if df_e_raw.empty and df_j_raw.empty:
    st.info("👈 Ladda upp data för att starta.")
    st.stop()

# --- 4. TVÄTTA DATA ---
def clean_df(df):
    if df.empty: return df
    df.columns = df.columns.str.strip()
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['År'] = df['Datum'].dt.year
    
    # Ras
    if 'Ras_Clean' not in df.columns:
        ras_candidates = ['rasnamn', 'Ras', 'Hundras']
        found_ras = next((c for c in ras_candidates if c in df.columns), None)
        df['Ras_Clean'] = df[found_ras] if found_ras else "Okänd"
    
    # Poäng (om de inte redan heter _Final från TXT)
    if 'Vatten_Final' not in df.columns:
        col_v = next((c for c in df.columns if "vatten" in c.lower() and "kritik" not in c.lower()), None)
        if col_v: df['Vatten_Final'] = pd.to_numeric(df[col_v], errors='coerce').fillna(0)
    
    if 'Spår_Final' not in df.columns:
        col_s = next((c for c in df.columns if ("spår" in c.lower() or "spar" in c.lower()) and "kritik" not in c.lower()), None)
        if col_s: df['Spår_Final'] = pd.to_numeric(df[col_s], errors='coerce').fillna(0)
        
    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- 5. FILTER ---
st.sidebar.divider()
st.sidebar.header("🔍 Filter")
all_years = sorted(list(set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))))
all_races = sorted(list(set(df_e.get('Ras_Clean', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras_Clean', pd.Series()).dropna().astype(str))))

valda_raser = st.sidebar.multiselect("Ras", all_races, default=all_races[:1] if all_races else None)
valda_ar = st.sidebar.multiselect("År", all_years, default=all_years)

def apply_filter(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser: temp = temp[temp['Ras_Clean'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    return temp

df_e_filt = apply_filter(df_e)
df_j_filt = apply_filter(df_j)

# --- 6. FLIKAR ---
tab1, tab2, tab3 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "🤖 AI-Assistent"])

# === FLIK 1: EFTERSÖK ===
with tab1:
    if not df_e_filt.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Starter", len(df_e_filt))
        
        if 'Vatten_Final' in df_e_filt.columns and 'Spår_Final' in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt['Vatten_Final'] >= 4) & (df_e_filt['Spår_Final'] >= 4)]
            full = df_e_filt[(df_e_filt['Vatten_Final'] == 10) & (df_e_filt['Spår_Final'] == 10)]
            c2.metric("Godkända", len(godkanda))
            c3.metric("10-10", len(full))
            
            cc1, cc2 = st.columns(2)
            with cc1: st.plotly_chart(px.histogram(df_e_filt, x='Vatten_Final', nbins=11, title="Vatten", color_discrete_sequence=['#3366CC']), use_container_width=True)
            with cc2: st.plotly_chart(px.histogram(df_e_filt, x='Spår_Final', nbins=11, title="Spår", color_discrete_sequence=['#109618']), use_container_width=True)
            
        st.dataframe(df_e_filt, use_container_width=True, hide_index=True)
    else: st.info("Ingen data.")

# === FLIK 2: FÄLT ===
with tab2:
    if not df_j_filt.empty:
        st.metric("Starter", len(df_j_filt))
        st.dataframe(df_j_filt, use_container_width=True, hide_index=True)
    else: st.info("Ingen data.")

# === FLIK 3: AI ASSISTENT ===
with tab3:
    st.subheader("🤖 Prata med din statistik")
    st.markdown("Här kan du ställa frågor direkt till datan. **OBS:** Kräver en OpenAI API-nyckel i menyn till vänster.")

    if not api_key:
        st.warning("⚠️ Du måste ange en API-nyckel i vänstermenyn för att aktivera AI:n.")
    else:
        # Välj vilken data vi ska fråga
        dataset_val = st.radio("Vilken data vill du analysera?", ["Eftersök", "Fältprov"], horizontal=True)
        
        target_df = df_e_filt if dataset_val == "Eftersök" else df_j_filt
        
        if target_df.empty:
            st.error("Det finns ingen data i det valda urvalet att analysera.")
        else:
            # Initiera AI
            llm = OpenAI(api_token=api_key)
            sdf = SmartDataframe(target_df, config={"llm": llm})

            # Chattruta
            fraga = st.text_area("Vad vill du veta?", placeholder="T.ex: Vilken hund har bäst snittbetyg på vatten? Eller: Visa ett diagram över resultaten.")
            
            if st.button("Skicka fråga"):
                with st.spinner("AI:n tänker..."):
                    try:
                        svar = sdf.chat(fraga)
                        st.success("Här är svaret:")
                        st.write(svar)
                    except Exception as e:
                        st.error(f"Kunde inte tolka frågan: {e}")

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: df_e_filt.to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: df_j_filt.to_excel(writer, index=False, sheet_name='Fält')
st.download_button("📥 Ladda ner Excel", output.getvalue(), "KLM_Statistik.xlsx", "application/vnd.ms-excel")
