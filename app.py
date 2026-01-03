import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Statistik", page_icon="🐕", layout="wide")

st.title("📊 KLM Statistik & Analys")
st.markdown("Analysera resultat från **Eftersök** och **Jaktprov/Fält**.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 Ladda upp data")
st.sidebar.info("Appen klarar Excel, CSV och gamla Textfiler.")

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

# --- 2. LADDA OCH SLÅ IHOP DATA ---
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

# --- 3. TVÄTTA DATA ---
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
    
    # Poäng
    if 'Vatten_Final' not in df.columns:
        col_v = next((c for c in df.columns if "vatten" in c.lower() and "kritik" not in c.lower()), None)
        if col_v: df['Vatten_Final'] = pd.to_numeric(df[col_v], errors='coerce').fillna(0)
    
    if 'Spår_Final' not in df.columns:
        col_s = next((c for c in df.columns if ("spår" in c.lower() or "spar" in c.lower()) and "kritik" not in c.lower()), None)
        if col_s: df['Spår_Final'] = pd.to_numeric(df[col_s], errors='coerce').fillna(0)
        
    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- 4. FILTER ---
st.sidebar.divider()
st.sidebar.header("🔍 Filter")

if df_e.empty and df_j.empty:
    st.info("👈 Börja med att ladda upp en fil i menyn.")
else:
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

# --- 5. VISNING (FLIKAR) ---
tab1, tab2, tab3 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "❓ Hjälp & Guide"])

# === FLIK 1: EFTERSÖK ===
with tab1:
    if not df_e_filt.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Starter", len(df_e_filt))
        c2.metric("Unika Hundar", df_e_filt['regnr'].nunique() if 'regnr' in df_e_filt.columns else 0)
        
        if 'Vatten_Final' in df_e_filt.columns and 'Spår_Final' in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt['Vatten_Final'] >= 4) & (df_e_filt['Spår_Final'] >= 4)]
            full = df_e_filt[(df_e_filt['Vatten_Final'] == 10) & (df_e_filt['Spår_Final'] == 10)]
            c3.metric("Godkända", f"{len(godkanda)} ({int(len(godkanda)/len(df_e_filt)*100)}%)")
            c4.metric("10-10", len(full))
            
            st.divider()
            cc1, cc2 = st.columns(2)
            with cc1: st.plotly_chart(px.histogram(df_e_filt, x='Vatten_Final', nbins=11, title="Vattenbetyg", color_discrete_sequence=['#3366CC']), use_container_width=True)
            with cc2: st.plotly_chart(px.histogram(df_e_filt, x='Spår_Final', nbins=11, title="Spårbetyg", color_discrete_sequence=['#109618']), use_container_width=True)
            
        st.dataframe(df_e_filt, use_container_width=True, hide_index=True)
    else: st.info("Ingen eftersöksdata hittades.")

# === FLIK 2: FÄLT ===
with tab2:
    if not df_j_filt.empty:
        c1, c2 = st.columns(2)
        c1.metric("Starter", len(df_j_filt))
        c2.metric("Unika Hundar", df_j_filt['regnr'].nunique() if 'regnr' in df_j_filt.columns else 0)
        
        st.divider()
        pris_col = next((c for c in df_j_filt.columns if "pris" in c.lower()), None)
        if pris_col:
            pc = df_j_filt[pris_col].value_counts().reset_index()
            pc.columns = ['Pris', 'Antal']
            st.plotly_chart(px.pie(pc, values='Antal', names='Pris', title="Prisfördelning", hole=0.4), use_container_width=True)
            
        st.dataframe(df_j_filt, use_container_width=True, hide_index=True)
    else: st.info("Ingen fältprovsdata hittades.")

# === FLIK 3: HJÄLP & GUIDE ===
with tab3:
    st.markdown("## 📘 Användarguide")
    st.markdown("""
    **1. Hämta dina filer**
    Ladda ner resultaten från SKK eller din vanliga källa. Appen klarar Excel (.xlsx), CSV (.csv) och Textfiler (.txt).
    
    **2. Ladda upp**
    Använd knapparna i menyn till vänster. Appen slår automatiskt ihop filerna om du laddar upp flera.
    
    **3. Filtrera**
    I menyn kan du välja:
    * **Ras:** Förvald på Kleiner Münsterländer (men du kan välja alla).
    * **År:** Välj vilka år du vill analysera.
    
    **Flikarna:**
    * **🌲 Eftersök:** Statistik för Vatten och Spår. Här ser du automatiskt antal Godkända och "Full pott" (10-10).
    * **🌾 Fältprov:** Statistik för jaktprov/fält.
    * **❓ Hjälp:** Denna guide.
    """)

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: df_e_filt.to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: df_j_filt.to_excel(writer, index=False, sheet_name='Fält')
st.download_button("📥 Ladda ner Excel", output.getvalue(), "KLM_Statistik.xlsx", "application/vnd.ms-excel")
