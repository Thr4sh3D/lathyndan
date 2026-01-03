import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="Jaktprovsstatistik", page_icon="🐕", layout="wide")

st.title("📊 Jaktprovsstatistik & Analys (Alla format)")
st.markdown("Stödjer **Excel (.xlsx)**, **CSV (.csv)** och **SKK Textfiler (.txt)**.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 Ladda upp data")
st.sidebar.info("Appen slår ihop data från alla filer du laddar upp.")

uploaded_excel = st.sidebar.file_uploader("Excel-fil (.xlsx)", type=["xlsx"])
uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"])
uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"])
uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"])

# --- HJÄLPFUNKTION: TOLKA TEXTFIL (FRÅN LAT-HUNDEN) ---
def parse_txt_to_df(txt_file):
    # Denna funktion konverterar den gamla textfilen till samma format som Excel
    content = txt_file.getvalue()
    try:
        text_data = content.decode("latin-1")
    except:
        text_data = content.decode("utf-8", errors="ignore")
    
    lines = text_data.split('\n')
    data = []
    
    for line in lines:
        # Leta efter rader som ser ut som provdata (börjar med datum)
        if re.match(r'^\d{2}-\d{2}-\d{2}', line):
            parts = line.split('\t')
            if len(parts) > 30:
                try:
                    # Mappa kolumner manuellt baserat på din gamla specifikation
                    entry = {
                        "Datum": parts[0],
                        "Klass": parts[1],
                        "Hund": parts[2],
                        "regnr": parts[3],
                        "Ras_Clean": parts[4].strip(), # Vi använder denna för filter
                        "Vatten": pd.to_numeric(parts[27], errors='coerce'), # Index 27
                        "Spår": pd.to_numeric(parts[28], errors='coerce'),   # Index 28
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except:
                    continue
    return pd.DataFrame(data)

# --- 2. LADDA OCH SLÅ IHOP DATA ---
@st.cache_data
def load_all_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    
    # 1. EXCEL
    if excel:
        try:
            xls = pd.ExcelFile(excel)
            sheet_e = next((s for s in xls.sheet_names if "eftersök" in s.lower()), None)
            sheet_j = next((s for s in xls.sheet_names if "jakt" in s.lower() or "fält" in s.lower()), None)
            if sheet_e: df_e_list.append(pd.read_excel(xls, sheet_name=sheet_e))
            if sheet_j: df_j_list.append(pd.read_excel(xls, sheet_name=sheet_j))
        except Exception as e:
            st.error(f"Excel-fel: {e}")

    # 2. CSV (Eftersök)
    if csv_e:
        try:
            df = pd.read_csv(csv_e, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_e, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_e_list.append(df)
        except: pass

    # 3. CSV (Jakt/Fält)
    if csv_j:
        try:
            df = pd.read_csv(csv_j, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_j, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_j_list.append(df)
        except: pass

    # 4. TEXTFIL (Antas vara Eftersök baserat på struktur)
    if txt:
        df_txt = parse_txt_to_df(txt)
        if not df_txt.empty:
            df_e_list.append(df_txt)

    # SLÅ IHOP ALLT
    df_e_total = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j_total = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()

    return df_e_total, df_j_total

df_e_raw, df_j_raw = load_all_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)

if df_e_raw.empty and df_j_raw.empty:
    st.info("👈 Ladda upp valfri filtyp (Excel, CSV eller TXT) för att starta.")
    st.stop()

# --- 3. STANDARDISERA DATA (TVÄTTA) ---
def standardisera_df(df):
    if df.empty: return df
    
    # 1. Rensa kolumnnamn
    df.columns = df.columns.str.strip()
    
    # 2. Fixa datum
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['År'] = df['Datum'].dt.year
    
    # 3. Hitta Ras (om den inte redan är fixad från txt-filen)
    if 'Ras_Clean' not in df.columns:
        ras_candidates = ['rasnamn', 'Ras', 'Hundras']
        found_ras = next((c for c in ras_candidates if c in df.columns), None)
        df['Ras_Clean'] = df[found_ras] if found_ras else "Okänd"
    
    # 4. Normalisera Poängkolumner (så Vatten heter Vatten oavsett källa)
    # Leta efter kolumn som heter något med "vatten"
    col_vatten = next((c for c in df.columns if "vatten" in c.lower() and "kritik" not in c.lower()), None)
    col_spar = next((c for c in df.columns if ("spår" in c.lower() or "spar" in c.lower()) and "kritik" not in c.lower()), None)
    
    if col_vatten: df['Vatten_Final'] = pd.to_numeric(df[col_vatten], errors='coerce').fillna(0)
    if col_spar: df['Spår_Final'] = pd.to_numeric(df[col_spar], errors='coerce').fillna(0)

    return df

df_e = standardisera_df(df_e_raw)
df_j = standardisera_df(df_j_raw)

# --- 4. FILTRERING (SIDEBAR) ---
st.sidebar.divider()
st.sidebar.header("🔍 Filter")

# Hämta unika värden för filter
all_years = sorted(list(set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))))
all_races = sorted(list(set(df_e.get('Ras_Clean', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras_Clean', pd.Series()).dropna().astype(str))))

valda_raser = st.sidebar.multiselect("Ras", all_races, default=all_races[:1] if all_races else None)
valda_ar = st.sidebar.multiselect("År", all_years, default=all_years)
sok_hund = st.sidebar.text_input("Sök hund/regnr", "")

# --- 5. APPLICERA FILTER ---
def filter_now(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser: temp = temp[temp['Ras_Clean'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    if sok_hund:
        temp['search'] = temp.astype(str).agg(' '.join, axis=1).str.lower()
        temp = temp[temp['search'].str.contains(sok_hund.lower())]
        temp = temp.drop(columns=['search'])
    return temp

df_e_filt = filter_now(df_e)
df_j_filt = filter_now(df_j)

# --- 6. VISUALISERING ---
tab1, tab2 = st.tabs(["🌲 Eftersök", "🌾 Fältprov"])

# === FLIK 1: EFTERSÖK ===
with tab1:
    if df_e_filt.empty:
        st.warning("Ingen data.")
    else:
        # KPIer
        total = len(df_e_filt)
        unika = df_e_filt['regnr'].nunique() if 'regnr' in df_e_filt.columns else 0
        
        # Räkna godkända (om vi har poängen)
        if 'Vatten_Final' in df_e_filt.columns and 'Spår_Final' in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt['Vatten_Final'] >= 4) & (df_e_filt['Spår_Final'] >= 4)]
            full_pott = df_e_filt[(df_e_filt['Vatten_Final'] == 10) & (df_e_filt['Spår_Final'] == 10)]
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Antal Starter", total)
            c2.metric("Unika Hundar", unika)
            c3.metric("Godkända", f"{len(godkanda)} ({int(len(godkanda)/total*100)}%)")
            c4.metric("10-10 (Full pott)", len(full_pott))
            
            st.divider()
            
            # Diagram
            cc1, cc2 = st.columns(2)
            with cc1:
                fig_v = px.histogram(df_e_filt, x='Vatten_Final', nbins=11, title="Vattenbetyg", color_discrete_sequence=['#3366CC'])
                fig_v.update_layout(bargap=0.2)
                st.plotly_chart(fig_v, use_container_width=True)
            with cc2:
                fig_s = px.histogram(df_e_filt, x='Spår_Final', nbins=11, title="Spårbetyg", color_discrete_sequence=['#109618'])
                fig_s.update_layout(bargap=0.2)
                st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.warning("Hittade inte poängkolumnerna.")

        st.dataframe(df_e_filt, use_container_width=True, hide_index=True)

# === FLIK 2: FÄLTPROV ===
with tab2:
    if df_j_filt.empty:
        st.warning("Ingen data.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("Antal Starter", len(df_j_filt))
        c2.metric("Unika Hundar", df_j_filt['regnr'].nunique() if 'regnr' in df_j_filt.columns else 0)
        
        st.divider()
        
        # Försök hitta Pris-kolumnen för diagram
        pris_col = next((c for c in df_j_filt.columns if "pris" in c.lower()), None)
        if pris_col:
            pc = df_j_filt[pris_col].value_counts().reset_index()
            pc.columns = ['Pris', 'Antal']
            fig_p = px.pie(pc, values='Antal', names='Pris', title="Prisfördelning", hole=0.4)
            st.plotly_chart(fig_p, use_container_width=True)

        st.dataframe(df_j_filt, use_container_width=True, hide_index=True)

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: df_e_filt.to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: df_j_filt.to_excel(writer, index=False, sheet_name='Fält')

st.download_button("📥 Ladda ner resultatet (Excel)", output.getvalue(), "KLM_Statistik_Komplett.xlsx", "application/vnd.ms-excel")
