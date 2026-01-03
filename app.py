import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="Jaktprovsstatistik", page_icon="🐕", layout="wide")

st.title("📊 Jaktprovsstatistik & Analys")
st.markdown("Analysera resultat från **Eftersök** och **Jaktprov/Fält**.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 1. Ladda upp data")
st.sidebar.markdown("Ladda upp Excel-filen (med flikar) ELLER separata CSV-filer.")

uploaded_excel = st.sidebar.file_uploader("Excel-fil (.xlsx)", type=["xlsx"])
uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"])
uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"])

# Funktion för att ladda data robust
@st.cache_data
def load_data(excel_file, csv_e, csv_j):
    df_e = pd.DataFrame()
    df_j = pd.DataFrame()
    
    # Alternativ A: Excel
    if excel_file:
        try:
            xls = pd.ExcelFile(excel_file)
            # Leta efter flikar
            sheet_e = next((s for s in xls.sheet_names if "eftersök" in s.lower()), None)
            sheet_j = next((s for s in xls.sheet_names if "jakt" in s.lower() or "fält" in s.lower()), None)
            
            if sheet_e: df_e = pd.read_excel(xls, sheet_name=sheet_e)
            if sheet_j: df_j = pd.read_excel(xls, sheet_name=sheet_j)
        except Exception as e:
            st.error(f"Fel vid inläsning av Excel: {e}")

    # Alternativ B: CSV (Skriver över Excel om båda finns)
    if csv_e:
        try:
            # Provar läsa med semikolon som är standard för SKK/Excel-CSV
            df_e = pd.read_csv(csv_e, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df_e.columns) < 2: # Om det misslyckades, testa komma
                csv_e.seek(0)
                df_e = pd.read_csv(csv_e, sep=',', encoding='utf-8', on_bad_lines='skip')
        except:
            st.warning("Kunde inte läsa Eftersöks-CSV. Kontrollera formatet.")

    if csv_j:
        try:
            df_j = pd.read_csv(csv_j, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df_j.columns) < 2:
                csv_j.seek(0)
                df_j = pd.read_csv(csv_j, sep=',', encoding='utf-8', on_bad_lines='skip')
        except:
            st.warning("Kunde inte läsa Jaktprovs-CSV.")

    return df_e, df_j

# Ladda datan
df_eftersok_raw, df_jakt_raw = load_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j)

# Kontroll om data finns
if df_eftersok_raw.empty and df_jakt_raw.empty:
    st.info("👈 Börja med att ladda upp filer i menyn till vänster.")
    st.stop()

# --- 2. TVÄTTA OCH FÖRBEREDA DATA ---
def clean_df(df):
    if df.empty: return df
    # Ta bort mellanslag i kolumnnamn
    df.columns = df.columns.str.strip()
    
    # Hantera Datum
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['År'] = df['Datum'].dt.year
    
    # Hitta Raskolumn (ofta 'rasnamn' eller 'Ras')
    ras_candidates = ['rasnamn', 'Ras', 'Hundras']
    found_ras = next((c for c in ras_candidates if c in df.columns), None)
    if found_ras:
        df['Ras_Clean'] = df[found_ras]
    else:
        df['Ras_Clean'] = "Okänd"
        
    return df

df_e = clean_df(df_eftersok_raw)
df_j = clean_df(df_jakt_raw)

# --- 3. FILTER (SIDEBAR) ---
st.sidebar.divider()
st.sidebar.header("🔍 Filtrering")

# Kombinera filterval från båda dataseten
all_years = sorted(list(set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))))
all_races = sorted(list(set(df_e.get('Ras_Clean', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras_Clean', pd.Series()).dropna().astype(str))))

# A. Rasfilter
valda_raser = st.sidebar.multiselect("Välj Ras(er)", all_races, default=all_races[:1] if all_races else None)

# B. Årsfilter
valda_ar = st.sidebar.multiselect("Välj År", all_years, default=all_years)

# C. Sök på hund
sok_hund = st.sidebar.text_input("Sök på Hundnamn eller Regnr", "")

# --- 4. APPLICERA FILTER ---
def filter_dataframe(df):
    if df.empty: return df
    temp = df.copy()
    
    # Ras
    if valda_raser:
        temp = temp[temp['Ras_Clean'].isin(valda_raser)]
    
    # År
    if valda_ar and 'År' in temp.columns:
        temp = temp[temp['År'].isin(valda_ar)]
        
    # Text-sökning (Regnr eller Namn)
    if sok_hund:
        # Skapa en söksträng av alla kolumner
        temp['search_col'] = temp.astype(str).agg(' '.join, axis=1).str.lower()
        temp = temp[temp['search_col'].str.contains(sok_hund.lower())]
        temp = temp.drop(columns=['search_col'])
        
    return temp

df_e_filt = filter_dataframe(df_e)
df_j_filt = filter_dataframe(df_j)

# --- 5. VISUALISERING (FLIKAR) ---
tab1, tab2 = st.tabs(["🌲 Eftersök", "🌾 Jaktprov / Fält"])

# === FLIK 1: EFTERSÖK ===
with tab1:
    if df_e_filt.empty:
        st.warning("Ingen eftersöksdata hittades med valda filter.")
    else:
        # Försök hitta poängkolumner
        col_vatten = next((c for c in df_e_filt.columns if "vatten" in c.lower() and "kritik" not in c.lower()), "Vatten")
        col_spar = next((c for c in df_e_filt.columns if "spår" in c.lower() or "spar" in c.lower() and "kritik" not in c.lower()), "Spår")

        # Gör om till siffror
        for col in [col_vatten, col_spar]:
            if col in df_e_filt.columns:
                df_e_filt[col] = pd.to_numeric(df_e_filt[col], errors='coerce').fillna(0)

        # KPI:er
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Antal Starter", len(df_e_filt))
        c2.metric("Unika Hundar", df_e_filt['regnr'].nunique() if 'regnr' in df_e_filt.columns else 0)
        
        # Räkna Godkända (Krav: Minst 4 på både vatten och spår)
        if col_vatten in df_e_filt.columns and col_spar in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt[col_vatten] >= 4) & (df_e_filt[col_spar] >= 4)]
            full_pott = df_e_filt[(df_e_filt[col_vatten] == 10) & (df_e_filt[col_spar] == 10)]
            
            c3.metric("Godkända", f"{len(godkanda)} ({round(len(godkanda)/len(df_e_filt)*100)}%)")
            c4.metric("10-10 (Full pott)", len(full_pott))
        
        st.divider()

        # DIAGRAM
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.markdown("### Vattenbetyg")
            if col_vatten in df_e_filt.columns:
                fig_vatten = px.histogram(df_e_filt, x=col_vatten, nbins=11, 
                                          title="Fördelning Vattenbetyg",
                                          labels={col_vatten: "Betyg"},
                                          color_discrete_sequence=['#3366CC'])
                fig_vatten.update_layout(bargap=0.2)
                st.plotly_chart(fig_vatten, use_container_width=True)

        with col_chart2:
            st.markdown("### Spårbetyg")
            if col_spar in df_e_filt.columns:
                fig_spar = px.histogram(df_e_filt, x=col_spar, nbins=11, 
                                        title="Fördelning Spårbetyg",
                                        labels={col_spar: "Betyg"},
                                        color_discrete_sequence=['#109618'])
                fig_spar.update_layout(bargap=0.2)
                st.plotly_chart(fig_spar, use_container_width=True)

        # TABELL
        st.markdown("### 📋 Resultatlista")
        st.dataframe(
            df_e_filt.sort_values('Datum', ascending=False),
            use_container_width=True,
            hide_index=True
        )

# === FLIK 2: JAKTPROV ===
with tab2:
    if df_j_filt.empty:
        st.warning("Ingen jaktprovsdata hittades med valda filter.")
    else:
        # KPI:er
        c1, c2 = st.columns(2)
        c1.metric("Antal Starter (Fält)", len(df_j_filt))
        c2.metric("Unika Hundar", df_j_filt['regnr'].nunique() if 'regnr' in df_j_filt.columns else 0)
        
        st.divider()
        
        # Försök hitta Pris-kolumn (ofta 'Pris', 'Premie' eller liknande)
        pris_col = next((c for c in df_j_filt.columns if "pris" in c.lower()), None)
        
        if pris_col:
            st.markdown("### Prisfördelning")
            # Räkna antal av varje pris
            pris_counts = df_j_filt[pris_col].value_counts().reset_index()
            pris_counts.columns = ['Pris', 'Antal']
            
            fig_pris = px.pie(pris_counts, values='Antal', names='Pris', 
                              title="Fördelning av Priser",
                              hole=0.4)
            st.plotly_chart(fig_pris, use_container_width=True)
        
        # TABELL
        st.markdown("### 📋 Resultatlista Fält")
        st.dataframe(
            df_j_filt.sort_values('Datum', ascending=False) if 'Datum' in df_j_filt.columns else df_j_filt,
            use_container_width=True,
            hide_index=True
        )

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: df_e_filt.to_excel(writer, index=False, sheet_name='Eftersök Urval')
    if not df_j_filt.empty: df_j_filt.to_excel(writer, index=False, sheet_name='Jaktprov Urval')

st.download_button(
    "📥 Ladda ner urvalet till Excel",
    data=output.getvalue(),
    file_name="Jaktprovsstatistik_Urval.xlsx",
    mime="application/vnd.ms-excel"
)
