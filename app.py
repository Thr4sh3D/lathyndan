import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="SVK Avelsstatistik", page_icon="🐕", layout="wide")

st.title("📊 SVK - Statistikverktyg för avel")
st.markdown("Automatiskt underlag för **årssammanställningar** och avelsuppföljning.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 1. Ladda upp data")

st.sidebar.markdown("### 🌲 Jakt & Eftersök (Master)")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

st.sidebar.markdown("### 🐗 Övriga Prov")
uploaded_fullbruk = st.sidebar.file_uploader("Fullbruksprov (Excel/CSV)", type=["xlsx", "csv"], key="fullbruk")
uploaded_viltspar = st.sidebar.file_uploader("Viltspårprov (Excel/CSV)", type=["xlsx", "csv"], key="viltspar")

st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil från Avelsdata (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- PARSING FUNKTIONER ---
def parse_generic_file(file):
    """Läser in Excel eller CSV oavsett format."""
    try:
        if file.name.endswith('.xlsx'):
            return pd.read_excel(file)
        else:
            # Försök med olika separatorer för CSV
            content = file.getvalue()
            try: txt = content.decode("latin-1")
            except: txt = content.decode("utf-8", errors="ignore")
            
            # Testa semikolon först (vanligast i Sverige)
            if ";" in txt.split('\n')[0]:
                return pd.read_csv(io.StringIO(txt), sep=';', on_bad_lines='skip')
            else:
                return pd.read_csv(io.StringIO(txt), sep=',', on_bad_lines='skip')
    except: return pd.DataFrame()

# --- DATALOAD ---
@st.cache_data
def load_prov_data(excel):
    df_e_list = []
    df_j_list = []
    status_msg = []

    if excel:
        try:
            xls = pd.ExcelFile(excel)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                df = df.dropna(how='all')
                
                cols = [str(c).lower() for c in df.columns]
                is_eftersok = False
                is_falt = False
                
                # Identifiering
                if "blad2" in sheet.lower() or any(x in cols for x in ["vatten", "spår"]):
                    is_eftersok = True
                elif "blad1" in sheet.lower() or any(x in cols for x in ["pris", "resultat", "fält"]):
                    is_falt = True
                
                if is_eftersok: df_e_list.append(df)
                elif is_falt: df_j_list.append(df)
                
        except Exception as e: status_msg.append(f"Fel: {e}")

    df_e = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    
    return df_e, df_j

df_e_raw, df_j_raw = load_prov_data(uploaded_excel)
df_halsa_raw = parse_generic_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()
df_full_raw = parse_generic_file(uploaded_fullbruk) if uploaded_fullbruk else pd.DataFrame()
df_vilt_raw = parse_generic_file(uploaded_viltspar) if uploaded_viltspar else pd.DataFrame()

# --- CLEANING ---
def clean_df(df, type="Standard"):
    if df.empty: return df
    df.columns = df.columns.str.strip()
    
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'datum' in cl: col_map['Datum'] = c
        if 'klass' in cl: col_map['Klass'] = c
        if 'ras' in cl: col_map['Ras'] = c
        if 'regnr' in cl or 'reg.nr' in cl: col_map['Regnr'] = c
        if 'kön' in cl or c == 'S': col_map['Kön'] = c
        if 'hund' in cl or 'namn' in cl: col_map['Hund'] = c
        
        # Specifika
        if 'vatten' in cl: col_map['Vatten'] = c
        if 'spår' in cl: col_map['Spår'] = c
        if 'resultat' in cl: col_map['Resultat'] = c
        elif 'pris' in cl: col_map['Resultat'] = c 
        if 'domare' in cl: col_map['Domare'] = c

    for standard, original in col_map.items():
        df[standard] = df[original]

    # Defaults
    if 'Ras' not in df.columns: df['Ras'] = "Okänd Ras"
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year.fillna(0).astype(int)
    else:
        df['År'] = 2024 # Fallback

    # SE-Hund Flagga (Svenskregistrerad)
    if 'Regnr' in df.columns:
        df['Svensk'] = df['Regnr'].astype(str).str.upper().str.match(r'^SE|^S\d')
    else:
        df['Svensk'] = False

    # Könfix
    if 'Kön' in df.columns:
        df['Kön'] = df['Kön'].astype(str).apply(lambda x: 'Hane' if 'h' in x.lower() else ('Tik' if 't' in x.lower() else 'Okänd'))
    
    # Klassfix
    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()

    # Poäng (bara för prov)
    for col in ['Vatten', 'Spår']:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)
df_full = clean_df(df_full_raw)
df_vilt = clean_df(df_vilt_raw)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

# Samla år från alla källor
all_years_set = set(df_e.get('År', [])) | set(df_j.get('År', [])) | set(df_full.get('År', [])) | set(df_vilt.get('År', []))
years = sorted(list(all_years_set), reverse=True)
if not years: years = [2024]

races = sorted(list(set(df_e.get('Ras', [])) | set(df_j.get('Ras', []))))
if not races: races = ["Ingen ras hittad"]

valda_raser = st.sidebar.multiselect("Välj Ras", races, default=races[0] if len(races)>0 else None)
valda_ar = st.sidebar.multiselect("Välj År (Verksamhetsår)", years, default=years[:1])

def apply_filter(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser and 'Ras' in temp.columns: temp = temp[temp['Ras'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    return temp

df_e_filt = apply_filter(df_e)
df_j_filt = apply_filter(df_j)
df_full_filt = apply_filter(df_full)
df_vilt_filt = apply_filter(df_vilt)
df_h_filt = apply_filter(df_halsa_raw) # Enkel filtrering för hälsa

# --- HJÄLPFUNKTIONER ---
def get_gender_text(df):
    if 'Kön' not in df.columns: return ""
    h = len(df[df['Kön']=='Hane'])
    t = len(df[df['Kön']=='Tik'])
    return f"({h} Hanar, {t} Tikar)"

def prepare_table(df, add_links=True):
    v = df.copy()
    cols = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass', 'Resultat', 'Pris', 'Vatten', 'Spår', 'Domare']
    final = [c for c in cols if c in v.columns]
    v = v[final].rename(columns={'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn'})
    
    if add_links and 'Reg.nr' in v.columns:
        base = "https://hundar.skk.se/hunddata/Hund_sok.aspx?sok="
        v['Reg.nr'] = v['Reg.nr'].apply(lambda x: f"{base}{x}" if pd.notna(x) else x)
    return v

# --- TABS ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "🐕 Fullbruk", "🩸 Viltspår", "🏥 Hälsa", "❓ Guide", "ℹ️ Info"])

# === EFTERSÖK ===
with tab1:
    if not df_e_filt.empty:
        st.subheader("Eftersök (Vatten & Spår)")
        c1, c2, c3 = st.columns(3)
        starts = len(df_e_filt)
        unika = df_e_filt['Regnr'].nunique() if 'Regnr' in df_e_filt.columns else 0
        
        c1.metric("Antal Starter", f"{starts}", get_gender_text(df_e_filt))
        c2.metric("Unika Individer", f"{unika}", help="Antal unika hundar som startat minst en gång")
        
        if 'Vatten' in df_e_filt.columns and 'Spår' in df_e_filt.columns:
            ok = df_e_filt[(df_e_filt['Vatten']>=4) & (df_e_filt['Spår']>=4)]
            full = df_e_filt[(df_e_filt['Vatten']==10) & (df_e_filt['Spår']==10)]
            c3.metric("Godkända (4+)", f"{len(ok)} ({int(len(ok)/starts*100)}%)")
            
            st.divider()
            c_a, c_b = st.columns(2)
            with c_a: 
                st.markdown("**Vattenbetyg**")
                st.dataframe(df_e_filt['Vatten'].value_counts().sort_index(ascending=False), use_container_width=True)
            with c_b: 
                st.markdown("**Spårbetyg**")
                st.dataframe(df_e_filt['Spår'].value_counts().sort_index(ascending=False), use_container_width=True)

        st.dataframe(prepare_table(df_e_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Eftersök.")

# === FÄLTPROV ===
with tab2:
    if not df_j_filt.empty:
        st.subheader("Jaktprov / Fält")
        c1, c2 = st.columns(2)
        unika_j = df_j_filt['Regnr'].nunique() if 'Regnr' in df_j_filt.columns else 0
        
        c1.metric("Antal Starter", len(df_j_filt), get_gender_text(df_j_filt))
        c2.metric("Unika Individer", f"{unika_j}", help="Antal unika hundar som startat minst en gång")
        
        st.divider()
        
        # --- SPECIALTABELL: PRISFÖRDELNING SE-HUNDAR (2 ÅR) ---
        st.markdown("### 🇸🇪 Prisfördelning Svenska Hundar (Senaste 2 åren)")
        
        # Hitta de två senaste åren i datasetet (inte bara filtret)
        all_avail_years = sorted(list(set(df_j['År'])), reverse=True)
        if len(all_avail_years) >= 2:
            two_years = all_avail_years[:2]
        else:
            two_years = all_avail_years

        # Filtrera: Rätt Ras, SE-hundar, Senaste 2 åren
        df_se_2y = df_j[
            (df_j['Ras'].isin(valda_raser) if valda_raser else True) & 
            (df_j['Svensk'] == True) & 
            (df_j['År'].isin(two_years))
        ]

        if not df_se_2y.empty and 'Resultat' in df_se_2y.columns and 'Klass' in df_se_2y.columns:
            st.caption(f"Visar data för år: {two_years}")
            
            # Skapa pivottabell: Klass x Pris
            pivot = pd.crosstab(df_se_2y['Klass'], df_se_2y['Resultat'], margins=True, margins_name="Totalt")
            st.dataframe(pivot, use_container_width=True)
        else:
            st.warning("Hittade inte tillräckligt med data för SE-hundar de senaste två åren.")

        st.divider()
        st.markdown("**Alla resultat (Valt år)**")
        st.dataframe(prepare_table(df_j_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Fält.")

# === FULLBRUK ===
with tab3:
    st.subheader("🐕 Fullbruksprov")
    if not df_full_filt.empty:
        c1, c2 = st.columns(2)
        c1.metric("Starter", len(df_full_filt), get_gender_text(df_full_filt))
        c2.metric("Unika SE-hundar", df_full_filt[df_full_filt['Svensk']]['Regnr'].nunique())
        
        st.divider()
        if 'Resultat' in df_full_filt.columns:
            st.markdown("**Resultatfördelning (Svenska hundar)**")
            res_se = df_full_filt[df_full_filt['Svensk']]['Resultat'].value_counts()
            st.dataframe(res_se, use_container_width=True)
            
        st.dataframe(prepare_table(df_full_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ladda upp Fullbruksprov-fil i menyn.")

# === VILTSPÅR ===
with tab4:
    st.subheader("🩸 Viltspårprov")
    if not df_vilt_filt.empty:
        c1, c2 = st.columns(2)
        c1.metric("Starter", len(df_vilt_filt), get_gender_text(df_vilt_filt))
        c2.metric("Unika SE-hundar", df_vilt_filt[df_vilt_filt['Svensk']]['Regnr'].nunique())
        
        st.divider()
        if 'Resultat' in df_vilt_filt.columns:
            st.markdown("**Resultatfördelning (Svenska hundar)**")
            res_se = df_vilt_filt[df_vilt_filt['Svensk']]['Resultat'].value_counts()
            st.dataframe(res_se, use_container_width=True)

        st.dataframe(prepare_table(df_vilt_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ladda upp Viltspår-fil i menyn.")

# === HÄLSA ===
with tab5:
    st.subheader("🏥 Hälsostatistik")
    if not df_h_filt.empty:
        res_col = next((c for c in df_h_filt.columns if 'res' in c.lower() or 'dia' in c.lower()), None)
        if res_col:
            st.markdown(f"**Statistik: {res_col}**")
            rc = df_h_filt[res_col].value_counts().rename("Antal")
            c1, c2 = st.columns(2)
            with c1: st.dataframe(rc, use_container_width=True)
            with c2: st.plotly_chart(px.bar(x=rc.index, y=rc.values), use_container_width=True)
        st.dataframe(df_h_filt, use_container_width=True)
    else: st.info("Ladda upp hälsofil i menyn.")

# === GUIDE ===
with tab6:
    st.markdown("## 📘 Användarguide (SVK)")
    st.markdown("### 1. Hämta Data")
    st.markdown("""
    * **Jakt & Eftersök:** Ladda ner "Årsstatistik"-filen från er Drive.
    * **Fullbruk/Viltspår:** Dessa ligger ofta i separata filer.
    * **Hälsa:** Hämta Excel från [SKK Avelsdata](https://hundar.skk.se/avelsdata).
    """)
    st.markdown("### 2. Statistik till Årssammanställning")
    st.markdown("""
    * **Unika Individer:** Se mätaren "Unika Individer" högst upp på varje flik.
    * **Svenskregistrerade:** Appen filtrerar automatiskt fram "SE-hundar" för specialtabellerna (t.ex. prisfördelning 2 år).
    * **Prisfördelning:** Finns under fliken Fältprov.
    """)

# === INFO ===
with tab7:
    st.header("ℹ️ Policy & Säkerhet")
    st.markdown("""
    * **Syfte:** Verktyg för SVK:s avelsfunktionärer.
    * **GDPR:** Personuppgifter (namn i resultatlistor) behandlas med stöd av *berättigat intresse* för avelsuppföljning.
    * **Ingen lagring:** Data raderas vid stängning.
    """)

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: prepare_table(df_e_filt, False).to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: prepare_table(df_j_filt, False).to_excel(writer, index=False, sheet_name='Fält')
    if not df_full_filt.empty: prepare_table(df_full_filt, False).to_excel(writer, index=False, sheet_name='Fullbruk')

st.download_button("📥 Ladda ner Excel", output.getvalue(), "SVK_Statistik.xlsx")
