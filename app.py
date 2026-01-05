import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="SVK Avelsstatistik", page_icon="🐕", layout="wide")

st.title("📊 SVK - Statistikverktyg för avel")
st.markdown("Automatiskt underlag för **årssammanställningar** och avelsuppföljning. Verktyget är till för **avelsråd och avelsfunktionärer** inom SVK.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 1. Ladda upp data")

st.sidebar.markdown("### 🌲 Jakt & Eftersök")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

with st.sidebar.expander("Ladda upp separata filer (CSV/TXT)"):
    uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"], key="csv_e")
    uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"], key="csv_j")
    uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"], key="txt_prov")

st.sidebar.markdown("### 🐗 Övriga Prov")
uploaded_fullbruk = st.sidebar.file_uploader("Fullbruksprov (Excel/CSV)", type=["xlsx", "csv"], key="fullbruk")
uploaded_viltspar = st.sidebar.file_uploader("Viltspårprov (Excel/CSV)", type=["xlsx", "csv"], key="viltspar")

st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil från Avelsdata (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- PARSING ---
def parse_txt_to_df(txt_file):
    content = txt_file.getvalue()
    try: text_data = content.decode("latin-1")
    except: text_data = content.decode("utf-8", errors="ignore")
    lines = text_data.split('\n')
    data = []
    for line in lines:
        if re.match(r'^\d{2}-\d{2}-\d{2}', line):
            parts = line.split('\t')
            if len(parts) > 30:
                try:
                    entry = {
                        "Datum": parts[0], "Klass": parts[1].strip(), "Hund": parts[2],
                        "regnr": parts[3], "Ras": parts[4].strip(), "Kön": parts[5].strip(),
                        "Vatten": pd.to_numeric(parts[27], errors='coerce'),
                        "Spår": pd.to_numeric(parts[28], errors='coerce'),
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except: continue
    return pd.DataFrame(data)

def parse_generic_file(file):
    try:
        if file.name.endswith('.xlsx'): return pd.read_excel(file)
        content = file.getvalue()
        try: txt = content.decode("latin-1")
        except: txt = content.decode("utf-8", errors="ignore")
        if ";" in txt.split('\n')[0]: return pd.read_csv(io.StringIO(txt), sep=';', on_bad_lines='skip')
        return pd.read_csv(io.StringIO(txt), sep=',', on_bad_lines='skip')
    except: return pd.DataFrame()

# --- DATALOAD ---
@st.cache_data
def load_prov_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    status_msg = []

    # 1. Excel Master
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
                
                if is_eftersok: 
                    df_e_list.append(df)
                    status_msg.append(f"✅ Flik '{sheet}': Eftersök")
                elif is_falt: 
                    df_j_list.append(df)
                    status_msg.append(f"✅ Flik '{sheet}': Fält")
        except Exception as e: status_msg.append(f"Excel-fel: {e}")

    # 2. Legacy CSV/TXT
    if csv_e: 
        try: df_e_list.append(pd.read_csv(csv_e, sep=';', encoding='latin1'))
        except: pass
    if csv_j:
        try: df_j_list.append(pd.read_csv(csv_j, sep=';', encoding='latin1'))
        except: pass
    if txt:
        df_txt = parse_txt_to_df(txt)
        if not df_txt.empty: df_e_list.append(df_txt)

    df_e = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    
    return df_e, df_j, status_msg

df_e_raw, df_j_raw, logs = load_prov_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)
df_halsa_raw = parse_generic_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()
df_full_raw = parse_generic_file(uploaded_fullbruk) if uploaded_fullbruk else pd.DataFrame()
df_vilt_raw = parse_generic_file(uploaded_viltspar) if uploaded_viltspar else pd.DataFrame()

# --- DIAGNOSRUTA (HOPFÄLLD) ---
if uploaded_excel and (not df_e_raw.empty or not df_j_raw.empty):
    with st.expander("ℹ️ Klicka här för inläsningsstatus", expanded=False):
        for l in logs: st.write(l)

# --- CLEANING ---
def clean_df(df):
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
        
        # HUNDNAMN: Prioritera 'namn'
        if c == 'namn': col_map['Hund'] = c
        elif 'hund' in cl and 'fader' not in cl and 'moder' not in cl and 'namn' not in col_map: col_map['Hund'] = c
        
        # Eftersök
        if 'vatten' in cl and 'passion' not in cl: col_map['Vatten'] = c
        if 'spår' in cl or 'spar' in cl: col_map['Spår'] = c 
        if 'passion' in cl: col_map['Vattenpassion'] = c # Nytt
        
        # Fält & Egenskaper
        if 'resultat' in cl: col_map['Resultat'] = c
        elif 'pris' in cl and 'egenskap' not in cl: col_map['Resultat'] = c 
        if 'domare' in cl: col_map['Domare'] = c
        if 'fart' in cl: col_map['Fart'] = c # Nytt
        if 'vidd' in cl: col_map['Vidd'] = c # Nytt
        if 'reviering' in cl: col_map['Reviering'] = c # Nytt
        if 'följsamhet' in cl or 'foljsamhet' in cl: col_map['Följsamhet'] = c # Nytt

    for standard, original in col_map.items():
        df[standard] = df[original]

    # Defaults
    if 'Ras' not in df.columns: df['Ras'] = "Okänd Ras"
    
    # Datumfix
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year.fillna(0).astype(int)
    else: df['År'] = 2024
    
    # SE-Hund Flagga
    if 'Regnr' in df.columns:
        df['Svensk'] = df['Regnr'].astype(str).str.upper().str.match(r'^SE|^S\d')
    else: df['Svensk'] = False

    # Könfix
    if 'Kön' in df.columns:
        df['Kön'] = df['Kön'].astype(str).apply(lambda x: 'Hane' if 'h' in x.lower() else ('Tik' if 't' in x.lower() else 'Okänd'))
    else: df['Kön'] = "Okänd"
    
    # Klassfix
    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()

    # Poäng och Egenskaper (siffror)
    numeric_cols = ['Vatten', 'Spår', 'Resultat', 'Fart', 'Vidd', 'Reviering', 'Följsamhet', 'Vattenpassion']
    for col in numeric_cols:
        if col in df.columns: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)
df_full = clean_df(df_full_raw)
df_vilt = clean_df(df_vilt_raw)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

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
df_h_filt = apply_filter(df_halsa_raw)

# --- VISNING ---
def get_gender_text(df):
    if 'Kön' not in df.columns: return ""
    h = len(df[df['Kön']=='Hane'])
    t = len(df[df['Kön']=='Tik'])
    return f"({h} Hanar, {t} Tikar)"

def prepare_table(df, add_links=True):
    v = df.copy()
    # Inkludera de nya egenskaperna
    cols = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass', 
            'Resultat', 'Pris', 
            'Vatten', 'Spår', 'Vattenpassion',
            'Fart', 'Vidd', 'Reviering', 'Följsamhet', 'Domare']
    
    final = [c for c in cols if c in v.columns]
    v = v[final].rename(columns={'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn', 'Resultat': 'Pris'})
    
    # Snygga till Pris (ta bort 0)
    if 'Pris' in v.columns:
        v['Pris'] = v['Pris'].apply(lambda x: f"{x}:a Pris" if x > 0 else "-")

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
            c_a, c_b, c_c = st.columns(3)
            with c_a: 
                st.markdown("**Vattenbetyg**")
                st.dataframe(df_e_filt['Vatten'].value_counts().sort_index(ascending=False), use_container_width=True)
            with c_b: 
                st.markdown("**Spårbetyg**")
                st.dataframe(df_e_filt['Spår'].value_counts().sort_index(ascending=False), use_container_width=True)
            with c_c:
                if 'Vattenpassion' in df_e_filt.columns:
                    st.markdown("**Vattenpassion**")
                    st.dataframe(df_e_filt['Vattenpassion'].value_counts().sort_index(ascending=False), use_container_width=True)

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
        c2.metric("Unika Individer", f"{unika_j}")
        
        st.divider()
        
        # --- SPECIALTABELL SE-HUNDAR ---
        st.markdown("### 🇸🇪 Prisfördelning Svenska Hundar (Senaste 2 åren)")
        
        sel_year = int(valda_ar[0]) if valda_ar else 2024
        two_years = [sel_year, sel_year-1]

        # Filtrera: Ras + SE + 2år + Pris>0 (Eftersom 0 inte är ett pris)
        df_se_2y = df_j[
            (df_j['Ras'].isin(valda_raser) if valda_raser else True) & 
            (df_j['Svensk'] == True) & 
            (df_j['År'].isin(two_years)) &
            (df_j['Resultat'] > 0) 
        ]

        if not df_se_2y.empty and 'Resultat' in df_se_2y.columns and 'Klass' in df_se_2y.columns:
            # Pivot med Resultat (1, 2, 3...)
            df_se_2y['Pris'] = df_se_2y['Resultat'].astype(int).astype(str) + ":a Pris"
            pivot = pd.crosstab(df_se_2y['Klass'], df_se_2y['Pris'], margins=True, margins_name="Totalt")
            cols = sorted(pivot.columns.tolist()) # Sortera 1, 2, 3
            st.dataframe(pivot[cols], use_container_width=True)
        else:
            st.warning(f"Inga svenska hundar med pris hittades för åren {two_years}.")

        st.divider()
        
        # --- EGENSKAPSBEDÖMNING ---
        st.markdown("### 📏 Egenskapsbedömning (Genomsnitt)")
        egenskaper = ['Fart', 'Vidd', 'Reviering', 'Följsamhet']
        valid_eg = [e for e in egenskaper if e in df_j_filt.columns]
        
        if valid_eg:
            # Räkna snitt per egenskap (exkludera 0or om det betyder "ej bedömd"?)
            # Ofta är betyg 1-6. 0 betyder ofta ej satt. Vi filtrerar >0.
            stats = {}
            for e in valid_eg:
                avg = df_j_filt[df_j_filt[e] > 0][e].mean()
                stats[e] = round(avg, 2) if pd.notna(avg) else 0
            
            c_eg = st.columns(len(valid_eg))
            for i, (k, v) in enumerate(stats.items()):
                c_eg[i].metric(k, v)
        
        st.divider()
        st.markdown("**Detaljlista (Alla starter)**")
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
            st.markdown("**Prisfördelning (Svenska hundar)**")
            # Filtrera bara svenska hundar och Pris > 0
            res_se = df_full_filt[(df_full_filt['Svensk']) & (df_full_filt['Resultat']>0)]['Resultat'].value_counts()
            if not res_se.empty:
                st.dataframe(res_se, use_container_width=True)
            else:
                st.write("Inga pris tagna av svenska hundar.")
            
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
            st.markdown("**Prisfördelning (Svenska hundar)**")
            res_se = df_vilt_filt[(df_vilt_filt['Svensk']) & (df_vilt_filt['Resultat']>0)]['Resultat'].value_counts()
            if not res_se.empty:
                st.dataframe(res_se, use_container_width=True)
            else: st.write("Inga pris tagna.")

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
    st.markdown("Detta verktyg är framtaget för **avelsråd och avelsfunktionärer inom SVK** för att underlätta arbetet med årssammanställningar.")
    
    st.markdown("""
    ### 📂 Steg 1: Hämta & Ladda upp Data
    
    **A. Provresultat (Jakt & Eftersök)**
    * Denna fil ("Årsstatistik") hämtas från Avelskommitténs **Gemensamma Google Drive**.
    * Om du saknar behörighet, kontakta din sammankallande.
    * Ladda upp Excel-filen under rubriken "Provresultat" i menyn.

    **B. Övriga Prov & Hälsa**
    * Fullbruk och Viltspår laddas upp som separata filer.
    * Hälsodata (HD/ED) hämtas som Excel från **[SKK Avelsdata](https://hundar.skk.se/avelsdata)**.

    ---

    ### 🎯 Steg 2: Filtrera och Avgränsa
    Använd menyn till vänster för att ställa in exakt vad du vill se:
    1. **Ras:** Välj din ras.
    2. **År:** Välj det verksamhetsår du arbetar med (t.ex. 2024).

    ---

    ### 📊 Steg 3: Analys & Sammanställning
    
    **🌲 Eftersök (Tabell 2 & 3 i mallen)**
    * Se rutan "Antal Starter" och "Unika Individer".
    * Använd tabellerna för "Vattenbetyg" och "Spårbetyg".
    * Här hittar du även statistik för **Vattenpassion**.

    **🌾 Fältprov (Tabell 1 i mallen)**
    * Här hittar du specialtabellen **"Prisfördelning Svenska Hundar (Senaste 2 åren)"** som krävs för rapporten.
    * Tabellen visar endast hundar som tagit pris (1, 2, 3).
    * Se även snittvärden för **Fart, Vidd, Reviering och Följsamhet**.

    **🔗 Tips:**
    I detaljlistorna är registreringsnumret en länk. Klicka på det för att komma direkt till hundens sida på SKK Hunddata.
    """)

# === INFO ===
with tab7:
    st.header("ℹ️ Information, Säkerhet & GDPR")
    
    st.markdown("""
    ### 🛡️ SVK Datapolicy
    Detta verktyg tillhandahålls för funktionärer inom Svenska Vorstehklubben (SVK).
    
    * **Syfte:** Effektivisera framtagandet av statistik till årssammanställningar och avelsutvärdering.
    * **Rättslig grund:** Personuppgiftsbehandlingen (namn i resultatlistor) sker med stöd av *berättigat intresse* för föreningens avelsarbete och verksamhetsuppföljning.
    
    ### 🔐 Datasäkerhet
    * **Ingen lagring:** De filer du laddar upp bearbetas endast i serverns tillfälliga arbetsminne.
    * **Automatisk radering:** Så fort du stänger webbläsarfliken eller laddar om sidan raderas all data omedelbart. Inga kopior sparas i någon databas.
    * **Kryptering:** All trafik är krypterad via HTTPS.

    ### 📞 Support
    Vid frågor om verktyget eller datahantering, kontakta Avelskommittén.
    """)

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: prepare_table(df_e_filt, False).to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: prepare_table(df_j_filt, False).to_excel(writer, index=False, sheet_name='Fält')
    if not df_full_filt.empty: prepare_table(df_full_filt, False).to_excel(writer, index=False, sheet_name='Fullbruk')

st.download_button("📥 Ladda ner Excel", output.getvalue(), "SVK_Statistik.xlsx")
