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

st.sidebar.markdown("### 🌲 Provresultat")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

with st.sidebar.expander("Ladda upp separata filer (CSV/TXT)"):
    uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"], key="csv_e")
    uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"], key="csv_j")
    uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"], key="txt_prov")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

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

def parse_health_file(file):
    try:
        if file.name.endswith('.xlsx'): return pd.read_excel(file)
        content = file.getvalue()
        try: txt = content.decode("latin-1")
        except: txt = content.decode("utf-8", errors="ignore")
        return pd.read_csv(io.StringIO(txt), sep='\t', on_bad_lines='skip')
    except: return pd.DataFrame()

# --- DATALOAD ---
@st.cache_data
def load_prov_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    status_msg = []

    if excel:
        try:
            xls = pd.ExcelFile(excel)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                df = df.dropna(how='all')
                
                # Enkel logik: Blad1=Fält, Blad2=Eftersök (Eller baserat på kolumner)
                cols = [str(c).lower() for c in df.columns]
                is_eftersok = False
                is_falt = False
                
                if "blad2" in sheet.lower() or any(x in cols for x in ["vatten", "spår"]):
                    is_eftersok = True
                elif "blad1" in sheet.lower() or any(x in cols for x in ["pris", "resultat", "fält"]):
                    is_falt = True
                
                if is_eftersok: df_e_list.append(df)
                elif is_falt: df_j_list.append(df)
                
                status_msg.append(f"Flik '{sheet}': {len(df)} rader -> {'Eftersök' if is_eftersok else 'Fält' if is_falt else 'Ignorerad'}")

        except Exception as e: status_msg.append(f"Excel-fel: {e}")

    # CSV/TXT support
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
df_halsa_raw = parse_health_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()

# --- DIAGNOSRUTA (Visas alltid om data finns) ---
if uploaded_excel and (not df_e_raw.empty or not df_j_raw.empty):
    with st.expander("ℹ️ Inläsningsstatus (Klicka för att se detaljer)", expanded=True):
        for l in logs: st.write(l)
        st.write(f"**Eftersök:** {len(df_e_raw)} rader totalt.")
        st.write(f"**Fält:** {len(df_j_raw)} rader totalt.")
        if not df_e_raw.empty: st.write(f"Kolumner Eftersök: {list(df_e_raw.columns)}")

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
        if 'hund' in cl or 'namn' in cl: col_map['Hund'] = c
        
        # Resultat
        if 'vatten' in cl: col_map['Vatten'] = c
        if 'spår' in cl: col_map['Spår'] = c
        if 'resultat' in cl: col_map['Resultat'] = c
        elif 'pris' in cl: col_map['Resultat'] = c # Mappa Pris till Resultat
        if 'domare' in cl: col_map['Domare'] = c

    for standard, original in col_map.items():
        df[standard] = df[original]

    # Defaults om kolumner saknas
    if 'Ras' not in df.columns: df['Ras'] = "Okänd Ras"
    if 'År' not in df.columns: df['År'] = 2024 # Fallback
    
    # Datumfix
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year.fillna(0).astype(int)
    
    # Könfix
    if 'Kön' in df.columns:
        df['Kön'] = df['Kön'].astype(str).apply(lambda x: 'Hane' if 'h' in x.lower() else ('Tik' if 't' in x.lower() else 'Okänd'))
    else: df['Kön'] = "Okänd"

    # Poäng
    for col in ['Vatten', 'Spår']:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

# Hämta filtervärden (Hantera om listorna är tomma)
years = sorted(list(set(df_e.get('År', [])) | set(df_j.get('År', []))), reverse=True)
if not years: years = [2024]

races = sorted(list(set(df_e.get('Ras', [])) | set(df_j.get('Ras', []))))
if not races: races = ["Ingen ras hittad"]

classes = sorted(list(set(df_e.get('Klass', [])) | set(df_j.get('Klass', []))))
sexes = ["Hane", "Tik", "Okänd"]

valda_raser = st.sidebar.multiselect("Välj Ras", races, default=races[0] if len(races)>0 else None)
valda_ar = st.sidebar.multiselect("Välj År", years, default=years[:1])
valda_klasser = st.sidebar.multiselect("Välj Klass", classes, default=classes)
valda_kon = st.sidebar.multiselect("Välj Kön", sexes, default=sexes)

def apply_filter(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser and 'Ras' in temp.columns: temp = temp[temp['Ras'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    if valda_klasser and 'Klass' in temp.columns: temp = temp[temp['Klass'].isin(valda_klasser)]
    if valda_kon and 'Kön' in temp.columns: temp = temp[temp['Kön'].isin(valda_kon)]
    return temp

df_e_filt = apply_filter(df_e)
df_j_filt = apply_filter(df_j)
df_h_filt = apply_filter(df_halsa_raw) # Enkel filtrering för hälsa

# --- VISNING ---
def prepare_table(df, type="Standard", add_links=True):
    v = df.copy()
    cols = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass']
    if type=="Eftersök": cols += ['Vatten', 'Spår']
    if type=="Fält": cols += ['Resultat', 'Domare']
    
    # Ta bara med kolumner som finns
    final = [c for c in cols if c in v.columns]
    v = v[final].rename(columns={'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn'})
    
    if add_links and 'Reg.nr' in v.columns:
        base = "https://hundar.skk.se/hunddata/Hund_sok.aspx?sok="
        v['Reg.nr'] = v['Reg.nr'].apply(lambda x: f"{base}{x}" if pd.notna(x) else x)
    return v

def get_gender_text(df):
    if 'Kön' not in df.columns: return ""
    h = len(df[df['Kön']=='Hane'])
    t = len(df[df['Kön']=='Tik'])
    return f"({h} H, {t} T)"

# --- TABS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "🏥 Hälsa", "❓ Guide", "ℹ️ Info"])

with tab1:
    if not df_e_filt.empty:
        c1, c2, c3 = st.columns(3)
        starts = len(df_e_filt)
        c1.metric("Starter", f"{starts}", get_gender_text(df_e_filt))
        
        # Godkända
        if 'Vatten' in df_e_filt.columns and 'Spår' in df_e_filt.columns:
            ok = df_e_filt[(df_e_filt['Vatten']>=4) & (df_e_filt['Spår']>=4)]
            full = df_e_filt[(df_e_filt['Vatten']==10) & (df_e_filt['Spår']==10)]
            c2.metric("Godkända (4+)", f"{len(ok)} ({int(len(ok)/starts*100)}%)")
            c3.metric("10-10", len(full))
            
            st.divider()
            cc1, cc2 = st.columns(2)
            with cc1:
                vc = df_e_filt['Vatten'].value_counts().sort_index(ascending=False)
                st.dataframe(vc, use_container_width=True)
            with cc2:
                sc = df_e_filt['Spår'].value_counts().sort_index(ascending=False)
                st.dataframe(sc, use_container_width=True)
        
        st.dataframe(prepare_table(df_e_filt, "Eftersök"), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Eftersök. Kontrollera inläsningsstatus ovan.")

with tab2:
    if not df_j_filt.empty:
        c1, c2 = st.columns(2)
        c1.metric("Starter", len(df_j_filt), get_gender_text(df_j_filt))
        
        if 'Resultat' in df_j_filt.columns:
            st.divider()
            pc = df_j_filt['Resultat'].value_counts()
            c1, c2 = st.columns([1,2])
            with c1: st.dataframe(pc, use_container_width=True)
            with c2: st.plotly_chart(px.pie(values=pc.values, names=pc.index), use_container_width=True)

        st.dataframe(prepare_table(df_j_filt, "Fält"), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Fält. Kontrollera inläsningsstatus ovan.")

with tab3:
    if not df_h_filt.empty:
        res_col = next((c for c in df_h_filt.columns if 'res' in c.lower() or 'dia' in c.lower()), None)
        if res_col:
            rc = df_h_filt[res_col].value_counts()
            st.bar_chart(rc)
        st.dataframe(df_h_filt, use_container_width=True)
    else: st.info("Ladda upp hälsofil.")

with tab4:
    st.markdown("## 📘 Guide")
    st.markdown("1. Ladda upp Excel-filen.\n2. Kontrollera att Ras och År stämmer i filtret.\n3. Läs av tabellerna.")

with tab5:
    st.header("ℹ️ Info")
    st.write("Verktyg för SVK. Ingen data sparas.")

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: prepare_table(df_e_filt, "Eftersök", False).to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: prepare_table(df_j_filt, "Fält", False).to_excel(writer, index=False, sheet_name='Fält')
st.download_button("Ladda ner Excel", output.getvalue(), "SVK_Statistik.xlsx")
