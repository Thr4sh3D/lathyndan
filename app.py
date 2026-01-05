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

# Provresultat
st.sidebar.markdown("---")
st.sidebar.markdown("### 🌲 Provresultat")
st.sidebar.info("Ladda upp Master-filen (Excel) med Blad 1 & 2.")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

# Fallback för gamla CSV/TXT
with st.sidebar.expander("Ladda upp separata filer (Gamla formatet)"):
    uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"], key="csv_e")
    uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"], key="csv_j")
    uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"], key="txt_prov")

# Hälsodata
st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil från Avelsdata (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- PARSING FUNKTIONER ---

def identify_sheet_type(df):
    """
    Analyserar kolumnerna för att avgöra om det är Eftersök eller Fält.
    """
    cols = [c.lower() for c in df.columns]
    
    # Eftersök har ofta "vatten" och "spår"
    if any("vatten" in c for c in cols) and any("spår" in c for c in cols):
        return "Eftersök"
    
    # Fält har ofta "fältbetyg", "egenskapspar" eller specifika priser
    if any("fält" in c for c in cols) or any("resultat" in c for c in cols):
        return "Fält"
        
    return "Okänd"

def parse_txt_to_df(txt_file):
    # (Behålls för bakåtkompatibilitet med gamla textfiler)
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
                    raw_kon = parts[5].strip() if len(parts) > 5 else "Okänd"
                    entry = {
                        "Datum": parts[0],
                        "Klass": parts[1].strip(),
                        "Hund": parts[2],
                        "Regnr": parts[3],
                        "Ras": parts[4].strip(),
                        "Kön": raw_kon,
                        "Vatten": pd.to_numeric(parts[27], errors='coerce'),
                        "Spår": pd.to_numeric(parts[28], errors='coerce'),
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except: continue
    return pd.DataFrame(data)

def parse_health_file(file):
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file)
        else:
            content = file.getvalue()
            try: txt = content.decode("latin-1")
            except: txt = content.decode("utf-8", errors="ignore")
            df = pd.read_csv(io.StringIO(txt), sep='\t', on_bad_lines='skip')
        return df
    except: return pd.DataFrame()

@st.cache_data
def load_prov_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    
    # 1. NYA MASTER-EXCELEN (Smart Inläsning)
    if excel:
        try:
            xls = pd.ExcelFile(excel)
            # Läs alla flikar och sortera dem
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                typ = identify_sheet_type(df)
                
                if typ == "Eftersök":
                    df_e_list.append(df)
                elif typ == "Fält":
                    df_j_list.append(df)
        except Exception as e:
            st.error(f"Kunde inte läsa Excel-filen: {e}")

    # 2. GAMLA CSV/TXT (Fallback)
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
    if txt:
        df_txt = parse_txt_to_df(txt)
        if not df_txt.empty: df_e_list.append(df_txt)

    df_e_tot = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j_tot = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    return df_e_tot, df_j_tot

df_e_raw, df_j_raw = load_prov_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)
df_halsa_raw = parse_health_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()

# --- DATATVÄTT (Standardisering) ---
def clean_df(df):
    if df.empty: return df
    
    # 1. Standardisera kolumnnamn (Ta bort mellanslag, gör gemener för matchning)
    # Vi sparar originalnamnen men skapar en "clean" map
    df.columns = df.columns.str.strip()
    
    # Mappning för att hitta rätt kolumn oavsett vad den heter i filen
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'datum' in cl and 'röntgen' not in cl: col_map['Datum'] = c
        if 'klass' in cl: col_map['Klass'] = c
        if 'ras' in cl and 'namn' not in cl: col_map['Ras'] = c # Undvik 'Rasnamn' om 'Ras' finns
        if 'regnr' in cl or 'reg.nr' in cl: col_map['Regnr'] = c
        if 'kön' in cl: col_map['Kön'] = c
        if 'hund' in cl and 'namn' in cl: col_map['Hund'] = c
        elif 'hund' in cl and 'fader' not in cl and 'moder' not in cl: col_map['Hund'] = c
        
        # Poäng Eftersök
        if 'vatten' in cl and 'poäng' not in cl: col_map['Vatten'] = c
        if 'spår' in cl and 'poäng' not in cl: col_map['Spår'] = c
        
        # Resultat Fält
        if 'resultat' in cl: col_map['Resultat'] = c
        if 'fältbetyg' in cl: col_map['Fältbetyg'] = c
        if 'domare' in cl: col_map['Domare'] = c

    # Skapa standardiserade kolumner
    for standard, original in col_map.items():
        df[standard] = df[original]

    # DATUM & ÅR
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year
    
    # KLASS
    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()

    # KÖN (Standardisering Hane/Tik)
    if 'Kön' in df.columns:
        def fix_kon(v):
            v = str(v).lower()
            if 'h' in v: return 'Hane'
            if 't' in v: return 'Tik'
            return 'Okänd'
        df['Kön'] = df['Kön'].apply(fix_kon)
    else:
        df['Kön'] = "Okänd"

    # POÄNG (Tvinga till siffror)
    if 'Vatten' in df.columns:
        df['Vatten'] = pd.to_numeric(df['Vatten'], errors='coerce').fillna(0)
    if 'Spår' in df.columns:
        df['Spår'] = pd.to_numeric(df['Spår'], errors='coerce').fillna(0)

    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

years_prov = set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))
# Hälsa år
if not df_halsa_raw.empty:
    date_col = next((c for c in df_halsa_raw.columns if 'datum' in c.lower()), None)
    if date_col:
        df_halsa_raw[date_col] = pd.to_datetime(df_halsa_raw[date_col], errors='coerce')
        df_halsa_raw['År'] = df_halsa_raw[date_col].dt.year
        years_prov = years_prov | set(df_halsa_raw['År'].dropna().astype(int))

all_years = sorted(list(years_prov), reverse=True)
all_races = sorted(list(set(df_e.get('Ras', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras', pd.Series()).dropna().astype(str))))
all_classes = sorted(list(set(df_e.get('Klass', pd.Series()).dropna().astype(str)) | set(df_j.get('Klass', pd.Series()).dropna().astype(str))))
all_sex = ["Hane", "Tik"]

if not all_races: all_races = ["Ladda fil först"]

valda_raser = st.sidebar.multiselect("Välj Ras", all_races, default=all_races[:1] if all_races else None)
valda_ar = st.sidebar.multiselect("Välj År", all_years, default=all_years[:1] if all_years else None)
valda_klasser = st.sidebar.multiselect("Välj Klass (Prov)", all_classes, default=all_classes)
valda_kon = st.sidebar.multiselect("Välj Kön (Prov)", all_sex, default=all_sex)

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
df_h_filt = apply_filter(df_halsa_raw)

# --- VISNINGSFUNKTION ---
def prepare_display_table(df, type="Standard", add_links=False):
    visning = df.copy()
    
    # Välj kolumner baserat på typ
    cols_to_show = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass']
    
    if type == "Eftersök":
        cols_to_show += ['Vatten', 'Spår', 'Domare']
    elif type == "Fält":
        # Lägg till specifika fältkolumner om de finns
        if 'Resultat' in visning.columns: cols_to_show.append('Resultat')
        if 'Fältbetyg' in visning.columns: cols_to_show.append('Fältbetyg')
        if 'Domare' in visning.columns: cols_to_show.append('Domare')

    # Filtrera kolumner som faktiskt finns
    final_cols = [c for c in cols_to_show if c in visning.columns]
    visning = visning[final_cols]
    
    # Byt namn till snygga rubriker
    rename_map = {
        'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn'
    }
    visning = visning.rename(columns=rename_map)

    # Länkar
    if add_links and 'Reg.nr' in visning.columns:
        base_url = "https://hundar.skk.se/hunddata/Hund_sok.aspx?sok="
        visning['Reg.nr'] = visning['Reg.nr'].apply(lambda x: f"{base_url}{x}" if pd.notna(x) else x)
        
    return visning

def get_gender_breakdown(df):
    if 'Kön' not in df.columns: return ""
    hanar = len(df[df['Kön'] == 'Hane'])
    tikar = len(df[df['Kön'] == 'Tik'])
    return f"({hanar} Hanar, {tikar} Tikar)"

# --- FLIKAR ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "🏥 Hälsa (HD/ED)", "❓ Guide", "ℹ️ Info & GDPR"])

# === FLIK 1: EFTERSÖK ===
with tab1:
    if not df_e_filt.empty:
        st.subheader("Eftersök (Vatten & Spår)")
        
        c1, c2, c3, c4 = st.columns(4)
        num_starts = len(df_e_filt)
        c1.metric("Antal Starter", f"{num_starts}", delta=get_gender_breakdown(df_e_filt), delta_color="off")
        c2.metric("Unika Hundar", df_e_filt['Regnr'].nunique() if 'Regnr' in df_e_filt.columns else 0)

        if 'Vatten' in df_e_filt.columns and 'Spår' in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt['Vatten'] >= 4) & (df_e_filt['Spår'] >= 4)]
            full = df_e_filt[(df_e_filt['Vatten'] == 10) & (df_e_filt['Spår'] == 10)]
            c3.metric("Godkända (4-4+)", f"{len(godkanda)} ({round(len(godkanda)/num_starts*100, 1)}%)")
            c4.metric("Full pott (10-10)", f"{len(full)}")

            st.divider()
            col_v1, col_v2 = st.columns([1, 2])
            with col_v1:
                st.markdown("##### 💧 Vattenbetyg")
                v_counts = df_e_filt['Vatten'].value_counts().sort_index(ascending=False).rename("Antal")
                st.dataframe(v_counts, use_container_width=True)
            with col_v2:
                fig_v = px.bar(x=v_counts.index, y=v_counts.values, labels={'x': 'Betyg', 'y': 'Antal'}, color_discrete_sequence=['#3366CC'])
                fig_v.update_layout(xaxis=dict(tickmode='linear', dtick=1), showlegend=False)
                st.plotly_chart(fig_v, use_container_width=True, key="v_chart")

            st.divider()
            col_s1, col_s2 = st.columns([1, 2])
            with col_s1:
                st.markdown("##### 🌲 Spårbetyg")
                s_counts = df_e_filt['Spår'].value_counts().sort_index(ascending=False).rename("Antal")
                st.dataframe(s_counts, use_container_width=True)
            with col_s2:
                fig_s = px.bar(x=s_counts.index, y=s_counts.values, labels={'x': 'Betyg', 'y': 'Antal'}, color_discrete_sequence=['#109618'])
                fig_s.update_layout(xaxis=dict(tickmode='linear', dtick=1), showlegend=False)
                st.plotly_chart(fig_s, use_container_width=True, key="s_chart")
            
        st.markdown("---")
        df_visning_e = prepare_display_table(df_e_filt, type="Eftersök", add_links=True)
        st.dataframe(
            df_visning_e, use_container_width=True, hide_index=True,
            column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")}
        )
    else: st.info("Ladda upp data och välj filter för att se statistik.")

# === FLIK 2: FÄLT ===
with tab2:
    if not df_j_filt.empty:
        st.subheader("Jaktprov / Fält")
        c1, c2 = st.columns(2)
        num_starts_j = len(df_j_filt)
        c1.metric("Antal Starter", f"{num_starts_j}", delta=get_gender_breakdown(df_j_filt), delta_color="off")
        c2.metric("Unika Hundar", df_j_filt['Regnr'].nunique() if 'Regnr' in df_j_filt.columns else 0)
        st.divider()

        # Prisfördelning (Baserat på Resultat eller Fältbetyg)
        pris_col = 'Resultat' if 'Resultat' in df_j_filt.columns else None
        
        if pris_col:
            st.markdown("##### 🏆 Prisfördelning")
            col_p1, col_p2 = st.columns([1, 2])
            p_counts = df_j_filt[pris_col].value_counts().rename("Antal")
            with col_p1: st.dataframe(p_counts, use_container_width=True)
            with col_p2: st.plotly_chart(px.pie(values=p_counts.values, names=p_counts.index, hole=0.4), use_container_width=True)
        
        st.markdown("---")
        df_visning_j = prepare_display_table(df_j_filt, type="Fält", add_links=True)
        st.dataframe(
            df_visning_j, use_container_width=True, hide_index=True,
            column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")}
        )
    else: st.info("Ladda upp data och välj filter för att se statistik.")

# === FLIK 3: HÄLSA ===
with tab3:
    st.subheader("🏥 Hälsostatistik")
    if df_h_filt.empty:
        st.info("Ingen hälsodata laddad. Ladda upp Excel från SKK Avelsdata.")
    else:
        # Försök hitta resultatkolumn
        cols = df_h_filt.columns
        res_col = next((c for c in cols if 'resultat' in c.lower() or 'diagnos' in c.lower()), None)
        
        if res_col:
            st.write(f"Statistik för: **{res_col}**")
            c1, c2 = st.columns(2)
            res_counts = df_h_filt[res_col].value_counts().rename("Antal")
            with c1: st.dataframe(res_counts, use_container_width=True)
            with c2: st.plotly_chart(px.bar(x=res_counts.index, y=res_counts.values, labels={'x': 'Resultat', 'y': 'Antal'}), use_container_width=True)
            st.markdown("---")
            st.dataframe(df_h_filt, use_container_width=True)
        else:
            st.warning("Hittade inte resultat-kolumnen. Visar listan:")
            st.dataframe(df_h_filt)

# === FLIK 4: GUIDE ===
with tab4:
    st.markdown("## 📘 Användarguide (SVK)")
    st.markdown("Verktyget används av **avelsråd** och **funktionärer** inom SVK för årssammanställningar.")
    
    st.markdown("""
    ### 📂 Steg 1: Hämta & Ladda upp
    1. **Provresultat:** Ladda ner den samlade "Årsstatistik"-filen (Excel) från den gemensamma driven.
    2. **Hälsodata:** Ladda ner Excel-listor från **[SKK Avelsdata](https://hundar.skk.se/avelsdata)**.
    3. **Ladda upp:** Använd menyn till vänster i detta verktyg.

    ### 🎯 Steg 2: Filtrera
    * **Ras:** Välj din ras (t.ex. Korthårig Vorsteh, KLM).
    * **År:** Välj verksamhetsår (t.ex. 2024).
    * **Klass:** Välj klass (UKL, ÖKL, EKL) för att få rätt underlag till rapporten.

    ### 📊 Steg 3: Analysera
    Använd flikarna ovan för att se:
    * **Eftersök:** Antal starter, könsuppdelning, och betygsfördelning (Vatten/Spår).
    * **Fält:** Prisfördelning.
    * **Hälsa:** HD/ED-statistik.
    """)

# === FLIK 5: INFO ===
with tab5:
    st.header("ℹ️ Information")
    st.markdown("Detta verktyg hanterar data temporärt för statistiska ändamål. Ingen data sparas permanent.")

# --- EXPORT ---
st.divider()
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: 
        prepare_display_table(df_e_filt, type="Eftersök", add_links=False).to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: 
        prepare_display_table(df_j_filt, type="Fält", add_links=False).to_excel(writer, index=False, sheet_name='Fält')
    if not df_h_filt.empty:
        df_h_filt.to_excel(writer, index=False, sheet_name='Hälsa')

st.download_button("📥 Ladda ner Urval (Excel)", output.getvalue(), "SVK_Statistik_Urval.xlsx", "application/vnd.ms-excel")
