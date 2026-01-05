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
st.sidebar.info("Ladda upp Master-filen (Excel) eller separata filer.")

uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

with st.sidebar.expander("Ladda upp separata filer (CSV/TXT)"):
    uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"], key="csv_e")
    uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"], key="csv_j")
    uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"], key="txt_prov")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- PARSING FUNKTIONER ---

def parse_txt_to_df(txt_file):
    # Gammal logik för textfiler (Lathunden)
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
                        "regnr": parts[3],
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
        if file.name.endswith('.xlsx'): return pd.read_excel(file)
        content = file.getvalue()
        try: txt = content.decode("latin-1")
        except: txt = content.decode("utf-8", errors="ignore")
        return pd.read_csv(io.StringIO(txt), sep='\t', on_bad_lines='skip')
    except: return pd.DataFrame()

# --- HUVUDINLÄSNING ---
@st.cache_data
def load_prov_data(excel, csv_e, csv_j, txt):
    df_e_list = []
    df_j_list = []
    log_msg = []

    # 1. EXCEL (Master-fil)
    if excel:
        try:
            xls = pd.ExcelFile(excel)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                df = df.dropna(how='all') # Rensa tomma rader
                cols = [str(c).lower() for c in df.columns]
                
                # LOGIK: Identifiera vad fliken innehåller
                is_eftersok = False
                is_falt = False
                
                # A. Namn-baserad identifiering (Starkast)
                if "blad2" in sheet.lower() or "eftersök" in sheet.lower():
                    is_eftersok = True
                elif "blad1" in sheet.lower() or "jakt" in sheet.lower() or "fält" in sheet.lower():
                    is_falt = True
                
                # B. Innehålls-baserad (Om namnet inte matchar)
                if not is_eftersok and not is_falt:
                    if any("vatten" in c for c in cols) and any("spår" in c for c in cols):
                        is_eftersok = True
                    elif any("pris" in c for c in cols) or any("resultat" in c for c in cols):
                        is_falt = True

                # Lägg till i rätt lista
                if is_eftersok:
                    df_e_list.append(df)
                    log_msg.append(f"Läste in '{sheet}' som **Eftersök** ({len(df)} rader)")
                elif is_falt:
                    df_j_list.append(df)
                    log_msg.append(f"Läste in '{sheet}' som **Fält/Jakt** ({len(df)} rader)")
                else:
                    log_msg.append(f"Ignorerade fliken '{sheet}' (Kunde inte identifiera typ)")

        except Exception as e:
            log_msg.append(f"Fel vid Excel-läsning: {e}")

    # 2. CSV/TXT (Legacy support)
    if csv_e:
        try:
            df = pd.read_csv(csv_e, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_e, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_e_list.append(df)
            log_msg.append("Läste in CSV (Eftersök)")
        except: pass
    
    if csv_j:
        try:
            df = pd.read_csv(csv_j, sep=';', encoding='latin1', on_bad_lines='skip')
            if len(df.columns) < 2: df = pd.read_csv(csv_j, sep=',', encoding='utf-8', on_bad_lines='skip')
            df_j_list.append(df)
            log_msg.append("Läste in CSV (Fält)")
        except: pass

    if txt:
        df_txt = parse_txt_to_df(txt)
        if not df_txt.empty:
            df_e_list.append(df_txt)
            log_msg.append("Läste in Textfil (Eftersök)")

    # Slå ihop allt
    df_e_tot = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j_tot = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    
    return df_e_tot, df_j_tot, log_msg

df_e_raw, df_j_raw, debug_logs = load_prov_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)
df_halsa_raw = parse_health_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()

# --- DATATVÄTT (Standardisering) ---
def clean_df(df):
    if df.empty: return df
    
    # 1. Rensa kolumnnamn
    df.columns = df.columns.str.strip()
    
    # 2. Mappa till standardnamn
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'datum' in cl and 'födelse' not in cl: col_map['Datum'] = c
        if 'klass' in cl: col_map['Klass'] = c
        if 'ras' in cl and 'namn' not in cl: col_map['Ras'] = c
        elif 'rasnamn' in cl: col_map['Ras'] = c # Fallback
        if 'regnr' in cl or 'reg.nr' in cl: col_map['Regnr'] = c
        
        # Kön
        if 'kön' in cl: col_map['Kön'] = c
        elif c == 'S': col_map['Kön'] = c # Ibland heter kön "S" (Sex)
        
        # Hund
        if 'hund' in cl and 'namn' in cl: col_map['Hund'] = c
        elif 'hund' in cl and 'fader' not in cl and 'moder' not in cl: col_map['Hund'] = c
        elif 'namn' in cl and 'ras' not in cl and 'ägare' not in cl and 'domare' not in cl: col_map['Hund'] = c
        
        # Eftersök
        if 'vatten' in cl and 'poäng' not in cl: col_map['Vatten'] = c
        if 'spår' in cl and 'poäng' not in cl: col_map['Spår'] = c
        
        # Fält
        if 'resultat' in cl: col_map['Resultat'] = c
        if 'pris' in cl and 'egenskap' not in cl: col_map['Pris'] = c
        if 'betyg' in cl: col_map['Fältbetyg'] = c
        if 'domare' in cl: col_map['Domare'] = c
        if 'kritik' in cl: col_map['Kritik'] = c

    # Applicera mappning
    for standard, original in col_map.items():
        df[standard] = df[original]

    # Om Resultat saknas men Pris finns -> Använd Pris som Resultat
    if 'Pris' in df.columns and 'Resultat' not in df.columns:
        df['Resultat'] = df['Pris']

    # 3. Typkonvertering
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year
    
    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()

    if 'Kön' in df.columns:
        def fix_kon(v):
            v = str(v).lower()
            if 'h' in v: return 'Hane'
            if 't' in v: return 'Tik'
            return 'Okänd'
        df['Kön'] = df['Kön'].apply(fix_kon)
    else:
        df['Kön'] = "Okänd"

    # Poäng till siffror
    for col in ['Vatten', 'Spår']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

years_prov = set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))
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

if not all_races: all_races = ["(Ingen data)"]

valda_raser = st.sidebar.multiselect("Välj Ras", all_races, default=all_races[0] if all_races else None)
valda_ar = st.sidebar.multiselect("Välj År", all_years, default=all_years[:1] if all_years else None)
valda_klasser = st.sidebar.multiselect("Välj Klass", all_classes, default=all_classes)
valda_kon = st.sidebar.multiselect("Välj Kön", all_sex, default=all_sex)

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
    cols_to_show = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass']
    
    if type == "Eftersök":
        cols_to_show += ['Vatten', 'Spår', 'Domare']
    elif type == "Fält":
        # Prioritera Resultat
        if 'Resultat' in visning.columns: cols_to_show.append('Resultat')
        if 'Fältbetyg' in visning.columns: cols_to_show.append('Fältbetyg')
        if 'Domare' in visning.columns: cols_to_show.append('Domare')
        if 'Kritik' in visning.columns: cols_to_show.append('Kritik')

    # Filtrera kolumner
    final_cols = [c for c in cols_to_show if c in visning.columns]
    visning = visning[final_cols]
    
    # Byt namn
    rename_map = {'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn', 'Kritik': 'Domarberättelse'}
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
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "🏥 Hälsa", "🔍 Felsökning (Rådata)", "❓ Guide", "ℹ️ Info"])

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
            c3.metric("Godkända (4-4+)", f"{len(godkanda)} ({round(len(godkanda)/num_starts*100, 1) if num_starts>0 else 0}%)")
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
    else: st.info("Ingen data. Ladda upp fil eller kontrollera 'Felsökning'.")

# === FLIK 2: FÄLT ===
with tab2:
    if not df_j_filt.empty:
        st.subheader("Jaktprov / Fält")
        c1, c2 = st.columns(2)
        num_starts_j = len(df_j_filt)
        c1.metric("Antal Starter", f"{num_starts_j}", delta=get_gender_breakdown(df_j_filt), delta_color="off")
        c2.metric("Unika Hundar", df_j_filt['Regnr'].nunique() if 'Regnr' in df_j_filt.columns else 0)
        st.divider()

        # Försök hitta pris/resultat
        pris_col = None
        if 'Resultat' in df_j_filt.columns: pris_col = 'Resultat'
        elif 'Pris' in df_j_filt.columns: pris_col = 'Pris'
        
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
    else: st.info("Ingen data. Ladda upp fil eller kontrollera 'Felsökning'.")

# === FLIK 3: HÄLSA ===
with tab3:
    st.subheader("🏥 Hälsostatistik")
    if df_h_filt.empty:
        st.info("Ingen hälsodata laddad.")
    else:
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
            st.warning("Hittade inte resultat-kolumnen.")
            st.dataframe(df_h_filt)

# === FLIK 4: FELSÖKNING (NY) ===
with tab4:
    st.markdown("### 🕵️‍♀️ Rådata & Felsökning")
    st.markdown("Om tabellerna är tomma kan du se här vad appen faktiskt har lyckats läsa in.")
    
    if uploaded_excel or uploaded_csv_e or uploaded_txt:
        st.markdown("**Inläsningslogg:**")
        for msg in debug_logs:
            st.text(f"- {msg}")
            
        st.markdown("---")
        st.markdown(f"**Rådata Eftersök ({len(df_e_raw)} rader):**")
        if not df_e_raw.empty:
            st.dataframe(df_e_raw.head(5))
            st.text(f"Kolumner: {list(df_e_raw.columns)}")
        else:
            st.warning("Eftersöks-tabellen är tom.")

        st.markdown(f"**Rådata Fält ({len(df_j_raw)} rader):**")
        if not df_j_raw.empty:
            st.dataframe(df_j_raw.head(5))
            st.text(f"Kolumner: {list(df_j_raw.columns)}")
        else:
            st.warning("Fält-tabellen är tom.")
            
    else:
        st.info("Ladda upp en fil först.")

# === FLIK 5: GUIDE ===
with tab5:
    st.markdown("## 📘 Användarguide (SVK)")
    st.markdown("Verktyget används av **avelsråd** och **funktionärer** inom SVK.")
    
    st.markdown("""
    ### 📂 Steg 1: Hämta & Ladda upp
    1. **Provresultat:** Ladda ner "Årsstatistik"-filen (Excel) från gemensam drive.
    2. **Hälsodata:** Hämta från **[SKK Avelsdata](https://hundar.skk.se/avelsdata)**.
    3. **Ladda upp:** Använd menyn till vänster.

    ### 🎯 Steg 2: Filtrera
    * **Ras, År, Klass:** Justera filtren för att avgränsa urvalet.
    * **Kön:** Filtrera på Hanar/Tikar vid behov.

    ### 📊 Steg 3: Analysera
    Använd flikarna för att se statistik och kopiera siffror till årssammanställningen.
    """)

# === FLIK 6: INFO ===
with tab6:
    st.header("ℹ️ Information")
    st.markdown("Detta verktyg hanterar data temporärt. Ingen data sparas permanent.")

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
