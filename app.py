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
st.sidebar.info("Ladda upp filer från SKK eller gemensam Drive.")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (Alla prov)", type=["xlsx"], key="excel")
uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"], key="csv_e")
uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"], key="csv_j")
uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"], key="txt_prov")

# Hälsodata
st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Hälsodata")
st.sidebar.info("Ladda upp Excel-fil från SKK Avelsdata.")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- PARSING PROVDATA ---
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
                    raw_kon = parts[5].strip() if len(parts) > 5 else "Okänd"
                    entry = {
                        "Datum": parts[0],
                        "Klass": parts[1].strip(),
                        "Hund": parts[2],
                        "regnr": parts[3],
                        "Ras_Clean": parts[4].strip(),
                        "Kön_Raw": raw_kon,
                        "Vatten_Final": pd.to_numeric(parts[27], errors='coerce'),
                        "Spår_Final": pd.to_numeric(parts[28], errors='coerce'),
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except: continue
    return pd.DataFrame(data)

# --- PARSING HÄLSODATA ---
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
    except:
        return pd.DataFrame()

@st.cache_data
def load_prov_data(excel, csv_e, csv_j, txt):
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

df_e_raw, df_j_raw = load_prov_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)
df_halsa_raw = parse_health_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()

# --- AVANCERAT: KOLUMN-KOLL ---
st.sidebar.divider()
valda_kon_kolumn = None
if not df_e_raw.empty:
    with st.sidebar.expander("🛠️ Felsökning: Provdata"):
        cols = ["(Auto-detektera)"] + list(df_e_raw.columns)
        val = st.selectbox("Om kön saknas, välj kolumn:", cols)
        if val != "(Auto-detektera)": valda_kon_kolumn = val

# --- CLEANING ---
def clean_df(df, manual_sex_col=None):
    if df.empty: return df
    df.columns = df.columns.str.strip()
    
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year

    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()
    
    # Ras
    if 'Ras_Clean' not in df.columns:
        ras_candidates = ['rasnamn', 'Ras', 'Hundras']
        found_ras = next((c for c in ras_candidates if c in df.columns), None)
        df['Ras_Clean'] = df[found_ras] if found_ras else "Okänd"
    
    # Kön
    found_kon = None
    if manual_sex_col and manual_sex_col in df.columns: found_kon = manual_sex_col
    else:
        kon_candidates = ['Kön', 'Sex', 'Gender', 'Kön_Raw']
        found_kon = next((c for c in kon_candidates if c in df.columns), None)
    
    if found_kon:
        def standardisera_kon(val):
            v = str(val).lower()
            if 'h' in v: return 'Hane'
            if 't' in v: return 'Tik'
            return 'Okänd'
        df['Kön'] = df[found_kon].apply(standardisera_kon)
    else:
        df['Kön'] = "Okänd"
            
    # Poäng
    if 'Vatten_Final' not in df.columns:
        col_v = next((c for c in df.columns if "vatten" in c.lower() and "kritik" not in c.lower()), None)
        if col_v: df['Vatten_Final'] = pd.to_numeric(df[col_v], errors='coerce').fillna(0)
    if 'Spår_Final' not in df.columns:
        col_s = next((c for c in df.columns if ("spår" in c.lower() or "spar" in c.lower()) and "kritik" not in c.lower()), None)
        if col_s: df['Spår_Final'] = pd.to_numeric(df[col_s], errors='coerce').fillna(0)
        
    return df

df_e = clean_df(df_e_raw, valda_kon_kolumn)
df_j = clean_df(df_j_raw, valda_kon_kolumn)

# --- FILTER ---
st.sidebar.header("🔍 2. Filtrering")

years_prov = set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))

# Hälso-datum logic
if not df_halsa_raw.empty:
    date_col = next((c for c in df_halsa_raw.columns if 'datum' in c.lower()), None)
    if date_col:
        df_halsa_raw[date_col] = pd.to_datetime(df_halsa_raw[date_col], errors='coerce')
        df_halsa_raw['År'] = df_halsa_raw[date_col].dt.year
        years_prov = years_prov | set(df_halsa_raw['År'].dropna().astype(int))

all_years = sorted(list(years_prov))
all_races = sorted(list(set(df_e.get('Ras_Clean', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras_Clean', pd.Series()).dropna().astype(str))))
all_classes = sorted(list(set(df_e.get('Klass', pd.Series()).dropna().astype(str)) | set(df_j.get('Klass', pd.Series()).dropna().astype(str))))
all_sex = sorted(list(set(df_e.get('Kön', pd.Series()).dropna().astype(str)) | set(df_j.get('Kön', pd.Series()).dropna().astype(str))))

if not all_races: all_races = ["Alla/Okänd"]

valda_raser = st.sidebar.multiselect("Välj Ras", all_races, default=all_races[:1] if all_races else None)
valda_ar = st.sidebar.multiselect("Välj År", all_years, default=all_years)
valda_klasser = st.sidebar.multiselect("Välj Klass (Prov)", all_classes, default=all_classes)
valda_kon = st.sidebar.multiselect("Välj Kön (Prov)", all_sex, default=all_sex)

def apply_filter(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser and 'Ras_Clean' in temp.columns: temp = temp[temp['Ras_Clean'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    if valda_klasser and 'Klass' in temp.columns: temp = temp[temp['Klass'].isin(valda_klasser)]
    if valda_kon and 'Kön' in temp.columns: temp = temp[temp['Kön'].isin(valda_kon)]
    return temp

df_e_filt = apply_filter(df_e)
df_j_filt = apply_filter(df_j)
df_h_filt = apply_filter(df_halsa_raw)

# --- FUNKTION: SNYGGA TABELLER MED LÄNKAR ---
def prepare_display_table(df, is_field=False, add_links=False):
    visning = df.copy()
    if 'Datum_Str' in visning.columns:
        if 'Datum' in visning.columns: visning = visning.drop(columns=['Datum'])
        visning = visning.rename(columns={'Datum_Str': 'Datum'})

    # Byt namn till svenska
    rename_map = {
        'regnr': 'Reg.nr', 'Hund': 'Hundnamn', 'namn': 'Hundnamn',
        'Vatten_Final': 'Vatten', 'Spår_Final': 'Spår', 'Ras_Clean': 'Ras',
        'Kritik': 'Domarberättelse', 'Kön': 'Kön'
    }
    if is_field:
        pris_col = next((c for c in df.columns if "pris" in c.lower()), None)
        if pris_col: rename_map[pris_col] = 'Pris'
    
    visning = visning.rename(columns=rename_map)
    visning = visning.loc[:, ~visning.columns.duplicated()]

    # SKAPA LÄNK (Endast om vi vill visa på skärm)
    if add_links and 'Reg.nr' in visning.columns:
        base_url = "https://hundar.skk.se/hunddata/Hund_sok.aspx?sok="
        visning['Reg.nr'] = visning['Reg.nr'].apply(lambda x: f"{base_url}{x}" if pd.notna(x) else x)

    # Välj kolumner
    priority_cols = ['Datum', 'Reg.nr', 'Hundnamn', 'Kön', 'Ras', 'Klass']
    if not is_field: priority_cols += ['Vatten', 'Spår']
    else: 
        if 'Pris' in visning.columns: priority_cols += ['Pris']
    if 'Domarberättelse' in visning.columns: priority_cols += ['Domarberättelse']
    
    final_cols = [c for c in priority_cols if c in visning.columns]
    return visning[final_cols]

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
        gender_info = get_gender_breakdown(df_e_filt)
        c1.metric("Antal Starter", f"{num_starts}", delta=gender_info, delta_color="off")
        c2.metric("Unika Hundar", df_e_filt['regnr'].nunique() if 'regnr' in df_e_filt.columns else 0)

        if 'Vatten_Final' in df_e_filt.columns and 'Spår_Final' in df_e_filt.columns:
            godkanda = df_e_filt[(df_e_filt['Vatten_Final'] >= 4) & (df_e_filt['Spår_Final'] >= 4)]
            full = df_e_filt[(df_e_filt['Vatten_Final'] == 10) & (df_e_filt['Spår_Final'] == 10)]
            c3.metric("Godkända (4-4+)", f"{len(godkanda)} ({round(len(godkanda)/num_starts*100, 1)}%)")
            c4.metric("Full pott (10-10)", f"{len(full)}")

            st.divider()
            col_v1, col_v2 = st.columns([1, 2])
            with col_v1:
                st.markdown("##### 💧 Vattenbetyg")
                v_counts = df_e_filt['Vatten_Final'].value_counts().sort_index(ascending=False).rename("Antal")
                st.dataframe(v_counts, use_container_width=True)
            with col_v2:
                fig_v = px.bar(x=v_counts.index, y=v_counts.values, labels={'x': 'Betyg', 'y': 'Antal'}, color_discrete_sequence=['#3366CC'])
                fig_v.update_layout(xaxis=dict(tickmode='linear', dtick=1), showlegend=False)
                st.plotly_chart(fig_v, use_container_width=True, key="v_chart")

            st.divider()
            col_s1, col_s2 = st.columns([1, 2])
            with col_s1:
                st.markdown("##### 🌲 Spårbetyg")
                s_counts = df_e_filt['Spår_Final'].value_counts().sort_index(ascending=False).rename("Antal")
                st.dataframe(s_counts, use_container_width=True)
            with col_s2:
                fig_s = px.bar(x=s_counts.index, y=s_counts.values, labels={'x': 'Betyg', 'y': 'Antal'}, color_discrete_sequence=['#109618'])
                fig_s.update_layout(xaxis=dict(tickmode='linear', dtick=1), showlegend=False)
                st.plotly_chart(fig_s, use_container_width=True, key="s_chart")
            
        st.markdown("---")
        st.markdown("##### Detaljlista (Klicka på Reg.nr)")
        df_visning_e = prepare_display_table(df_e_filt, is_field=False, add_links=True)
        st.dataframe(
            df_visning_e,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)"),
                "Kön": st.column_config.TextColumn(width="small"),
                "Vatten": st.column_config.NumberColumn(width="small"),
                "Spår": st.column_config.NumberColumn(width="small"),
                "Domarberättelse": st.column_config.TextColumn(width="large"),
            }
        )
    else: st.info("Ingen provdata laddad eller matchar filtret.")

# === FLIK 2: FÄLT ===
with tab2:
    if not df_j_filt.empty:
        st.subheader("Jaktprov / Fält")
        c1, c2 = st.columns(2)
        num_starts_j = len(df_j_filt)
        gender_info_j = get_gender_breakdown(df_j_filt)
        c1.metric("Antal Starter", f"{num_starts_j}", delta=gender_info_j, delta_color="off")
        c2.metric("Unika Hundar", df_j_filt['regnr'].nunique() if 'regnr' in df_j_filt.columns else 0)
        st.divider()

        pris_col = next((c for c in df_j_filt.columns if "pris" in c.lower()), None)
        if pris_col:
            st.markdown("##### 🏆 Prisfördelning")
            col_p1, col_p2 = st.columns([1, 2])
            p_counts = df_j_filt[pris_col].value_counts().rename("Antal")
            with col_p1: st.dataframe(p_counts, use_container_width=True)
            with col_p2: st.plotly_chart(px.pie(values=p_counts.values, names=p_counts.index, hole=0.4), use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### Detaljlista (Klicka på Reg.nr)")
        df_visning_j = prepare_display_table(df_j_filt, is_field=True, add_links=True)
        st.dataframe(
            df_visning_j,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)"),
                "Kön": st.column_config.TextColumn(width="small"),
                "Pris": st.column_config.TextColumn(width="small"),
                "Domarberättelse": st.column_config.TextColumn(width="large"),
            }
        )
    else: st.info("Ingen provdata matchar filtret.")

# === FLIK 3: HÄLSA ===
with tab3:
    st.subheader("🏥 Hälsostatistik (HD/ED)")
    if df_h_filt.empty:
        st.info("Ingen hälsodata laddad. Se 'Guide' för instruktioner.")
    else:
        cols = df_h_filt.columns
        res_col = next((c for c in cols if 'resultat' in c.lower() or 'diagnos' in c.lower()), None)
        
        if res_col:
            st.write(f"Visar statistik för: **{res_col}**")
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
    st.markdown("Detta verktyg är ett stöd för dig som är **avelsråd eller avelsfunktionär** inom Svenska Vorstehklubben (SVK).")
    
    st.markdown("""
    ### 📂 Steg 1: Hämta Data
    För att använda verktyget behöver du ladda upp filer.

    **A. Provresultat (Eftersök/Fält)**
    * Dessa filer hämtas vanligtvis från er **Gemensamma Google Drive**.
    * Om du saknar åtkomst, kontakta avelskommittén.

    **B. Hälsodata (HD/ED)**
    * Gå till **[SKK Avelsdata](https://hundar.skk.se/avelsdata)**.
    * Sök på din ras (t.ex. Korthårig Vorsteh eller Kleiner Münsterländer).
    * Klicka på fliken **"Hälsa"** -> Välj diagnos (t.ex. HD).
    * Klicka på **Excel-ikonen** för att ladda ner listan.

    ---

    ### 📂 Steg 2: Ladda upp & Filtrera
    1. Ladda upp filerna i menyn till vänster.
    2. Välj **Ras** och **År** i filtret.
    3. Om du ska göra **årssammanställning**, välj även **Klass**.
    4. Om filen saknar "Kön", använd felsökningsverktyget i menyn.

    ### 📊 Steg 3: Analysera
    * **Eftersök:** Betygsfördelning, godkända hundar och könsuppdelning.
    * **Fältprov:** Prisfördelning och detaljer.
    * **Hälsa:** Stapeldiagram över HD/ED-resultat.
    
    *Tips: Klicka på registreringsnumret i listorna för att se hunden på SKK.*
    """)

# === FLIK 5: INFO ===
with tab5:
    st.header("ℹ️ Information & Säkerhet")
    st.markdown("""
    ### 🛡️ SVK Datapolicy
    Detta verktyg tillhandahålls för avelsfunktionärer inom SVK för att underlätta arbetet med statistik.
    
    * **Syfte:** Effektivisera avelsuppföljning och framtagande av årssammanställningar.
    * **Lagring:** Ingen data sparas. All bearbetning sker i arbetsminnet och raderas vid stängning.
    * **Personuppgifter:** Verktyget hanterar resultatlistor med stöd av *berättigat intresse* för föreningens avelsarbete.
    """)

# --- EXPORT ---
st.divider()
output = io.BytesIO()

with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: 
        prepare_display_table(df_e_filt, is_field=False, add_links=False).to_excel(writer, index=False, sheet_name='Eftersök')
    if not df_j_filt.empty: 
        prepare_display_table(df_j_filt, is_field=True, add_links=False).to_excel(writer, index=False, sheet_name='Fält')
    if not df_h_filt.empty:
        df_h_filt.to_excel(writer, index=False, sheet_name='Hälsa')

st.download_button("📥 Ladda ner Statistik (Excel)", output.getvalue(), "SVK_Statistik.xlsx", "application/vnd.ms-excel")
