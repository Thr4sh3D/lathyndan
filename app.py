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

st.sidebar.markdown("### 🌲 Provresultat")
st.sidebar.info("Ladda upp Master-filen (Excel).")
uploaded_excel = st.sidebar.file_uploader("Excel-fil (SVK Årsstatistik)", type=["xlsx"], key="excel")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Hälsodata")
uploaded_halsa = st.sidebar.file_uploader("Hälsofil från Avelsdata (Excel)", type=["xlsx", "txt", "csv"], key="halsa")

# --- IDENTIFIERING AV FLIKAR ---
def identify_sheet_type(df, sheet_name):
    """
    Gissar om en flik är Eftersök eller Fält baserat på kolumner.
    """
    cols = [str(c).lower().strip() for c in df.columns]
    
    # Debug-info (sparas för att visas om det krånglar)
    debug_msg = f"Flik: '{sheet_name}' | Kolumner: {cols[:5]}..."

    # Fältprov har ofta "pris", "egenskap", "fält", "resultat"
    if any(k in c for c in cols for k in ["fält", "resultat", "pris", "egenskap"]):
        return "Fält", debug_msg
    
    # Eftersök har ofta "vatten", "spår", "eftersök"
    if any(k in c for c in cols for k in ["vatten", "spår"]):
        return "Eftersök", debug_msg
        
    return "Okänd", debug_msg

# --- DATALOAD ---
@st.cache_data
def load_prov_data(excel):
    df_e_list = []
    df_j_list = []
    debug_log = []
    
    if excel:
        try:
            xls = pd.ExcelFile(excel)
            for sheet_name in xls.sheet_names:
                # Läs in och rensa lite direkt
                df = pd.read_excel(xls, sheet_name=sheet_name)
                df = df.dropna(how='all') # Ta bort tomma rader
                
                typ, msg = identify_sheet_type(df, sheet_name)
                debug_log.append(f"{msg} -> **{typ}**")
                
                if typ == "Eftersök":
                    df_e_list.append(df)
                elif typ == "Fält":
                    df_j_list.append(df)
                else:
                    # FALLBACK: Om fliken heter "Blad1" gissa Fält, "Blad2" gissa Eftersök
                    if "1" in sheet_name:
                        df_j_list.append(df)
                        debug_log.append(f"⚠️ Tvingade '{sheet_name}' till Fält (Fallback)")
                    elif "2" in sheet_name:
                        df_e_list.append(df)
                        debug_log.append(f"⚠️ Tvingade '{sheet_name}' till Eftersök (Fallback)")

        except Exception as e:
            return pd.DataFrame(), pd.DataFrame(), [f"Fel vid inläsning: {e}"]

    df_e_tot = pd.concat(df_e_list, ignore_index=True) if df_e_list else pd.DataFrame()
    df_j_tot = pd.concat(df_j_list, ignore_index=True) if df_j_list else pd.DataFrame()
    
    return df_e_tot, df_j_tot, debug_log

def parse_health_file(file):
    try:
        if file.name.endswith('.xlsx'): return pd.read_excel(file)
        content = file.getvalue()
        try: txt = content.decode("latin-1")
        except: txt = content.decode("utf-8", errors="ignore")
        return pd.read_csv(io.StringIO(txt), sep='\t', on_bad_lines='skip')
    except: return pd.DataFrame()

df_e_raw, df_j_raw, debug_info = load_prov_data(uploaded_excel)
df_halsa_raw = parse_health_file(uploaded_halsa) if uploaded_halsa else pd.DataFrame()

# --- FELSÖKNINGSRUTA I MENYN ---
if uploaded_excel:
    with st.sidebar.expander("🛠️ Vad läser appen?"):
        for log in debug_info:
            st.write(log)
        st.write(f"Antal rader Eftersök: {len(df_e_raw)}")
        st.write(f"Antal rader Fält: {len(df_j_raw)}")

# --- CLEANING ---
def clean_df(df):
    if df.empty: return df
    
    # Rensa kolumnnamn (ta bort mellanslag)
    df.columns = df.columns.str.strip()
    
    # Mappning för att hitta rätt kolumn oavsett vad den heter
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'datum' in cl and 'röntgen' not in cl: col_map['Datum'] = c
        if 'klass' in cl: col_map['Klass'] = c
        if 'ras' in cl: col_map['Ras'] = c
        if 'regnr' in cl or 'reg.nr' in cl: col_map['Regnr'] = c
        
        # Kön-detektiv
        if 'kön' in cl: col_map['Kön'] = c
        elif c.lower() == 'sex': col_map['Kön'] = c
        
        # Hundnamn
        if 'hund' in cl and 'namn' in cl: col_map['Hund'] = c
        elif 'hund' in cl and 'fader' not in cl and 'moder' not in cl: col_map['Hund'] = c
        elif 'namn' in cl and 'ras' not in cl and 'ägare' not in cl: col_map['Hund'] = c
        
        # Poäng Eftersök
        if 'vatten' in cl: col_map['Vatten'] = c
        if 'spår' in cl: col_map['Spår'] = c
        
        # Resultat Fält
        if 'resultat' in cl: col_map['Resultat'] = c
        if 'betyg' in cl: col_map['Fältbetyg'] = c
        if 'domare' in cl: col_map['Domare'] = c
        if 'pris' in cl: col_map['Pris'] = c # Ibland heter det Pris istället för Resultat

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
    cols_to_show = ['Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass']
    
    if type == "Eftersök":
        cols_to_show += ['Vatten', 'Spår']
        if 'Domare' in visning.columns: cols_to_show.append('Domare')
    elif type == "Fält":
        # Prioritera Resultat-kolumner
        if 'Resultat' in visning.columns: cols_to_show.append('Resultat')
        elif 'Pris' in visning.columns: cols_to_show.append('Pris')
        if 'Fältbetyg' in visning.columns: cols_to_show.append('Fältbetyg')
        if 'Domare' in visning.columns: cols_to_show.append('Domare')

    # Filtrera
    final_cols = [c for c in cols_to_show if c in visning.columns]
    visning = visning[final_cols]
    
    # Byt namn
    rename_map = {'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn'}
    # Om Resultat saknas men Pris finns, döp om Pris till Resultat för konsekvens
    if 'Pris' in visning.columns and 'Resultat' not in visning.columns:
        rename_map['Pris'] = 'Resultat'
        
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
    else: st.info("Ladda upp data och välj filter för att se statistik.")

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
