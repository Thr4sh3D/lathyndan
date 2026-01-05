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
        
        # HUNDNAMN (Ignorera ägare)
        if 'ag_fnamn' in cl or 'ägare' in cl: continue
        if c == 'namn': col_map['Hund'] = c
        elif 'hund' in cl and 'fader' not in cl and 'moder' not in cl: col_map['Hund'] = c
        
        # EFTERSÖK
        if 'vatten' in cl and 'passion' not in cl: col_map['Vatten'] = c
        if 'spår' in cl or 'spar' in cl: col_map['Spår'] = c 
        if 'passion' in cl: col_map['Vattenpassion'] = c
        
        # FÄLT
        if 'resultat' in cl: col_map['Resultat'] = c
        elif 'pris' in cl and 'egenskap' not in cl: col_map['Resultat'] = c 
        if 'domare' in cl: col_map['Domare'] = c
        
        # EGENSKAPER
        if 'fart' in cl: col_map['Fart'] = c
        if 'vidd' in cl: col_map['Vidd'] = c
        if 'reviering' in cl: col_map['Reviering'] = c
        if 'följsamhet' in cl or 'foljsamhet' in cl: col_map['Följsamhet'] = c

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
        df['Svensk'] = df['Regnr'].astype(str).str.upper().str.match(r'^(SE|S\d)')
    else: df['Svensk'] = False

    # Könfix
    if 'Kön' in df.columns:
        df['Kön'] = df['Kön'].astype(str).apply(lambda x: 'Hane' if 'h' in x.lower() else ('Tik' if 't' in x.lower() else 'Okänd'))
    else: df['Kön'] = "Okänd"
    
    # Klassfix
    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()

    # Siffror (Poäng & Egenskaper)
    num_cols = ['Vatten', 'Spår', 'Resultat', 'Fart', 'Vidd', 'Reviering', 'Följsamhet', 'Vattenpassion']
    for col in num_cols:
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
    
    # 1. Definiera önskade kolumner
    desired_cols = [
        'Datum_Str', 'Regnr', 'Hund', 'Kön', 'Ras', 'Klass', 
        'Resultat', # Resultat behövs för logik, men omdöps till Pris senare
        'Vatten', 'Spår', 'Vattenpassion',
        'Fart', 'Vidd', 'Reviering', 'Följsamhet', 'Domare'
    ]
    
    # 2. Plocka ENDAST ut kolumner som finns i v (undvik KeyError)
    existing_cols = [c for c in desired_cols if c in v.columns]
    v = v[existing_cols] # Nu har vi rensat bort allt skräp
    
    # 3. Döp om till snygga rubriker
    rename_map = {'Datum_Str': 'Datum', 'Regnr': 'Reg.nr', 'Hund': 'Hundnamn', 'Resultat': 'Pris'}
    v = v.rename(columns=rename_map)
    
    # 4. Formatera Pris (Endast om kolumnen Pris nu finns)
    if 'Pris' in v.columns:
        # Se till att det är heltal för snyggare formatting
        v['Pris'] = pd.to_numeric(v['Pris'], errors='coerce').fillna(0).astype(int)
        v['Pris'] = v['Pris'].apply(lambda x: f"{x}:a Pris" if x > 0 else "")

    # 5. Skapa länkar
    if add_links and 'Reg.nr' in v.columns:
        base = "https://hundar.skk.se/hunddata/Hund_sok.aspx?sok="
        v['Reg.nr'] = v['Reg.nr'].apply(lambda x: f"{base}{x}" if pd.notna(x) else x)
        
    return v

# --- TABS ---
tab_names = ["📄 Underlag Årsrapport", "🌲 Eftersök", "🌾 Fältprov", "🐕 Fullbruk", "🩸 Viltspår", "🏥 Hälsa", "❓ Guide", "ℹ️ Info"]
tabs = st.tabs(tab_names)

# === TAB 1: ÅRSRAPPORT ===
with tabs[0]:
    st.header("📄 Underlag för Årssammanställning")
    st.markdown("Här samlas statistiken exakt så som den efterfrågas i rapportmallen. Notera att **rasfiltret** i menyn används, men **årsfiltret ignoreras** här för att visa jämförelser.")
    
    if not valda_raser:
        st.warning("Välj en ras i menyn till vänster.")
    else:
        df_j_ras = df_j[df_j['Ras'].isin(valda_raser)] if not df_j.empty else pd.DataFrame()
        df_e_ras = df_e[df_e['Ras'].isin(valda_raser)] if not df_e.empty else pd.DataFrame()
        df_full_ras = df_full[df_full['Ras'].isin(valda_raser)] if not df_full.empty else pd.DataFrame()
        df_vilt_ras = df_vilt[df_vilt['Ras'].isin(valda_raser)] if not df_vilt.empty else pd.DataFrame()

        # TABELL 2
        st.subheader("Tabell 2: Antal starter och unika hundar (per år)")
        if not df_j_ras.empty or not df_e_ras.empty:
            f_stats = df_j_ras.groupby('År').agg(Starter_Fält=('Regnr', 'count'), Unika_Fält=('Regnr', 'nunique'))
            e_stats = df_e_ras.groupby('År').agg(Starter_Eftersök=('Regnr', 'count'), Unika_Eftersök=('Regnr', 'nunique'))
            tab2 = pd.concat([f_stats, e_stats], axis=1).sort_index(ascending=False).fillna(0).astype(int)
            st.dataframe(tab2, use_container_width=True)
        else: st.info("Ingen data för Tabell 2.")

        # TABELL 3
        st.subheader("Tabell 3: Prisfördelning svenskregistrerade hundar (Senaste 2 åren)")
        if not df_j_ras.empty:
            # Endast SE-hundar och Pris > 0
            df_se = df_j_ras[(df_j_ras['Svensk'] == True) & (df_j_ras['Resultat'] > 0)]
            avail_years = sorted(list(set(df_j_ras['År'])), reverse=True)[:2]
            
            if df_se.empty:
                st.info("Inga pris tagna av svenska hundar.")
            else:
                for y in avail_years:
                    st.markdown(f"**År {y}**")
                    df_y = df_se[df_se['År'] == y]
                    if not df_y.empty:
                        df_y['PrisLabel'] = df_y['Resultat'].astype(int).astype(str) + ":a pris"
                        pivot = pd.crosstab(df_y['Resultat'], df_y['Klass'])
                        cols = [c for c in ['EKL', 'ÖKL', 'UKL'] if c in pivot.columns]
                        pivot = pivot[cols]
                        pivot['Totalt'] = pivot.sum(axis=1)
                        pivot.index = [f"{i}:a pris" for i in pivot.index]
                        st.dataframe(pivot, use_container_width=True)
                    else: st.text("Ingen data.")
        else: st.info("Ingen data för Tabell 3.")

        # TABELL 4
        st.subheader("Tabell 4: Andel med optimalt värde (4) i %")
        egenskaper = ['Fart', 'Vidd', 'Reviering', 'Följsamhet', 'Vattenpassion']
        df_all_props = pd.concat([df_j_ras, df_e_ras], ignore_index=True)
        
        if not df_all_props.empty:
            res_list = []
            avail_years_3 = sorted(list(set(df_all_props['År'])), reverse=True)[:3]
            for y in sorted(avail_years_3):
                row = {'År': y}
                df_y = df_all_props[df_all_props['År'] == y]
                for prop in egenskaper:
                    if prop in df_y.columns:
                        scored = df_y[df_y[prop] > 0]
                        if len(scored) > 0:
                            fours = len(scored[scored[prop] == 4])
                            perc = round((fours / len(scored)) * 100)
                            row[prop] = f"{perc}%"
                        else: row[prop] = "-"
                    else: row[prop] = "-"
                res_list.append(row)
            if res_list: st.dataframe(pd.DataFrame(res_list).set_index('År'), use_container_width=True)
        else: st.info("Ingen data för Tabell 4.")

        # TABELL 5 & 6 (Fullbruk & Viltspår)
        col5, col6 = st.columns(2)
        with col5:
            st.subheader("Tabell 5: Fullbruk (SE)")
            if not df_full_ras.empty:
                df_full_se = df_full_ras[df_full_ras['Svensk']==True]
                if not df_full_se.empty:
                    fb_stats = df_full_se.groupby('År').agg(
                        Starter=('Regnr', 'count'), Unika=('Regnr', 'nunique'),
                        Pris_1=('Resultat', lambda x: (x==1).sum()),
                        Pris_2=('Resultat', lambda x: (x==2).sum()),
                        Pris_3=('Resultat', lambda x: (x==3).sum())
                    ).sort_index(ascending=False)
                    st.dataframe(fb_stats, use_container_width=True)
                else: st.text("Inga svenska hundar.")
            else: st.info("Ladda upp Fullbruksfil.")

        with col6:
            st.subheader("Tabell 6: Viltspår (SE)")
            if not df_vilt_ras.empty:
                df_vilt_se = df_vilt_ras[df_vilt_ras['Svensk']==True]
                if not df_vilt_se.empty:
                    v_res = []
                    for y in sorted(list(set(df_vilt_se['År'])), reverse=True):
                        d = df_vilt_se[df_vilt_se['År'] == y]
                        akl = d[d['Klass'].str.contains('anlag', case=False, na=False)]
                        okl = d[d['Klass'].str.contains('öppen', case=False, na=False)]
                        godkanda_akl = len(akl[akl['Resultat'].astype(str).str.contains('godk', case=False) | (akl['Resultat']==1)])
                        row = {
                            'År': y, 'Starter AKL': len(akl), 'Godkända AKL': godkanda_akl,
                            'Starter ÖKL': len(okl),
                            '1:a ÖKL': len(okl[okl['Resultat']==1]), '2:a ÖKL': len(okl[okl['Resultat']==2]), '3:a ÖKL': len(okl[okl['Resultat']==3]),
                        }
                        v_res.append(row)
                    st.dataframe(pd.DataFrame(v_res).set_index('År'), use_container_width=True)
            else: st.info("Ladda upp Viltspårfil.")

# === EFTERSÖK (TAB 2) ===
with tabs[1]:
    if not df_e_filt.empty:
        st.subheader("Eftersök (Vatten & Spår)")
        c1, c2, c3 = st.columns(3)
        starts = len(df_e_filt)
        unika = df_e_filt['Regnr'].nunique()
        c1.metric("Antal Starter", f"{starts}", get_gender_text(df_e_filt))
        c2.metric("Unika Individer", f"{unika}")
        
        if 'Vatten' in df_e_filt.columns and 'Spår' in df_e_filt.columns:
            ok = df_e_filt[(df_e_filt['Vatten'] >= 4) & (df_e_filt['Spår'] >= 4)]
            c3.metric("Godkända (4+)", f"{len(ok)}")
            
            st.divider()
            c_a, c_b, c_c = st.columns(3)
            with c_a: 
                st.markdown("**💧 Vattenbetyg**")
                st.dataframe(df_e_filt['Vatten'].value_counts().sort_index(ascending=False), use_container_width=True)
            with c_b: 
                st.markdown("**🌲 Spårbetyg**")
                st.dataframe(df_e_filt['Spår'].value_counts().sort_index(ascending=False), use_container_width=True)
            with c_c:
                if 'Vattenpassion' in df_e_filt.columns:
                    st.markdown("**🌊 Vattenpassion**")
                    st.dataframe(df_e_filt['Vattenpassion'].value_counts().sort_index(ascending=False), use_container_width=True)

        st.dataframe(prepare_table(df_e_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Eftersök.")

# === FÄLTPROV (TAB 3) ===
with tabs[2]:
    if not df_j_filt.empty:
        st.subheader("Jaktprov / Fält")
        c1, c2 = st.columns(2)
        c1.metric("Antal Starter", len(df_j_filt), get_gender_text(df_j_filt))
        c2.metric("Unika Individer", f"{df_j_filt['Regnr'].nunique()}")
        
        st.divider()
        st.markdown("### 🇸🇪 Prisfördelning Svenska Hundar (Senaste 2 åren)")
        
        sel_year = int(valda_ar[0]) if valda_ar else 2024
        two_years = [sel_year, sel_year-1]

        df_se_2y = df_j[
            (df_j['Ras'].isin(valda_raser) if valda_raser else True) & 
            (df_j['Svensk'] == True) & 
            (df_j['År'].isin(two_years)) &
            (df_j['Resultat'] > 0) 
        ]

        if not df_se_2y.empty and 'Resultat' in df_se_2y.columns and 'Klass' in df_se_2y.columns:
            df_se_2y['PrisLabel'] = df_se_2y['Resultat'].astype(int).astype(str) + ":a Pris"
            pivot = pd.crosstab(df_se_2y['Klass'], df_se_2y['PrisLabel'], margins=True, margins_name="Totalt")
            cols = sorted([c for c in pivot.columns if c != "Totalt"]) + ["Totalt"]
            st.dataframe(pivot[cols], use_container_width=True)
        else:
            st.warning(f"Inga svenska hundar med pris hittades för åren {two_years}.")

        st.divider()
        st.markdown("### 📏 Egenskaper i Fält (Medelvärde > 0)")
        egenskaper = ['Fart', 'Vidd', 'Reviering', 'Följsamhet']
        valid_eg = [e for e in egenskaper if e in df_j_filt.columns]
        
        if valid_eg:
            stats = {}
            for e in valid_eg:
                avg = df_j_filt[df_j_filt[e] > 0][e].mean()
                stats[e] = round(avg, 2) if pd.notna(avg) else 0
            cols = st.columns(len(valid_eg))
            for i, (k, v) in enumerate(stats.items()):
                cols[i].metric(k, f"{v}")
        
        st.divider()
        st.markdown("**Alla starter (Detaljlista)**")
        st.dataframe(prepare_table(df_j_filt), use_container_width=True, hide_index=True,
                     column_config={"Reg.nr": st.column_config.LinkColumn("Reg.nr", display_text=r"sok=(.*)")})
    else: st.info("Ingen data för Fält.")

# === ÖVRIGA TABS (Kortade men funktionella) ===
with tabs[3]: # Fullbruk
    if not df_full_filt.empty: st.dataframe(prepare_table(df_full_filt), use_container_width=True, hide_index=True)
    else: st.info("Ladda fil.")
with tabs[4]: # Viltspår
    if not df_vilt_filt.empty: st.dataframe(prepare_table(df_vilt_filt), use_container_width=True, hide_index=True)
    else: st.info("Ladda fil.")
with tabs[5]: # Hälsa
    if not df_h_filt.empty: st.dataframe(df_h_filt, use_container_width=True)
    else: st.info("Ladda fil.")

with tabs[6]: # Guide
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

    ### 📊 Steg 3: Analys & Sammanställning (Flik 1)
    
    **📄 Underlag Årsrapport**
    * Denna flik samlar alla tabeller du behöver till rapporten (Tabell 2-6).
    * Den räknar ut starter, unika hundar, egenskapsvärden och prisfördelning för SE-hundar.

    **🔗 Tips:**
    I detaljlistorna är registreringsnumret en länk. Klicka på det för att komma direkt till hundens sida på SKK Hunddata.
    """)

with tabs[7]: # Info
    st.header("ℹ️ Information, Säkerhet & GDPR")
    st.markdown("""
    ### 🛡️ SVK Datapolicy
    Detta verktyg tillhandahålls för funktionärer inom Svenska Vorstehklubben (SVK).
    * **Syfte:** Effektivisera framtagandet av statistik till årssammanställningar och avelsutvärdering.
    * **Rättslig grund:** Personuppgiftsbehandlingen (namn i resultatlistor) sker med stöd av *berättigat intresse* för föreningens avelsarbete och verksamhetsuppföljning.
    ### 🔐 Datasäkerhet
    * **Ingen lagring:** De filer du laddar upp bearbetas endast i serverns tillfälliga arbetsminne.
    * **Automatisk radering:** Så fort du stänger webbläsarfliken eller laddar om sidan raderas all data omedelbart.
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
