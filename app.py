import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Statistik", page_icon="🐕", layout="wide")

st.title("📊 KLM Statistikverktyg (Dev-mode)")
st.markdown("Automatiskt underlag för **Verksamhetsberättelse** och avelsuppföljning.")

# --- 1. LADDA UPP FILER ---
st.sidebar.header("📂 1. Ladda upp data")
st.sidebar.info("Ladda upp Excel, CSV eller Textfiler.")

uploaded_excel = st.sidebar.file_uploader("Excel-fil (.xlsx)", type=["xlsx"])
uploaded_csv_e = st.sidebar.file_uploader("Eftersök (.csv)", type=["csv"])
uploaded_csv_j = st.sidebar.file_uploader("Jaktprov (.csv)", type=["csv"])
uploaded_txt   = st.sidebar.file_uploader("Textfil SKK (.txt)", type=["txt"])

# --- PARSING & DATALOAD ---
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
                    # Försök gissa kön baserat på om data finns i kolumn 5 (ofta kön i SKK-filer)
                    raw_kon = parts[5].strip() if len(parts) > 5 else "Okänd"
                    
                    entry = {
                        "Datum": parts[0],
                        "Klass": parts[1].strip(),
                        "Hund": parts[2],
                        "regnr": parts[3],
                        "Ras_Clean": parts[4].strip(),
                        "Kön_Raw": raw_kon, # Sparar för tvätt senare
                        "Vatten_Final": pd.to_numeric(parts[27], errors='coerce'),
                        "Spår_Final": pd.to_numeric(parts[28], errors='coerce'),
                        "Kritik": " ".join(parts[40:]).strip()
                    }
                    data.append(entry)
                except: continue
    return pd.DataFrame(data)

@st.cache_data
def load_all_data(excel, csv_e, csv_j, txt):
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

df_e_raw, df_j_raw = load_all_data(uploaded_excel, uploaded_csv_e, uploaded_csv_j, uploaded_txt)

# --- CLEANING & KÖNS-DETEKTIV ---
def clean_df(df):
    if df.empty: return df
    df.columns = df.columns.str.strip()
    
    if 'Datum' in df.columns:
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df['Datum_Str'] = df['Datum'].dt.strftime('%Y-%m-%d')
        df['År'] = df['Datum'].dt.year

    if 'Klass' in df.columns:
        df['Klass'] = df['Klass'].astype(str).str.upper().str.strip()
    
    # --- RAS ---
    if 'Ras_Clean' not in df.columns:
        ras_candidates = ['rasnamn', 'Ras', 'Hundras']
        found_ras = next((c for c in ras_candidates if c in df.columns), None)
        df['Ras_Clean'] = df[found_ras] if found_ras else "Okänd"
    
    # --- KÖN (Ny logik) ---
    # Vi letar efter en kolumn som heter "Kön", "Sex" eller liknande.
    # Om vi hittar "H" eller "Hane" -> Hane, "T" eller "Tik" -> Tik.
    if 'Kön' not in df.columns:
        kon_candidates = ['Kön', 'Sex', 'Gender', 'Kön_Raw']
        found_kon = next((c for c in kon_candidates if c in df.columns), None)
        
        if found_kon:
            # Standardisera till "Hane" och "Tik"
            def standardisera_kon(val):
                v = str(val).lower()
                if 'h' in v: return 'Hane'
                if 't' in v: return 'Tik'
                return 'Okänd'
            df['Kön'] = df[found_kon].apply(standardisera_kon)
        else:
            df['Kön'] = "Okänd (Saknas i fil)"
            
    # --- POÄNG ---
    if 'Vatten_Final' not in df.columns:
        col_v = next((c for c in df.columns if "vatten" in c.lower() and "kritik" not in c.lower()), None)
        if col_v: df['Vatten_Final'] = pd.to_numeric(df[col_v], errors='coerce').fillna(0)
    if 'Spår_Final' not in df.columns:
        col_s = next((c for c in df.columns if ("spår" in c.lower() or "spar" in c.lower()) and "kritik" not in c.lower()), None)
        if col_s: df['Spår_Final'] = pd.to_numeric(df[col_s], errors='coerce').fillna(0)
        
    return df

df_e = clean_df(df_e_raw)
df_j = clean_df(df_j_raw)

# --- FILTER ---
st.sidebar.divider()
st.sidebar.header("🔍 2. Filtrering")

if df_e.empty and df_j.empty:
    st.warning("👈 Börja med att ladda upp data.")
else:
    all_years = sorted(list(set(df_e.get('År', pd.Series()).dropna().astype(int)) | set(df_j.get('År', pd.Series()).dropna().astype(int))))
    all_races = sorted(list(set(df_e.get('Ras_Clean', pd.Series()).dropna().astype(str)) | set(df_j.get('Ras_Clean', pd.Series()).dropna().astype(str))))
    all_classes = sorted(list(set(df_e.get('Klass', pd.Series()).dropna().astype(str)) | set(df_j.get('Klass', pd.Series()).dropna().astype(str))))
    all_sex = sorted(list(set(df_e.get('Kön', pd.Series()).dropna().astype(str)) | set(df_j.get('Kön', pd.Series()).dropna().astype(str))))

    valda_raser = st.sidebar.multiselect("Välj Ras", all_races, default=all_races[:1] if all_races else None)
    valda_ar = st.sidebar.multiselect("Välj År", all_years, default=all_years)
    valda_klasser = st.sidebar.multiselect("Välj Klass", all_classes, default=all_classes)
    valda_kon = st.sidebar.multiselect("Välj Kön", all_sex, default=all_sex)

def apply_filter(df):
    if df.empty: return df
    temp = df.copy()
    if valda_raser: temp = temp[temp['Ras_Clean'].isin(valda_raser)]
    if valda_ar and 'År' in temp.columns: temp = temp[temp['År'].isin(valda_ar)]
    if valda_klasser and 'Klass' in temp.columns: temp = temp[temp['Klass'].isin(valda_klasser)]
    if valda_kon and 'Kön' in temp.columns: temp = temp[temp['Kön'].isin(valda_kon)]
    return temp

df_e_filt = apply_filter(df_e)
df_j_filt = apply_filter(df_j)

# --- FUNKTION: SNYGGA TABELLER ---
def prepare_display_table(df, is_field=False):
    visning = df.copy()
    if 'Datum_Str' in visning.columns:
        if 'Datum' in visning.columns: visning = visning.drop(columns=['Datum'])
        visning = visning.rename(columns={'Datum_Str': 'Datum'})

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

    priority_cols = ['Datum', 'Reg.nr', 'Hundnamn', 'Kön', 'Ras', 'Klass']
    if not is_field: priority_cols += ['Vatten', 'Spår']
    else: 
        if 'Pris' in visning.columns: priority_cols += ['Pris']
    if 'Domarberättelse' in visning.columns: priority_cols += ['Domarberättelse']
    
    final_cols = [c for c in priority_cols if c in visning.columns]
    return visning[final_cols]

# --- HJÄLPFUNKTION: BERÄKNA KÖNSFÖRDELNING ---
def get_gender_breakdown(df):
    if 'Kön' not in df.columns: return ""
    hanar = len(df[df['Kön'] == 'Hane'])
    tikar = len(df[df['Kön'] == 'Tik'])
    okanda = len(df) - hanar - tikar
    return f"({hanar} Hanar, {tikar} Tikar)"

# --- FLIKAR ---
tab1, tab2, tab3, tab4 = st.tabs(["🌲 Eftersök", "🌾 Fältprov", "❓ Guide", "ℹ️ Info & GDPR"])

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
            
            # Tabeller & Diagram
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
        st.markdown("##### Detaljlista")
        
        df_visning_e = prepare_display_table(df_e_filt, is_field=False)
        st.dataframe(
            df_visning_e,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Reg.nr": st.column_config.TextColumn(width="small"),
                "Kön": st.column_config.TextColumn(width="small"),
                "Vatten": st.column_config.NumberColumn(width="small"),
                "Spår": st.column_config.NumberColumn(width="small"),
                "Domarberättelse": st.column_config.TextColumn(width="large"),
            }
        )
    else: st.info("Ingen data matchar filtret.")

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
        st.markdown("##### Detaljlista")
        
        df_visning_j = prepare_display_table(df_j_filt, is_field=True)
        st.dataframe(
            df_visning_j,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Reg.nr": st.column_config.TextColumn(width="small"),
                "Kön": st.column_config.TextColumn(width="small"),
                "Pris": st.column_config.TextColumn(width="small"),
                "Domarberättelse": st.column_config.TextColumn(width="large"),
            }
        )
    else: st.info("Ingen data matchar filtret.")

# === FLIK 3: GUIDE ===
with tab3:
    st.markdown("## 📘 Hjälp & Instruktioner")
    
    st.markdown("""
    ### 📂 Här hittar du filerna
    Resultatfilerna som ska laddas upp finns på vår **gemensamma Google Drive**. 
    * Du behöver behörighet från **Avelskommittén** för att komma åt mappen.
    * Ladda ner filen till din dator först, sedan laddar du upp den här i appen.

    ---

    ### 🎯 Så här gör du (Steg-för-steg)

    #### 1. Ladda upp filen
    Titta i menyn till vänster (på mobil: klicka på pilen `>` högst upp till vänster).
    * Klicka på knappen **Browse files** under rätt rubrik (t.ex. Excel eller Textfil).
    * Välj filen du hämtade från Google Drive.

    #### 2. Välj vad du vill titta på
    I menyn kan du nu filtrera:
    * **Ras:** Kontrollera att det står "Kleiner Münsterländer".
    * **År:** Välj det år du jobbar med (t.ex. 2024).
    * **Klass:** Välj klass (t.ex. UKL eller ÖKL) för att få rätt siffror till rapporten.
    * **Kön:** Du kan nu välja att se bara **Tikar** eller **Hanar**.

    #### 3. Läs av siffrorna
    Nu är det bara att skriva av siffrorna till rapporten!
    * **Eftersök (Tabell 2 & 3):** * Titta på siffran under "Antal Starter". Där står det t.ex. *(12 Hanar, 14 Tikar)*.
        * Använd tabellerna för "Vattenbetyg" och "Spårbetyg".
    * **Fältprov (Tabell 1):** * Gå till fliken **🌾 Fältprov**. Läs av "Prisfördelning" och antalet Hanar/Tikar högst upp.

    ---
    **Tips!**
    Om siffrorna ser konstiga ut, kontrollera att du inte råkat välja fel år eller ras i menyn.
    """)

# === FLIK 4: GDPR & INFO ===
with tab4:
    st.header("ℹ️ Information, Säkerhet & GDPR")
    
    st.markdown("""
    ### 🔐 Datasäkerhet och Lagring
    Denna applikation är utformad enligt principen **"Privacy by Design"**.
    
    * **Ingen lagring:** De filer du laddar upp (Excel/CSV/Txt) bearbetas endast i serverns arbetsminne (RAM). Så fort du stänger webbläsarfliken eller laddar om sidan raderas all data permanent. Inga kopior sparas i någon databas.
    * **Kryptering:** All trafik mellan din dator och servern är krypterad (HTTPS).

    ### 🛡️ Personuppgiftspolicy (GDPR)
    Då vi hanterar resultatlistor som innehåller namn på hundägare/förare, gäller följande:
    
    1.  **Ändamål:** Syftet med behandlingen är att sammanställa anonymiserad statistik för avelsutvärdering och verksamhetsberättelse för Svenska Vorstehklubben (SVK) / KLM.
    2.  **Rättslig grund:** Berättigat intresse (föreningsverksamhet och avelsuppföljning).
    3.  **Lagringstid:** Eftersom ingen data sparas i appen, upphör behandlingen omedelbart efter utfört arbete.

    ### 📞 Support & Kontakt
    Vid tekniska problem eller frågor om appen, kontakta avelskommittén.
    """)

# --- EXPORT ---
st.divider()
output = io.BytesIO()

# Förbered export
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    if not df_e_filt.empty: 
        prepare_display_table(df_e_filt, is_field=False).to_excel(writer, index=False, sheet_name='Eftersök Data')
    if not df_j_filt.empty: 
        prepare_display_table(df_j_filt, is_field=True).to_excel(writer, index=False, sheet_name='Fält Data')

st.download_button("📥 Ladda ner Statistik (Excel)", output.getvalue(), "KLM_Statistik_Rapport.xlsx", "application/vnd.ms-excel")
