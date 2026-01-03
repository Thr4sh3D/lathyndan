import streamlit as st
import pandas as pd
import io

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Provstatistik", page_icon="🐕", layout="wide")

st.title("📊 Statistikverktyg för Kleiner Münsterländer (2023-2024)")
st.markdown("Ladda upp Excel-filen med flikarna **Eftersök** och **Jaktprov**.")

# --- 1. LADDA UPP EXCEL-FIL ---
uploaded_file = st.sidebar.file_uploader("Ladda upp Excel (.xlsx)", type=["xlsx"])

# Funktion för att tvätta kolumnnamn (tar bort mellanslag på slutet etc)
def clean_columns(df):
    df.columns = df.columns.str.strip()
    return df

if uploaded_file:
    try:
        # Läs in hela Excel-filen (alla flikar)
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Försök hitta rätt flikar automatiskt
        eftersok_sheet = next((s for s in sheet_names if "eftersök" in s.lower()), None)
        falt_sheet = next((s for s in sheet_names if "jakt" in s.lower() or "fält" in s.lower()), None)

        # --- DATAHANTERING ---
        df_eftersok = pd.DataFrame()
        df_falt = pd.DataFrame()

        if eftersok_sheet:
            df_eftersok = pd.read_excel(xls, sheet_name=eftersok_sheet)
            df_eftersok = clean_columns(df_eftersok)
            # Fixa datumformat
            if 'Datum' in df_eftersok.columns:
                df_eftersok['Datum'] = pd.to_datetime(df_eftersok['Datum'], errors='coerce')
                df_eftersok['År'] = df_eftersok['Datum'].dt.year
        
        if falt_sheet:
            df_falt = pd.read_excel(xls, sheet_name=falt_sheet)
            df_falt = clean_columns(df_falt)
            if 'Datum' in df_falt.columns:
                df_falt['Datum'] = pd.to_datetime(df_falt['Datum'], errors='coerce')
                df_falt['År'] = df_falt['Datum'].dt.year

        # Slå ihop allt för filter (för att hitta alla raser och år)
        all_data = pd.concat([df_eftersok, df_falt], ignore_index=True)
        
        # Hantera kolumnnamn för Ras (i din PDF heter den "rasnamn")
        ras_col = 'rasnamn' if 'rasnamn' in all_data.columns else 'Ras'
        if ras_col not in all_data.columns:
            st.error(f"Hittar inte kolumnen för Ras. Heter den 'rasnamn' i Excel-filen?")
            st.stop()

        # --- 2. SIDEBAR FILTER (GLOBALTA) ---
        st.sidebar.header("🔍 Filtrering")

        # A. Välj Ras
        alla_raser = sorted(all_data[ras_col].dropna().unique().astype(str))
        default_ras = ["Kleiner Münsterländer"] if "Kleiner Münsterländer" in alla_raser else alla_raser[:1]
        valda_raser = st.sidebar.multiselect("Välj Ras", alla_raser, default=default_ras)

        # B. Välj År
        if 'År' in all_data.columns:
            alla_ar = sorted(all_data['År'].dropna().unique().astype(int))
            valda_ar = st.sidebar.multiselect("Välj År", alla_ar, default=alla_ar)
        else:
            valda_ar = []

        # C. Välj Klass
        if 'Klass' in all_data.columns:
            alla_klasser = sorted(all_data['Klass'].dropna().astype(str).unique())
            valda_klasser = st.sidebar.multiselect("Välj Klass", alla_klasser, default=alla_klasser)
        else:
            valda_klasser = []

        # --- 3. FILTRERA DATAN ---
        def filter_data(df):
            if df.empty: return df
            temp_df = df.copy()
            # Filtrera Ras
            if valda_raser:
                temp_df = temp_df[temp_df[ras_col].isin(valda_raser)]
            # Filtrera År
            if valda_ar and 'År' in temp_df.columns:
                temp_df = temp_df[temp_df['År'].isin(valda_ar)]
            # Filtrera Klass
            if valda_klasser and 'Klass' in temp_df.columns:
                temp_df = temp_df[temp_df['Klass'].astype(str).isin(valda_klasser)]
            return temp_df

        df_e_filtered = filter_data(df_eftersok)
        df_f_filtered = filter_data(df_falt)

        # --- 4. VISA RESULTAT (FLIKAR) ---
        tab1, tab2 = st.tabs(["🌲 Eftersök (Vatten/Spår)", "🌾 Fältprov"])

        # === FLIK 1: EFTERSÖK ===
        with tab1:
            if df_e_filtered.empty:
                st.info("Ingen eftersöksdata hittades för detta urval.")
            else:
                st.subheader(f"Statistik Eftersök ({len(df_e_filtered)} starter)")
                
                # Identifiera kolumner för poäng (Spår/Vatten)
                # I din PDF heter de "Vatten" och "Spar" (eller Spår)
                vatten_col = 'Vatten'
                spar_col = 'Spår' if 'Spår' in df_e_filtered.columns else 'Spar'

                # Validera att kolumnerna finns
                if vatten_col in df_e_filtered.columns and spar_col in df_e_filtered.columns:
                    # RÄKNA GODKÄNDA (Minst 4 i båda grenar)
                    # Se till att det är siffror
                    df_e_filtered[vatten_col] = pd.to_numeric(df_e_filtered[vatten_col], errors='coerce').fillna(0)
                    df_e_filtered[spar_col] = pd.to_numeric(df_e_filtered[spar_col], errors='coerce').fillna(0)

                    godkanda = df_e_filtered[
                        (df_e_filtered[vatten_col] >= 4) & 
                        (df_e_filtered[spar_col] >= 4)
                    ]
                    
                    full_pott = df_e_filtered[
                        (df_e_filtered[vatten_col] == 10) & 
                        (df_e_filtered[spar_col] == 10)
                    ]

                    # KPI-KPI:er
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Antal Starter", len(df_e_filtered))
                    c2.metric("Unika Hundar", df_e_filtered['regnr'].nunique() if 'regnr' in df_e_filtered else 0)
                    c3.metric("Godkända", f"{len(godkanda)} st ({round(len(godkanda)/len(df_e_filtered)*100)}%)")
                    c4.metric("10-10 (Full pott)", len(full_pott))

                    st.divider()

                    # Detaljlista
                    st.markdown("#### Detaljerad lista")
                    visnings_cols = ['Datum', 'Klass', 'regnr', 'namn', vatten_col, spar_col]
                    # Filtrera så vi bara visar kolumner som faktiskt finns
                    visnings_cols = [c for c in visnings_cols if c in df_e_filtered.columns]
                    
                    st.dataframe(df_e_filtered[visnings_cols].sort_values('Datum', ascending=False), use_container_width=True)
                else:
                    st.warning(f"Kunde inte hitta poängkolumnerna 'Vatten' och '{spar_col}'. Kontrollera Excel-filen.")
                    st.write("Hittade kolumner:", df_e_filtered.columns.tolist())

        # === FLIK 2: FÄLTPROV ===
        with tab2:
            if df_f_filtered.empty:
                st.info("Ingen fältprovsdata hittades för detta urval (eller så heter fliken något annat än 'Jakt'/'Fält').")
            else:
                st.subheader(f"Statistik Fält ({len(df_f_filtered)} starter)")
                
                # Här visar vi bara rådata först eftersom jag inte vet exakta kolumnnamn för fältbetyg än
                st.dataframe(df_f_filtered, use_container_width=True)
                
                st.info("Tips: Om du vill ha specifik statistik för fält (t.ex. antal 1:a pris), berätta vad kolumnen heter som innehåller priset/poängen!")

        # --- EXPORT ---
        st.divider()
        st.subheader("📥 Ladda ner urval")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            if not df_e_filtered.empty:
                df_e_filtered.to_excel(writer, index=False, sheet_name='Eftersök Urval')
            if not df_f_filtered.empty:
                df_f_filtered.to_excel(writer, index=False, sheet_name='Fält Urval')
        
        st.download_button(
            label="Ladda ner Excel-fil med detta urval",
            data=output.getvalue(),
            file_name="KLM_Statistik_Urval.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Ett fel uppstod vid inläsning av filen: {e}")
        st.markdown("### Felsökning tips:")
        st.markdown("""
        1. Kontrollera att filen är en **.xlsx** (inte gammal .xls).
        2. Heter flikarna ungefär "Eftersök" och "Jaktprov/Fält"?
        3. Heter kolumnen för ras **rasnamn**?
        """)

else:
    st.info("👈 Börja med att ladda upp din Excel-fil i menyn till vänster.")
