import streamlit as st
import pandas as pd
import io

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Statistik", page_icon="🐕", layout="wide")

st.title("📊 Statistikverktyg för Kleiner Münsterländer")
st.markdown("Ladda upp resultatfilen (**Eftersök 2025.txt**) för att generera underlag till verksamhetsberättelsen.")

# --- 1. LADDA UPP FIL ---
uploaded_file = st.file_uploader("Välj textfil", type=["txt", "csv"])

if uploaded_file is not None:
    # Läs in filen
    content = uploaded_file.getvalue()
    
    # --- FIX FÖR SVENSKA TECKEN (ÅÄÖ) ---
    try:
        # 1. Vi testar utf-8-sig först. Det fixar både ÅÄÖ och de konstiga tecknen i början av filen.
        text_data = content.decode("utf-8-sig")
    except:
        # 2. Om det misslyckas, testa det äldre Windows-formatet (latin-1)
        text_data = content.decode("latin-1")
    
    lines = text_data.split('\n')
    data = []

    # --- 2. REGELVERK ---
    def bedom_eftersok(vatten, spar):
        if vatten < 4 or spar < 4:
            return "0 Pris (Ej Godkänd)"
        if vatten == 10 and spar == 10:
            return "Godkänd (Full pott 10-10)"
        return "Godkänd"

    # --- 3. TOLKA FILEN ---
    for line in lines:
        parts = line.split('\t')
        
        # Vi behöver rader som är tillräckligt långa
        if len(parts) > 40:
            ras = parts[4].strip()
            
            # FILTER: VI SPARAR BARA KLEINER MÜNSTERLÄNDER
            if "Kleiner" in ras or "Münsterländer" in ras or "kleiner" in ras: 
                try:
                    # Vi läser kolumn 36 (Vatten) och 37 (Spår)
                    raw_vatten = parts[36].strip()
                    raw_spar = parts[37].strip()
                    
                    poang_vatten = int(raw_vatten) if raw_vatten.isdigit() else 0
                    poang_spar = int(raw_spar) if raw_spar.isdigit() else 0
                except:
                    poang_vatten = 0
                    poang_spar = 0

                # Skapa datapunkten
                entry = {
                    "Datum": parts[0],      
                    "Klass": parts[1],      
                    "Hund": parts[2],       
                    "Regnr": parts[3],      
                    "Vatten (Indata)": poang_vatten, # Kolumn 36
                    "Spår (Indata)": poang_spar,     # Kolumn 37
                    "Resultat": bedom_eftersok(poang_vatten, poang_spar),
                    "Ras": ras,
                    "Kritik": f"Vatten: {parts[66]} Spår: {parts[67]}" if len(parts) > 67 else "" 
                }
                data.append(entry)

    if not data:
        st.error("Inga hundar av rasen Kleiner Münsterländer hittades.")
    else:
        df_klm = pd.DataFrame(data)

        # --- 4. VISA RESULTATET ---
        tab1, tab2 = st.tabs(["📈 Statistik & Rapport", "🔍 Validering"])

        with tab1:
            st.success(f"Inläsningen klar! Hittade {len(df_klm)} starter för KLM.")
            st.divider()
            
            col1, col2 = st.columns(2)
            col1.metric("Antal starter", len(df_klm))
            col2.metric("Unika individer", df_klm['Regnr'].nunique())
            
            st.divider()
            st.subheader("Resultatfördelning")
            st.bar_chart(df_klm['Resultat'].value_counts())

            # Excel-export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_klm.to_excel(writer, index=False, sheet_name='Data 2025')
            
            st.download_button("📥 Ladda ner Excel", output.getvalue(), "KLM_Statistik_2025.xlsx", "application/vnd.ms-excel")

        with tab2:
            st.markdown("### 🕵️‍♀️ Validering")
            st.dataframe(df_klm)
            
            # Varningar för 0:or
            nollor = df_klm[(df_klm['Vatten (Indata)'] == 0) | (df_klm['Spår (Indata)'] == 0)]
            if not nollor.empty:
                st.warning(f"Dessa rader har 0 poäng (kolla så det stämmer):")
                st.dataframe(nollor)

else:
    st.info("Väntar på filuppladdning...")
