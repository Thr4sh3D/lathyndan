import streamlit as st
import pandas as pd
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Statistik", page_icon="🐕", layout="wide")

st.title("📊 Statistikverktyg för Kleiner Münsterländer")
st.markdown("Ladda upp resultatfilen (**Eftersök 2025.txt**) för att generera underlag till verksamhetsberättelsen.")

# --- 1. LADDA UPP FIL ---
uploaded_file = st.file_uploader("Välj textfil", type=["txt"])

if uploaded_file is not None:
    # Läs in filen (hanterar svenska tecken)
    content = uploaded_file.getvalue()
    try:
        text_data = content.decode("latin-1")
    except:
        text_data = content.decode("utf-8", errors="ignore")
    
    lines = text_data.split('\n')
    data = []

    # --- 2. REGELVERK (LOGIKEN) ---
    def bedom_eftersok(vatten, spar):
        # Enligt § 8: Lägst betyg 4 krävs i samtliga grenar för pris
        if vatten < 4 or spar < 4:
            return "0 Pris (Ej Godkänd)"
        if vatten == 10 and spar == 10:
            return "Godkänd (Full pott 10-10)"
        return "Godkänd"

    # --- 3. TOLKA FILEN ---
    for line in lines:
        # Vi letar efter rader som börjar med ett datum (t.ex. 25-05-15)
        if re.match(r'^\d{2}-\d{2}-\d{2}', line):
            parts = line.split('\t')
            
            # Grovsortering: Är raden tillräckligt lång?
            if len(parts) > 30:
                ras = parts[4].strip()
                
                # FILTER: VI SPARAR BARA KLEINER MÜNSTERLÄNDER
                if "Kleiner" in ras or "Münsterländer" in ras: 
                    try:
                        # Hämta poängen. I din fil ligger de ofta på index 27 (Vatten) och 28 (Spår).
                        # Vi gör det robust genom att tvinga till heltal.
                        raw_vatten = parts[27]
                        raw_spar = parts[28]
                        
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
                        "Vatten (Indata)": poang_vatten,
                        "Spår (Indata)": poang_spar,
                        "Resultat": bedom_eftersok(poang_vatten, poang_spar),
                        "Kritik": " ".join(parts[40:]).strip() # Slår ihop kritiken på slutet
                    }
                    data.append(entry)

    if not data:
        st.error("Inga hundar av rasen Kleiner Münsterländer hittades. Är det rätt fil?")
    else:
        df_klm = pd.DataFrame(data)

        # --- 4. VISA RESULTATET (FLIKAR) ---
        tab1, tab2 = st.tabs(["📈 Statistik & Rapport", "🔍 Validering (Kontrollera data)"])

        # --- FLIK 1: STATISTIK ---
        with tab1:
            st.success(f"Inläsningen klar! Hittade {len(df_klm)} starter för KLM.")
            st.divider()
            
            # Siffror till Tabell 2 i Word-dokumentet
            st.subheader("Underlag till Tabell 2 (Verksamhetsberättelsen)")
            col1, col2 = st.columns(2)
            col1.metric("Antal starter (Eftersök)", len(df_klm))
            col2.metric("Unika individer (Eftersök)", df_klm['Regnr'].nunique())
            
            st.info("Ovanstående siffror fyller du i för år 2025 i Tabell 2.")
            
            st.divider()

            # Resultatfördelning (Till Tabell 3)
            st.subheader("Resultatfördelning")
            resultat_stats = df_klm['Resultat'].value_counts()
            
            c1, c2 = st.columns([1, 2])
            with c1:
                st.dataframe(resultat_stats, use_container_width=True)
            with c2:
                st.bar_chart(resultat_stats)

            # Ladda ner Excel
            st.subheader("Exportera")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_klm.to_excel(writer, index=False, sheet_name='Data 2025')
            
            st.download_button(
                label="📥 Ladda ner färdig Excel-fil",
                data=output.getvalue(),
                file_name="KLM_Statistik_2025.xlsx",
                mime="application/vnd.ms-excel"
            )

        # --- FLIK 2: VALIDERING ---
        with tab2:
            st.markdown("### 🕵️‍♀️ Validering")
            st.markdown("""
            Använd denna lista för att stickprovskontrollera att appen läst rätt siffror.
            1. Öppna din **original-textfil**.
            2. Jämför en hund i listan nedan mot textfilen.
            3. Stämmer kolumnerna **Vatten** och **Spår**?
            """)

            # Visa hela tabellen sökbar
            st.dataframe(df_klm, use_container_width=True)

            st.divider()
            st.markdown("### ⚠️ Varningar (Hundar med 0 poäng)")
            nollor = df_klm[(df_klm['Vatten (Indata)'] == 0) | (df_klm['Spår (Indata)'] == 0)]
            
            if not nollor.empty:
                st.warning(f"Hittade {len(nollor)} starter med 0 poäng. Kontrollera att dessa stämmer:")
                st.dataframe(nollor)
            else:
                st.success("Inga nollor hittades i poängkolumnerna.")

else:
    st.info("Väntar på filuppladdning...")
