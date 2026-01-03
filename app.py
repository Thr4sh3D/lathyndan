import streamlit as st
import pandas as pd
import io
import altair as alt  # <-- Nytt bibliotek för snyggare grafer

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
        # 1. Vi testar utf-8-sig först.
        text_data = content.decode("utf-8-sig")
    except:
        # 2. Fallback till latin-1
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
                    # Lägger till kritik om den finns
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
            
            # --- NY GRAF KOD MED SIFFROR PÅ STAPLARNA ---
            # Förbered data för grafen
            chart_data = df_klm['Resultat'].value_counts().reset_index()
            chart_data.columns = ['Resultat', 'Antal']
            
            # Skapa grafen
            base = alt.Chart(chart_data).encode(
                x=alt.X('Resultat', sort='-y', title="Resultat"),
                y=alt.Y('Antal', title="Antal starter")
            )

            # Staplarna
            bars = base.mark_bar().encode(
                color=alt.value("#4c78a8") # Du kan byta färg här om du vill
            )

            # Texten (Siffrorna ovanpå)
            text = base.mark_text(
                align='center',
                baseline='bottom',
                dy=-5,  # Flyttar texten lite uppåt från stapeln
                fontSize=14,
                color='white'  # Färgen på siffran (ändra till 'black' om du har ljus bakgrund)
            ).encode(
                text='Antal'
            )

            # Rita ut allt
            st.altair_chart(bars + text, use_container_width=True)
            # ---------------------------------------------

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
