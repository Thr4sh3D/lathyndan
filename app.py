import streamlit as st
import pandas as pd
import io
import re

# --- SID-INSTÄLLNINGAR ---
st.set_page_config(page_title="KLM Statistik", page_icon="🐕", layout="wide")

st.title("📊 Statistikverktyg för Kleiner Münsterländer")
st.markdown("Ladda upp resultatfilen (**Eftersök 2025.txt**) för att generera underlag till verksamhetsberättelsen.")

# --- 1. LADDA UPP FIL ---
uploaded_file = st.file_uploader("Välj textfil", type=["txt", "csv"])

if uploaded_file is not None:
    # Läs in filen
    content = uploaded_file.getvalue()
    try:
        # Försök läsa med latin-1 (vanligast för svenska Excel-exporter)
        text_data = content.decode("latin-1")
    except:
        # Fallback till utf-8 om latin-1 misslyckas
        text_data = content.decode("utf-8", errors="ignore")
    
    lines = text_data.split('\n')
    data = []
    
    # Hitta rubrikraden (oftast rad 1) för att veta var kolumnerna finns
    header = lines[0].split('\t')
    
    # Hjälpfunktion för att hitta kolumnindex (svarar -1 om den inte hittas)
    def get_col_index(headers, possible_names):
        for name in possible_names:
            for i, h in enumerate(headers):
                if name.lower() == h.strip().lower():
                    return i
        return -1

    # Identifiera viktiga kolumner dynamiskt
    idx_datum = get_col_index(header, ["Datum", "Provdatum"])
    idx_ras = get_col_index(header, ["rasnamn", "Ras", "Hundras"])
    idx_hund = get_col_index(header, ["namn", "Hund", "Hundnamn"])
    idx_regnr = get_col_index(header, ["regnr", "Reg.nr", "Registreringsnummer"])
    idx_klass = get_col_index(header, ["Klass", "Provklass"])
    idx_vatten = get_col_index(header, ["Vatten", "Vattenarbete"])
    idx_spar = get_col_index(header, ["Spar", "Spår", "Spårarbete"])
    idx_kritik = get_col_index(header, ["Kritik", "Domarkritik", "KritikVatten"]) # Tar första bästa kritik-kolumn

    # Kontroll: Hittade vi Vatten och Spår?
    if idx_vatten == -1 or idx_spar == -1:
        st.error(f"Kunde inte hitta kolumnerna för 'Vatten' och 'Spar/Spår'. Kontrollera filens rubriker. Hittade rubriker: {header}")
    else:
        # --- 2. REGELVERK (LOGIKEN) ---
        def bedom_eftersok(vatten, spar):
            if vatten < 4 or spar < 4:
                return "0 Pris (Ej Godkänd)"
            if vatten == 10 and spar == 10:
                return "Godkänd (Full pott 10-10)"
            return "Godkänd"

        # --- 3. TOLKA FILEN ---
        for line in lines[1:]: # Hoppa över rubrikraden
            parts = line.split('\t')
            
            # Kontrollera att raden har data
            if len(parts) > max(idx_vatten, idx_spar):
                # Hämta ras om kolumnen finns, annars anta att det är rätt fil
                ras = parts[idx_ras].strip() if idx_ras != -1 else "Okänd"
                
                # FILTER: VI SPARAR BARA KLEINER MÜNSTERLÄNDER (Om vi kan läsa rasen)
                if idx_ras == -1 or ("Kleiner" in ras or "Münsterländer" in ras or "kleiner" in ras): 
                    try:
                        raw_vatten = parts[idx_vatten]
                        raw_spar = parts[idx_spar]
                        
                        # Hantera "Eg" (Egenskapsbedömd/Ej godkänd) eller tomma värden som 0
                        poang_vatten = int(raw_vatten) if raw_vatten.strip().isdigit() else 0
                        poang_spar = int(raw_spar) if raw_spar.strip().isdigit() else 0
                    except:
                        poang_vatten = 0
                        poang_spar = 0

                    # Hämta övrig data
                    datum = parts[idx_datum] if idx_datum != -1 else "-"
                    hund = parts[idx_hund] if idx_hund != -1 else "-"
                    regnr = parts[idx_regnr] if idx_regnr != -1 else "-"
                    klass = parts[idx_klass] if idx_klass != -1 else "-"
                    kritik = parts[idx_kritik] if idx_kritik != -1 else ""

                    entry = {
                        "Datum": datum,
                        "Klass": klass,
                        "Hund": hund,
                        "Regnr": regnr,
                        "Vatten (Indata)": poang_vatten,
                        "Spår (Indata)": poang_spar,
                        "Resultat": bedom_eftersok(poang_vatten, poang_spar),
                        "Kritik": kritik
                    }
                    data.append(entry)

        if not data:
            st.warning("Inga hundar hittades. Kontrollera att filen innehåller 'Kleiner Münsterländer' om den har en ras-kolumn.")
        else:
            df_klm = pd.DataFrame(data)

            # --- 4. VISA RESULTATET ---
            tab1, tab2 = st.tabs(["📈 Statistik & Rapport", "🔍 Validering"])

            with tab1:
                st.success(f"Inläsningen klar! Hittade {len(df_klm)} starter.")
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
                st.dataframe(df_klm)
                
                st.markdown("### ⚠️ Varningar (0 poäng)")
                nollor = df_klm[(df_klm['Vatten (Indata)'] == 0) | (df_klm['Spår (Indata)'] == 0)]
                if not nollor.empty:
                    st.warning(f"Dessa {len(nollor)} rader har 0 i poäng (kontrollera om det är rätt):")
                    st.dataframe(nollor)
else:
    st.info("Väntar på filuppladdning...")
