import streamlit as st
import mapping_logic  # Importiert deine Logik aus der anderen Datei
import pandas as pd

st.title("POC: Swisspor IFC-Mapping & Kosten")

# 1. INPUT PHASE
st.header("1. Upload")
ifc_file = st.file_uploader("IFC Modell hochladen", type=['ifc'])
csv_file = st.file_uploader("Swisspor Referenzdaten (CSV) hochladen", type=['csv'])

if ifc_file and csv_file:
    # 2. VERARBEITUNG (Aufruf deiner Logik)
    st.header("2. Mapping & Analyse")
    with st.spinner("Berechne Mapping und Kosten..."):
        # Hier wird die Funktion aus deinem anderen Skript aufgerufen
        results = mapping_logic.process_mapping(ifc_file, csv_file)
    
    # 3. OUTPUT (Anzeige der Ergebnisse)
    st.success("Mapping abgeschlossen!")
    st.dataframe(results['summary_table']) # Kostentabelle anzeigen
    
    # Download Button für das neue IFC
    st.download_button("Neues IFC herunterladen", data=results['new_ifc'], file_name="gemappt_swisspor.ifc")