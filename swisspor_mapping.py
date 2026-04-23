import ifcopenshell
import ifcopenshell.guid
import csv, io, re
import pandas as pd
from datetime import datetime

# ============================================================
# KONFIGURATION & MAPS (Deine Original-Logik)
# ============================================================
PSET_IDENT  = "Pset_Swisspor_Identifikation"
PSET_PROD   = "Pset_Swisspor_Produktdaten"
PSET_TECH   = "Pset_Swisspor_TechnischeWerte"
PSET_KOSTEN = "Pset_Swisspor_Kostencodes"
PSET_GESAMT = "Pset_Swisspor_Gesamtaufbau"

FIELD_MAP_PROD = {
    "Produktname": "IfcLabel",
    "Material": "IfcLabel",
    "Dicke_mm": "IfcReal"
}
FIELD_MAP_TECH = {
    "Lambda_W_mK": "IfcReal",
    "Brandkennziffer": "IfcLabel",
    "Druckfestigkeit_kPa": "IfcInteger"
}
FIELD_MAP_KOSTEN = {
    "Baukostenplan_eBKP_H": "IfcLabel",
    "Nutzungsdauer_Jahre": "IfcInteger"
}
FIELD_MAP_GESAMT = {
    "Systemname": "IfcLabel",
    "Aufbaubeschreibung": "IfcText"
}

# Hilfsfunktion zum Schreiben von Properties (deine Original-Logik)
def write_prop(ifc, element, pset_name, prop_name, value, value_type="IfcLabel"):
    if value is None or str(value).strip() == "" or str(value).lower() == "nan":
        return False
    
    try:
        if value_type == "IfcReal": value = float(str(value).replace(',', '.'))
        elif value_type == "IfcInteger": value = int(float(str(value).replace(',', '.')))
    except: pass

    # Pset suchen oder erstellen
    psets = [p for p in element.IsDefinedBy if p.is_a("IfcRelDefinesByProperties") and p.RelatingPropertyDefinition.Name == pset_name]
    if psets:
        pset_def = psets[0].RelatingPropertyDefinition
        props = list(pset_def.HasProperties)
        for p in props:
            if p.Name == prop_name:
                p.NominalValue = ifc.create_entity(value_type, value)
                return True
        new_prop = ifc.create_entity("IfcPropertySingleValue", prop_name, None, ifc.create_entity(value_type, value), None)
        pset_def.HasProperties = props + [new_prop]
    else:
        new_prop = ifc.create_entity("IfcPropertySingleValue", prop_name, None, ifc.create_entity(value_type, value), None)
        pset_def = ifc.create_entity("IfcPropertySet", ifcopenshell.guid.new(), None, pset_name, None, [new_prop])
        ifc.create_entity("IfcRelDefinesByProperties", ifcopenshell.guid.new(), None, None, [element], pset_def)
    return True

# ============================================================
# DIE HAUPTFUNKTION FÜR STREAMLIT
# ============================================================
def process_mapping(ifc_file_raw, csv_files_list):
    report = []
    report.append(f"START MAPPING: {datetime.now().strftime('%H:%M:%S')}")
    
    # 1. IFC aus Streamlit laden
    ifc_bytes = ifc_file_raw.getvalue()
    ifc = ifcopenshell.file.from_string(ifc_bytes.decode("utf-8"))
    
    # 2. CSVs laden und Lookup erstellen
    lookup = {}
    for csv_file in csv_files_list:
        content = csv_file.getvalue().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(content))
        for row in reader:
            sid = str(row.get("SystemID", "")).strip().zfill(2)
            fun = str(row.get("Function", "")).strip()
            if sid != "00" and fun:
                lookup[(sid, fun)] = row

    # 3. MAPPING LOGIK (Deine Kern-Logik übernommen)
    all_elements = ifc.by_type("IfcElement")
    mapped_cnt = 0
    miss_cnt = 0
    gesamtdach = []

    for element in all_elements:
        # Identifikation suchen
        sid, function = None, None
        for rel in element.IsDefinedBy:
            if rel.is_a("IfcRelDefinesByProperties"):
                pset = rel.RelatingPropertyDefinition
                if pset.Name == PSET_IDENT:
                    for p in pset.HasProperties:
                        if p.Name == "SystemID": sid = str(p.NominalValue.wrappedValue).zfill(2)
                        if p.Name == "Function": function = str(p.NominalValue.wrappedValue)

        if not sid or not function: continue

        # FALL A: Modellierte Schicht
        if function != "Gesamtaufbau":
            data = lookup.get((sid, function))
            if data:
                mapped_cnt += 1
                for pname, typ in FIELD_MAP_PROD.items(): write_prop(ifc, element, PSET_PROD, pname, data.get(pname), typ)
                for pname, typ in FIELD_MAP_TECH.items(): write_prop(ifc, element, PSET_TECH, pname, data.get(pname), typ)
                for pname, typ in FIELD_MAP_KOSTEN.items(): write_prop(ifc, element, PSET_KOSTEN, pname, data.get(pname), typ)
            else:
                miss_cnt += 1
        
        # FALL B: Gesamtdachbauteil
        else:
            gesamtdach.append(element)

    # Nachbearbeitung Gesamtdach (vereinfacht für den Report)
    report.append(f"Erfolgreich gemappt: {mapped_cnt} Schichten")
    report.append(f"Nicht im CSV gefunden: {miss_cnt}")

    # Dummy Summary Table für Streamlit
    summary_df = pd.DataFrame([
        {"Kategorie": "Gemappte Schichten", "Anzahl": mapped_cnt},
        {"Kategorie": "Fehlende Einträge", "Anzahl": miss_cnt},
        {"Kategorie": "Gesamtdach-Elemente", "Anzahl": len(gesamtdach)}
    ])

    # 4. Rückgabe an app.py
    return {
        'new_ifc': ifc.to_string(),
        'report': "\n".join(report),
        'summary_table': summary_df
    }