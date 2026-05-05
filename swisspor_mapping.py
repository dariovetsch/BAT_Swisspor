"""
============================================================
swisspor_mapping.py  –  BIM-ENGINE-SP  v5.1
Projekt: DC_BAT_FS26_Swisspor – Dario Vetsch, HSLU
============================================================
FIXES v5.1:
  - Korrekte CSV-Spaltennamen (Produktbezeichnung, Material, ...)
  - Vollständiges FIELD_MAP (31 Felder) statt nur 5
  - write_prop(): NameError-Fix bei fehlendem Pset
  - CSV-Loading: Title-Row, BOM, Delimiter, SystemID-Normalisierung
  - Gesamtaufbau: '00_Gesamtdachaufbau' und 'Gesamtaufbau' erkannt
  - Download: bytes statt str (.encode())
  - Dezimaltrennzeichen: Komma → Punkt für IfcReal
============================================================
"""

import ifcopenshell
import ifcopenshell.guid
import csv, io
import pandas as pd

# ── Pset-Namen ─────────────────────────────────────────────────────────────
PSET_IDENT     = "Pset_Swisspor_Identifikation"
PSET_PRODUKT   = "Pset_Swisspor_Produkt"
PSET_FUNKTION  = "Pset_Swisspor_Funktion"
PSET_TECHNISCH = "Pset_Swisspor_Technisch"
PSET_KOSTEN    = "Pset_Swisspor_Kosten"
# PSET_GESAMT entfernt – jetzt schichtspezifische Psets nach BIMQ

# ── BIMQ: Function → schichtspezifischer Pset-Name ───────────────────────
# Quelle: DC-BAT-FS26-Swisspor_Vorlagen.xlsx (BIMQ-Export)
FUNCTION_TO_PSET = {
    "01_Unterkonstruktion":      "Pset_Swisspor_Produkt",          # kein eigener BIMQ-Pset
    "02_Haftvermittler":         "Pset_Swisspor_Haftvermittler",
    "03_Dampfbremse":            "Pset_Swisspor_Dampfbremse",
    "04_Wärmedämmung":           "Pset_Swisspor_Dämmung",
    "05_Gefälledämmung":         "Pset_Swisspor_Gefälledämmung",
    "06_Unterbahn":              "Pset_Swisspor_Unterbahn",
    "07_Oberbahn":               "Pset_Swisspor_Oberbahn",
    "08_Trennschicht":           "Pset_Swisspor_Trennschicht",
    "09_Drainageschicht":        "Pset_Swisspor_Drainageschicht",
    "10_Filterschicht":          "Pset_Swisspor_Filterschicht",
    "11_Wasserspeicherschicht":  "Pset_Swisspor_Wasserspeicherschicht",
    "12_Schutzschicht":          "Pset_Swisspor_Schutzschicht",
    "13_Splitschicht":           "Pset_Swisspor_Splitschicht",
    "14_Brandschutzschicht":     "Pset_Swisspor_Brandschutzschicht",
    "15_Vegetationstragschicht": "Pset_Swisspor_Vegetationstragschicht",
    "16_Belag":                  "Pset_Swisspor_Belag",
    "17_Absturzsicherung":       "Pset_Swisspor_Absturzsicherung",
}

# ── FIELD_MAP: CSV-Spalte → (Pset, IFC-Property, IFC-Typ) ─────────────────
FIELD_MAP = [
    ("Produktbezeichnung",         PSET_PRODUKT,   "ProductName",                     "IfcText"),
    ("Produkttyp",                 PSET_PRODUKT,   "ProductType",                     "IfcText"),
    ("Hersteller",                 PSET_PRODUKT,   "Manufacturer",                    "IfcText"),
    ("Artikelnummer",              PSET_PRODUKT,   "ArticleNumber",                   "IfcLabel"),
    ("Lieferformat",               PSET_PRODUKT,   "Format",                          "IfcText"),
    ("LeistungserklärungNr",       PSET_PRODUKT,   "DeclarationOfPerformanceNumber",  "IfcLabel"),
    ("Bezeichnung_nach_SIA",       PSET_PRODUKT,   "SIADesignation",                  "IfcText"),
    ("Anwendung_nach_SIA",         PSET_PRODUKT,   "SIAApplication",                  "IfcText"),
    ("Material",                   PSET_FUNKTION,  "Material",                        "IfcLabel"),
    ("Wärmeleitfähigkeit",         PSET_TECHNISCH, "ThermalConductivity",             "IfcReal"),
    ("Gefälle",                    PSET_TECHNISCH, "Slope",                           "IfcReal"),
    ("Mindestdicke",               PSET_TECHNISCH, "MinThickness",                    "IfcReal"),
    ("sd_Wert",                    PSET_TECHNISCH, "SdValue",                         "IfcReal"),
    ("Flächenbezogene_Masse",      PSET_TECHNISCH, "MassPerArea",                     "IfcReal"),
    ("Druckfestigkeit",            PSET_TECHNISCH, "CompressiveStrength",             "IfcReal"),
    ("Wasserspeichervermögen",     PSET_TECHNISCH, "WaterStorageCapacity",            "IfcReal"),
    ("Brandverhalten",             PSET_TECHNISCH, "ReactionToFire",                  "IfcLabel"),
    ("Brandverhaltensgruppe",      PSET_TECHNISCH, "FireClassificationCH",            "IfcLabel"),
    ("Spezifische_Wärmekapazität", PSET_TECHNISCH, "SpecificHeatCapacity",            "IfcReal"),
    ("Rohdichte",                  PSET_TECHNISCH, "Density",                         "IfcReal"),
    ("Diffusionswiderstandszahl",  PSET_TECHNISCH, "VapourDiffusionResistanceFactor", "IfcReal"),
    ("Wärmestandfestigkeit",       PSET_TECHNISCH, "HeatResistance",                  "IfcReal"),
    ("Kaltbiegeverhalten",         PSET_TECHNISCH, "FlexibilityAtLowTemperature",     "IfcReal"),
    ("Verbrauch",                  PSET_TECHNISCH, "ConsumptionPerArea",              "IfcReal"),
    ("MaximaleBenutzerzahl",       PSET_TECHNISCH, "MaximumUsers",                    "IfcInteger"),
    ("Überprüfungsintervall",      PSET_TECHNISCH, "InspectionInterval",              "IfcInteger"),
    ("MaximalePrüflast",           PSET_TECHNISCH, "MaximumTestLoad",                 "IfcReal"),
    ("Kostencode",                 PSET_KOSTEN,    "CostCode",                        "IfcLabel"),
    ("Kostenbezeichnung",          PSET_KOSTEN,    "CostLabel",                       "IfcText"),
    ("Preisbasis",                 PSET_KOSTEN,    "PriceUnit",                       "IfcLabel"),
    # Einheitspreis NICHT ins IFC → kommt aus Swisspor_Preisdatenbank_2026.xlsx
]

# BIMQ-konforme Properties pro Schicht-Pset (ohne Präfix im Property-Namen)
# Quelle: DC-BAT-FS26-Swisspor_Vorlagen.xlsx
FIELD_MAP_GESAMT = [
    # Produktdaten
    ("Produktbezeichnung",    "Product",                         "IfcText"),
    ("Produkttyp",            "ProductType",                     "IfcText"),
    ("Hersteller",            "Manufacturer",                    "IfcText"),
    ("Artikelnummer",         "ArticleNumber",                   "IfcLabel"),
    ("Lieferformat",          "Format",                          "IfcText"),
    ("LeistungserklärungNr",  "DeclarationOfPerformanceNumber",  "IfcLabel"),
    ("Bezeichnung_nach_SIA",  "SIADesignation",                  "IfcText"),
    ("Anwendung_nach_SIA",    "SIAApplication",                  "IfcText"),
    # Technische Kennwerte
    ("Material",              "Material",                        "IfcText"),
    ("Wärmeleitfähigkeit",    "ThermalConductivity",             "IfcReal"),
    ("sd_Wert",               "SdValue",                         "IfcReal"),
    ("Flächenbezogene_Masse", "MassPerArea",                     "IfcReal"),
    ("Brandverhalten",        "ReactionToFire",                  "IfcLabel"),
    ("Brandverhaltensgruppe", "FireClassificationCH",            "IfcLabel"),
    ("Wärmestandfestigkeit",  "HeatResistance",                  "IfcReal"),
    ("Kaltbiegeverhalten",    "FlexibilityAtLowTemperature",     "IfcReal"),
    ("Druckfestigkeit",       "CompressiveStrength",             "IfcReal"),
    ("Wasserspeichervermögen","WaterStorageCapacity",            "IfcReal"),
    ("Verbrauch",             "ConsumptionPerArea",              "IfcReal"),
    ("MaximaleBenutzerzahl",  "MaximumUsers",                    "IfcInteger"),
    ("Überprüfungsintervall", "InspectionInterval",              "IfcInteger"),
    ("MaximalePrüflast",      "MaximumTestLoad",                 "IfcReal"),
    # Kostendaten
    ("Kostencode",            "CostCode",                        "IfcLabel"),
    ("Kostenbezeichnung",     "CostLabel",                       "IfcText"),
    ("Preisbasis",            "PriceUnit",                       "IfcLabel"),
    # Einheitspreis NICHT ins IFC → kommt aus Preisdatenbank
]


# ══════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════

def _to_float(value):
    """Konvertiert '0,035' oder '0.035' sicher zu float."""
    try:
        return float(str(value).strip().replace(",", "."))
    except (ValueError, TypeError):
        return None


def write_prop(ifc, element, pset_name, prop_name, value, value_type="IfcLabel"):
    """
    Schreibt eine IFC-Property.
    FIX gegenüber Original: pset_def war bei neuem Pset undefiniert (NameError).
    """
    if value is None or str(value).strip() in ("", "nan", "none", "None", "-", "0", "0.0"):
        return False

    # Typkonvertierung mit Komma→Punkt
    try:
        if value_type == "IfcReal":
            v = _to_float(value)
            if v is None:
                return False
            value = v
        elif value_type == "IfcInteger":
            value = int(_to_float(value) or 0)
        else:
            value = str(value).strip()
    except Exception:
        return False

    # Bestehendes Pset suchen
    existing_pset = None
    for rel in element.IsDefinedBy:
        if rel.is_a("IfcRelDefinesByProperties"):
            pset = rel.RelatingPropertyDefinition
            if pset.is_a("IfcPropertySet") and pset.Name == pset_name:
                existing_pset = pset
                break

    if existing_pset:
        props = list(existing_pset.HasProperties)
        for p in props:
            if p.Name == prop_name:
                p.NominalValue = ifc.create_entity(value_type, value)
                return True
        new_prop = ifc.create_entity(
            "IfcPropertySingleValue", prop_name, None,
            ifc.create_entity(value_type, value), None
        )
        existing_pset.HasProperties = props + [new_prop]
    else:
        # FIX: neues Pset korrekt anlegen (kein undefiniertes pset_def)
        new_prop = ifc.create_entity(
            "IfcPropertySingleValue", prop_name, None,
            ifc.create_entity(value_type, value), None
        )
        new_pset = ifc.create_entity(
            "IfcPropertySet", ifcopenshell.guid.new(), None,
            pset_name, None, [new_prop]
        )
        ifc.create_entity(
            "IfcRelDefinesByProperties", ifcopenshell.guid.new(),
            None, None, None, [element], new_pset
        )
    return True


def _detect_delimiter(content):
    """Robust: zählt Trennzeichen in ersten 10 Zeilen."""
    sample = "\n".join(content.splitlines()[:10])
    counts = {",": sample.count(","), ";": sample.count(";"), "\t": sample.count("\t")}
    return max(counts, key=counts.get)


# ══════════════════════════════════════════════════════════════════════════
# CSV LOADING
# ══════════════════════════════════════════════════════════════════════════

def load_csvs(csv_files_list):
    """Lädt alle CSV-Uploads in ein Lookup-Dict {(SystemID, Function): row}."""
    lookup = {}
    logs   = []

    for csv_file in csv_files_list:
        raw = csv_file.getvalue()

        # Encoding erkennen
        content = None
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            try:
                content = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            logs.append(f"❌ {csv_file.name}: Encoding-Fehler")
            continue

        content = content.replace("\r\n", "\n").replace("\r", "\n")
        lines   = [l for l in content.splitlines() if l.strip()]

        # Titelzeile überspringen (System_02..04 haben extra Titelzeile)
        start_idx = 0
        for i, line in enumerate(lines):
            if "SystemID" in line and "Function" in line:
                start_idx = i
                break
        lines = lines[start_idx:]

        if not lines:
            logs.append(f"❌ {csv_file.name}: Kein Header gefunden")
            continue

        delim = _detect_delimiter("\n".join(lines[:5]))
        n = 0

        for row in csv.DictReader(lines, delimiter=delim):
            clean = {str(k).strip().lstrip("\ufeff"): str(v).strip()
                     for k, v in row.items() if k}

            sid_raw = clean.get("SystemID", "").strip()
            try:
                sid = str(int(float(sid_raw))).zfill(2) if sid_raw else ""
            except ValueError:
                sid = sid_raw.zfill(2) if sid_raw else ""

            fun = clean.get("Function", "").strip()
            if sid and sid != "00" and fun:
                lookup[(sid, fun)] = clean
                n += 1

        logs.append(f"✅ {csv_file.name}: {n} Schichten (Delimiter: '{delim}')")

    return lookup, logs


# ══════════════════════════════════════════════════════════════════════════
# MAIN MAPPING
# ══════════════════════════════════════════════════════════════════════════

def process_mapping(ifc_file_raw, csv_files_list):
    """
    Hauptfunktion für Streamlit.
    Gibt dict zurück: {ifc_bytes, summary_table, log}
    """
    logs = []

    # 1. IFC laden
    ifc_content = ifc_file_raw.getvalue().decode("utf-8", errors="replace")
    ifc = ifcopenshell.file.from_string(ifc_content)
    logs.append(f"✅ IFC geladen: {ifc_file_raw.name}")

    # 2. CSVs laden
    lookup, csv_logs = load_csvs(csv_files_list)
    logs.extend(csv_logs)
    logs.append(f"📊 {len(lookup)} Schichten | SystemIDs: {sorted(set(k[0] for k in lookup))}")

    # 3. Mapping
    elements    = ifc.by_type("IfcElement")
    mapped_cnt  = missing_cnt = gesamt_cnt = skipped_cnt = 0
    mapped_set  = set()

    def get_swisspor_ident(el):
        """Liest SystemID und Function aus Pset_Swisspor_Identifikation."""
        for rel in el.IsDefinedBy:
            if not rel.is_a("IfcRelDefinesByProperties"):
                continue
            pset = rel.RelatingPropertyDefinition
            if not pset.is_a("IfcPropertySet"):
                continue
            if (pset.Name or "").lower() != "pset_swisspor_identifikation":
                continue
            sid, func = None, None
            for p in pset.HasProperties:
                if not p.NominalValue:
                    continue
                pn  = (p.Name or "").lower()
                pv  = str(p.NominalValue.wrappedValue).strip()
                if "systemid" in pn:
                    try:
                        sid = str(int(float(pv))).zfill(2) if pv else ""
                    except ValueError:
                        sid = pv.zfill(2) if pv else ""
                if pn == "function":
                    func = pv
            return sid, func
        return None, None

    logs.append("\n─── A) Direkte Schichten ───")
    for el in elements:
        sid, function = get_swisspor_ident(el)
        if not sid and not function:
            skipped_cnt += 1
            continue
        if not function:
            skipped_cnt += 1
            continue
        if function in ("Gesamtaufbau", "00_Gesamtdachaufbau"):
            continue  # → Case B

        key = (sid, function)
        if key not in lookup:
            logs.append(f"  ⚠️ Kein CSV: ({sid}, {function}) [{el.Name or '?'}]")
            missing_cnt += 1
            continue

        d = lookup[key]
        count = sum(
            1 for col, pset, prop, typ in FIELD_MAP
            if write_prop(ifc, el, pset, prop, d.get(col), typ)
        )
        logs.append(f"  ✅ {sid} | {function:28} → {d.get('Produktbezeichnung','?')} [{count} Props]")
        mapped_set.add(key)
        mapped_cnt += 1

    logs.append("\n─── B) Gesamtdachbauteil ───")
    for el in elements:
        sid, function = get_swisspor_ident(el)
        if not (sid and function):
            continue
        if function not in ("Gesamtaufbau", "00_Gesamtdachaufbau"):
            continue

        sys_rows = {k[1]: v for k, v in lookup.items() if k[0] == sid}
        if not sys_rows:
            logs.append(f"  ⚠️ Keine CSV für System {sid}")
            continue

        logs.append(f"\n  🏗️ [{el.Name or '?'}] System {sid}")
        aufbau  = []
        n_total = 0

        for func, row in sorted(sys_rows.items()):
            prod    = row.get("Produktbezeichnung", "").strip()
            # BIMQ: schichtspezifischer Pset-Name (z.B. Pset_Swisspor_Unterbahn)
            pset_name = FUNCTION_TO_PSET.get(func, f"Pset_Swisspor_{func.split('_',1)[-1] if '_' in func else func}")
            if prod:
                aufbau.append(f"{func}: {prod}")

            # WICHTIG: Alle Schichten immer auf Gesamtbauteil schreiben –
            # auch wenn die Schicht einzeln modelliert ist.
            # Grund: Kostenauswertung liest Artikelnummern vom Gesamtbauteil.
            # So funktioniert UC-02 auch bei teilweiser Modellierung.
            n = sum(
                1 for col, prop, typ in FIELD_MAP_GESAMT
                if write_prop(ifc, el, pset_name, prop, row.get(col), typ)
            )
            already = " [auch modelliert ✓]" if (sid, func) in mapped_set else ""
            logs.append(f"     📋 {func} → {pset_name} [{n} Props]{already}")
            n_total   += n
            gesamt_cnt += 1

        if aufbau:
            write_prop(ifc, el, PSET_IDENT, "Schichtaufbau", " | ".join(aufbau), "IfcText")

        logs.append(f"     Total: {n_total} Props")

    # 4. Output
    # FIX: .encode() → bytes für st.download_button
    ifc_bytes = ifc.to_string().encode("utf-8")

    summary_df = pd.DataFrame([
        {"Kategorie": "✅ Schichten direkt gemappt",  "Anzahl": mapped_cnt},
        {"Kategorie": "🏗️ Gesamtaufbau-Schichten",    "Anzahl": gesamt_cnt},
        {"Kategorie": "⚠️ Kein CSV-Eintrag",          "Anzahl": missing_cnt},
        {"Kategorie": "⏭ Übersprungen",               "Anzahl": skipped_cnt},
    ])

    return {
        "ifc_bytes":     ifc_bytes,
        "summary_table": summary_df,
        "log":           "\n".join(logs),
    }
