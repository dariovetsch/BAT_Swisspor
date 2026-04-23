"""
============================================================
Swisspor IFC Mapping Script v5.0
Projekt: DC_BAT_FS26_Swisspor – Dario Vetsch, HSLU
============================================================

FIXES v5.0 (gegenüber v4.1):
  - Preise (Einheitspreis, Preisstand) NICHT mehr ins IFC gemappt
    → Preise kommen aus separater Preisdatenbank (kostenauswertung.py)
  - FUNCTION_PREFIX auf neue BIMQ-Codes aktualisiert (00–17)
  - FIELD_MAP_GESAMT: Einheitspreis entfernt

MAPPING-SCHLÜSSEL:  (SystemID, Function)
  Beide Properties in Pset_Swisspor_Identifikation

LOGIK:
  A) Modellierte Schichten → CSV-Properties direkt ins Element
     (Produktdaten, technische Werte, Kostencodes – KEIN Preis)
  B) Gesamtdachbauteil (Function=Gesamtaufbau) →
     Nicht modellierte Schichten als Pset_Swisspor_Gesamtaufbau
     + Schichtaufbau-String in Pset_Swisspor_Identifikation

VERWENDUNG:
  pip install ifcopenshell
  python swisspor_mapping_v5.py
============================================================
"""

import ifcopenshell
import ifcopenshell.guid
import csv, os, glob, re
from datetime import datetime

# ============================================================
# KONFIGURATION
# ============================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
IFC_INPUT   = os.path.join(BASE_DIR, "Modell_Architektur.ifc")
IFC_OUTPUT  = os.path.join(BASE_DIR, "Modell_Architektur_enriched.ifc")
REPORT_FILE = os.path.join(BASE_DIR, "mapping_report.txt")
CSV_PATTERN = os.path.join(BASE_DIR, "System_*.csv")

IFC_TYPES   = ["IfcRoof", "IfcSlab", "IfcBuildingElementProxy", "IfcCovering", "IfcPlate"]

# Pset-Namen
PSET_IDENT       = "Pset_Swisspor_Identifikation"
PSET_PRODUKT     = "Pset_Swisspor_Produkt"
PSET_FUNKTION    = "Pset_Swisspor_Funktion"   # für Material
PSET_TECHNISCH   = "Pset_Swisspor_Technisch"
PSET_KOSTEN      = "Pset_Swisspor_Kosten"
PSET_GESAMT      = "Pset_Swisspor_Gesamtaufbau"

# Kurzpräfix pro Funktion (für Gesamtdachbauteil) – BIMQ Codes 00–17
FUNCTION_PREFIX = {
    "00_Gesamtdachaufbau":       "Gesamtdachaufbau",
    "01_Unterkonstruktion":      "Unterkonstruktion",
    "02_Haftvermittler":         "Haftvermittler",
    "03_Dampfbremse":            "Dampfbremse",
    "04_Wärmedämmung":           "Waermedaemmung",
    "05_Gefälledämmung":         "Gefaelledaemmung",
    "06_Unterbahn":              "Unterbahn",
    "07_Oberbahn":               "Oberbahn",
    "08_Trennschicht":           "Trennschicht",
    "09_Drainageschicht":        "Drainageschicht",
    "10_Filterschicht":          "Filterschicht",
    "11_Wasserspeicherschicht":  "Wasserspeicherschicht",
    "12_Schutzschicht":          "Schutzschicht",
    "13_Splitschicht":           "Splitschicht",
    "14_Brandschutzschicht":     "Brandschutzschicht",
    "15_Vegetationstragschicht": "Vegetationstragschicht",
    "16_Belag":                  "Belag",
    "17_Absturzsicherung":       "Absturzsicherung",
}

# ============================================================
# FIELD MAP: CSV-Spalte → (Pset, IFC-Property, IFC-Typ)
# ============================================================
FIELD_MAP = [
    ("Produktbezeichnung",        PSET_PRODUKT,   "ProductName",                    "IfcText"),
    ("Produkttyp",                PSET_PRODUKT,   "ProductType",                    "IfcText"),
    ("Hersteller",                PSET_PRODUKT,   "Manufacturer",                   "IfcText"),
    ("Artikelnummer",             PSET_PRODUKT,   "ArticleNumber",                  "IfcLabel"),
    ("Lieferformat",              PSET_PRODUKT,   "Format",                         "IfcText"),
    ("LeistungserklärungNr",      PSET_PRODUKT,   "DeclarationOfPerformanceNumber", "IfcLabel"),
    ("Bezeichnung_nach_SIA",      PSET_PRODUKT,   "SIADesignation",                 "IfcText"),
    ("Anwendung_nach_SIA",        PSET_PRODUKT,   "SIAApplication",                 "IfcText"),
    ("Material",                  PSET_FUNKTION,  "Material",                       "IfcLabel"),
    ("Wärmeleitfähigkeit",        PSET_TECHNISCH, "ThermalConductivity",            "IfcReal"),
    ("Gefälle",                   PSET_TECHNISCH, "Slope",                          "IfcReal"),
    ("Mindestdicke",              PSET_TECHNISCH, "MinThickness",                   "IfcReal"),
    ("sd_Wert",                   PSET_TECHNISCH, "SdValue",                        "IfcReal"),
    ("Flächenbezogene_Masse",     PSET_TECHNISCH, "MassPerArea",                    "IfcReal"),
    ("Druckfestigkeit",           PSET_TECHNISCH, "CompressiveStrength",            "IfcReal"),
    ("Wasserspeichervermögen",    PSET_TECHNISCH, "WaterStorageCapacity",           "IfcReal"),
    ("Brandverhalten",            PSET_TECHNISCH, "ReactionToFire",                 "IfcLabel"),
    ("Brandverhaltensgruppe",     PSET_TECHNISCH, "FireClassificationCH",           "IfcLabel"),
    ("Spezifische_Wärmekapazität",PSET_TECHNISCH, "SpecificHeatCapacity",           "IfcReal"),
    ("Rohdichte",                 PSET_TECHNISCH, "Density",                        "IfcReal"),
    ("Diffusionswiderstandszahl", PSET_TECHNISCH, "VapourDiffusionResistanceFactor","IfcReal"),
    ("Wärmestandfestigkeit",      PSET_TECHNISCH, "HeatResistance",                 "IfcReal"),
    ("Kaltbiegeverhalten",        PSET_TECHNISCH, "FlexibilityAtLowTemperature",    "IfcReal"),
    ("Verbrauch",                 PSET_TECHNISCH, "ConsumptionPerArea",             "IfcReal"),
    ("MaximaleBenutzerzahl",      PSET_TECHNISCH, "MaximumUsers",                   "IfcInteger"),
    ("Überprüfungsintervall",     PSET_TECHNISCH, "InspectionInterval",             "IfcInteger"),
    ("MaximalePrüflast",          PSET_TECHNISCH, "MaximumTestLoad",                "IfcReal"),
    ("Kostencode",                PSET_KOSTEN,    "CostCode",                       "IfcLabel"),
    ("Kostenbezeichnung",         PSET_KOSTEN,    "CostLabel",                      "IfcText"),
    ("Preisbasis",                PSET_KOSTEN,    "PriceUnit",                      "IfcLabel"),
    # Einheitspreis und Preisstand werden NICHT ins IFC gemappt
    # → Preise kommen aus Swisspor_Preisdatenbank_2026.xlsx (kostenauswertung.py)
]

# Kurzversion für Gesamtaufbau-Pset (KEIN Einheitspreis)
FIELD_MAP_GESAMT = [
    ("Produktbezeichnung",   "Produktbezeichnung",   "IfcText"),
    ("Produkttyp",           "Produkttyp",           "IfcText"),
    ("Hersteller",           "Hersteller",           "IfcText"),
    ("Artikelnummer",        "Artikelnummer",        "IfcLabel"),
    ("Material",             "Material",             "IfcLabel"),
    ("Wärmeleitfähigkeit",   "Waermeleitfaehigkeit", "IfcReal"),
    ("sd_Wert",              "SdWert",               "IfcReal"),
    ("Brandverhalten",       "Brandverhalten",       "IfcLabel"),
    ("Brandverhaltensgruppe","Brandverhaltensgruppe","IfcLabel"),
    ("Preisbasis",           "Preisbasis",           "IfcLabel"),
    ("Kostencode",           "Kostencode",           "IfcLabel"),
    # Einheitspreis: NICHT ins IFC → kommt aus Preisdatenbank
]


# ============================================================
# KERN-FUNKTIONEN
# ============================================================

def has_value(v):
    """CSV-Wert schreibenswert?"""
    return bool(v and str(v).strip() not in ("", "0"))


def make_ifc_value(ifc, v, typ):
    try:
        if typ == "IfcReal":    return ifc.createIfcReal(float(v))
        if typ == "IfcInteger": return ifc.createIfcInteger(int(float(v)))
        if typ == "IfcText":    return ifc.createIfcText(str(v))
        if typ == "IfcLabel":   return ifc.createIfcLabel(str(v))
        return ifc.createIfcLabel(str(v))
    except Exception as e:
        print(f"    ⚠  Konvertierung '{v}' → {typ}: {e}")
        return None


def build_element_pset_map(ifc):
    """
    Baut direkt aus IfcRelDefinesByProperties eine Karte:
    element_id → {pset_name → {prop_name → wrappedValue}}
    
    Umgeht get_psets()-Bugs durch direkte Entity-Iteration.
    """
    elem_map = {}
    for rel in ifc.by_type("IfcRelDefinesByProperties"):
        pdef = rel.RelatingPropertyDefinition
        if not pdef.is_a("IfcPropertySet"):
            continue
        pset_name = pdef.Name
        props = {}
        for prop in pdef.HasProperties:
            if prop.is_a("IfcPropertySingleValue") and prop.NominalValue:
                props[prop.Name] = prop.NominalValue.wrappedValue
        for element in rel.RelatedObjects:
            eid = element.id()
            if eid not in elem_map:
                elem_map[eid] = {}
            elem_map[eid][pset_name] = props
    return elem_map


def get_pset_value_from_map(elem_map, element, pset_name, prop_name):
    return elem_map.get(element.id(), {}).get(pset_name, {}).get(prop_name)


def get_or_create_pset(ifc, element, pset_name):
    """Sucht oder erstellt einen PropertySet für ein Element."""
    for rel in element.IsDefinedBy:
        if rel.is_a("IfcRelDefinesByProperties"):
            pset = rel.RelatingPropertyDefinition
            if pset.is_a("IfcPropertySet") and pset.Name == pset_name:
                return pset
    pset = ifc.createIfcPropertySet(
        ifcopenshell.guid.new(), None, pset_name, None, []
    )
    ifc.createIfcRelDefinesByProperties(
        ifcopenshell.guid.new(), None, None, None, [element], pset
    )
    return pset


def write_prop(ifc, element, pset_name, prop_name, value, ifc_type):
    """Schreibt Property – überschreibt nur leere."""
    if not has_value(value):
        return False
    pset = get_or_create_pset(ifc, element, pset_name)
    props = list(pset.HasProperties)
    for prop in props:
        if prop.Name == prop_name:
            cur = prop.NominalValue.wrappedValue if prop.NominalValue else None
            if cur and str(cur).strip() not in ("", "0"):
                return False  # bereits gesetzt
            nv = make_ifc_value(ifc, value, ifc_type)
            if nv:
                prop.NominalValue = nv
                return True
            return False
    nv = make_ifc_value(ifc, value, ifc_type)
    if nv:
        np_ = ifc.createIfcPropertySingleValue(prop_name, None, nv, None)
        pset.HasProperties = props + [np_]
        return True
    return False


def write_all_from_row(ifc, element, row):
    """Schreibt alle FIELD_MAP Properties aus CSV-Zeile."""
    count = 0
    for csv_col, pset, prop, typ in FIELD_MAP:
        val = row.get(csv_col, "").strip()
        if write_prop(ifc, element, pset, prop, val, typ):
            count += 1
    return count


# ============================================================
# CSV LADEN
# ============================================================

def load_csvs(pattern):
    lookup = {}
    files = sorted(glob.glob(pattern))
    if not files:
        print(f"⚠  Keine CSV für Muster: {pattern}")
        return lookup
    for fp in files:
        with open(fp, encoding="utf-8-sig") as f:   # utf-8-sig removes Excel BOM
            lines = [l for l in f if not l.strip().startswith("#") and l.strip()]
        n = 0
        for row in csv.DictReader(lines):
            # Normalize SystemID: Excel entfernt führende Nullen → "1" → "01"
            sid_raw = row.get("SystemID", "").strip()
            sid = str(sid_raw).zfill(2) if sid_raw else ""
            fun = row.get("Function", "").strip()
            if sid and fun:
                lookup[(sid, fun)] = row
                n += 1
        print(f"  📄 {os.path.basename(fp)}: {n} Schichten")
    return lookup


# ============================================================
# HAUPTPROGRAMM
# ============================================================

def run(ifc_input, csv_pattern, ifc_output, report_path):
    print("=" * 60)
    print("Swisspor IFC Mapping v5.0")
    print(f"Start: {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 60)

    if not os.path.exists(ifc_input):
        print(f"❌ IFC nicht gefunden: {ifc_input}"); return

    print("\n📂 Lade CSV-Dateien...")
    lookup = load_csvs(csv_pattern)
    print(f"   Total: {len(lookup)} Schichten | SystemIDs: {sorted(set(k[0] for k in lookup))}")

    print(f"\n📐 Öffne IFC...")
    ifc = ifcopenshell.open(ifc_input)

    # Alle Elemente
    all_elements = []
    for t in IFC_TYPES:
        all_elements.extend(ifc.by_type(t))

    # Pset-Map aufbauen (robust, kein get_psets)
    print(f"🔍 Lese Psets aus {len(all_elements)} Elementen...")
    elem_pset_map = build_element_pset_map(ifc)

    # Klassifizieren
    normal, gesamtdach, sonstige = [], [], []
    for el in all_elements:
        sys_id = get_pset_value_from_map(elem_pset_map, el, PSET_IDENT, "SystemID")
        func   = get_pset_value_from_map(elem_pset_map, el, PSET_IDENT, "Function")
        if not sys_id and not func:
            sonstige.append(el)
        elif func and str(func).strip() == "Gesamtaufbau":
            # Normalize SystemID: "1" → "01"
            sid_norm = str(sys_id).strip().zfill(2) if sys_id else None
            gesamtdach.append((el, sid_norm))
        elif sys_id and not func:
            # SystemID vorhanden aber Function leer → überspringen mit Hinweis
            name = el.Name or "(kein Name)"
            print(f"  ⏭  Übersprungen: SystemID={sys_id}, Function leer [{name}]")
            sonstige.append(el)
        else:
            # Normalize SystemID: "1" → "01"
            sid = str(sys_id).strip().zfill(2) if sys_id else ""
            fn  = str(func).strip()            if func   else ""
            if sid and fn:
                normal.append((el, sid, fn))
            else:
                sonstige.append(el)

    print(f"\n📦 Klassifizierung:")
    print(f"   Modellierte Schichten:  {len(normal)}")
    print(f"   Gesamtdachbauteil:      {len(gesamtdach)}")
    print(f"   Ohne Swisspor-Pset:     {len(sonstige)}")

    report = [
        "Swisspor IFC Mapping Report v5.0",
        f"Datum:  {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"IFC:    {ifc_input}",
        f"Output: {ifc_output}",
        "=" * 60, "",
        "A) MODELLIERTE SCHICHTEN", "-" * 40,
    ]

    # --------------------------------------------------------
    # A) Modellierte Schichten
    # --------------------------------------------------------
    print(f"\n{'='*40}")
    print("A) Modellierte Schichten")
    print(f"{'='*40}")

    mapped_set  = set()
    mapped_cnt  = 0
    miss_cnt    = 0

    for element, sys_id, function in normal:
        key  = (sys_id, function)
        name = element.Name or "(kein Name)"

        if key not in lookup:
            print(f"  ⚠  Kein CSV: ({sys_id}, {function}) [{name}]")
            report.append(f"KEIN CSV | {sys_id} | {function} | {name}")
            miss_cnt += 1
            continue

        row   = lookup[key]
        count = write_all_from_row(ifc, element, row)
        prod  = row.get("Produktbezeichnung", "–")

        print(f"  ✅ SysID={sys_id} | {function:28} → {prod} [{count} Props]")
        report.append(f"OK | {sys_id} | {function} | {prod} | {count} Props")
        mapped_set.add(key)
        mapped_cnt += 1

    # --------------------------------------------------------
    # B) Gesamtdachbauteil
    # --------------------------------------------------------
    print(f"\n{'='*40}")
    print("B) Gesamtdachbauteil")
    print(f"{'='*40}")
    report += ["", "B) GESAMTDACHBAUTEIL", "-" * 40]

    for element, sys_id in gesamtdach:
        name = element.Name or "(kein Name)"
        if not sys_id:
            print(f"  ⚠  Kein SystemID [{name}]"); continue
        # sys_id ist bereits normalisiert (zfill(2) aus Klassifizierung)

        print(f"\n  🏗  [{name}] SystemID={sys_id}")
        report.append(f"\nGESAMT | {sys_id} | {name}")

        # Alle Schichten dieses Systems aus CSV
        sys_rows = {k[1]: v for k, v in lookup.items() if k[0] == sys_id}
        if not sys_rows:
            print(f"     ⚠  Keine CSV-Daten für SystemID={sys_id}"); continue

        aufbau = []
        gesamt_cnt = 0

        for function in sorted(sys_rows.keys()):
            row    = sys_rows[function]
            prod   = row.get("Produktbezeichnung", "").strip()
            prefix = FUNCTION_PREFIX.get(function, function.replace("_", ""))
            already = (sys_id, function) in mapped_set

            if prod:
                aufbau.append(f"{function}: {prod}")

            if already:
                print(f"     ✓  {function} → bereits modelliert")
                continue

            # Nicht modelliert → auf Gesamtdachbauteil
            print(f"     📋 {function} → {prod or '(kein Produkt)'}")
            n = 0
            for csv_col, prop_suffix, typ in FIELD_MAP_GESAMT:
                val = row.get(csv_col, "").strip()
                pname = f"{prefix}_{prop_suffix}"
                if write_prop(ifc, element, PSET_GESAMT, pname, val, typ):
                    n += 1
            print(f"        {n} Properties → {PSET_GESAMT}")
            report.append(f"  {function} | {prod} | {n} Props")
            gesamt_cnt += n

        # Schichtaufbau-String
        if aufbau:
            aufbau_str = " | ".join(aufbau)
            write_prop(ifc, element, PSET_IDENT, "Schichtaufbau", aufbau_str, "IfcText")
            print(f"\n     📝 Schichtaufbau: {len(aufbau)} Schichten")

        print(f"     Total Props Gesamtdachbauteil: {gesamt_cnt}")

    # Speichern
    ifc.write(ifc_output)

    # Report
    report += [
        "", "=" * 60, "ZUSAMMENFASSUNG",
        f"  Elemente total:              {len(all_elements)}",
        f"  Modellierte Schichten:       {mapped_cnt}",
        f"  Kein CSV-Eintrag:            {miss_cnt}",
        f"  Gesamtdachbauteil:           {len(gesamtdach)}",
        f"  Ohne Swisspor-Pset:          {len(sonstige)}",
        "", "NÄCHSTE SCHRITTE:",
        "  → swisspor_mengen.py  (Schritt 2)",
        "  → swisspor_kosten.py  (Schritt 3)",
    ]
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report))

    print(f"\n{'='*60}")
    print(f"💾 {ifc_output}")
    print(f"📄 {report_path}")
    print(f"✅ Fertig – {mapped_cnt} Schichten gemappt, {len(gesamtdach)} Gesamtdachbauteil")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    run(IFC_INPUT, CSV_PATTERN, IFC_OUTPUT, REPORT_FILE)
