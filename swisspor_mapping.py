"""
============================================================
swisspor_mapping.py  –  BIM-ENGINE-SP  v6.2
Projekt: DC_BAT_FS26_Swisspor – Dario Vetsch, HSLU
============================================================
v6.0 – Vollständige Überarbeitung gemäss BIMQ-Zuordnung
  Quelle: DC-BAT-FS26-Swisspor_Projektanforderungen_v14.xlsx

Änderungen gegenüber v5.1:
  FIELD_MAP (OT001-OT009 direkte Einzelschichten):
    + Preisstand -> Pset_Swisspor_Kosten.PriceDate         [fehlte komplett]
    + Deklaration -> Pset_Swisspor_Technisch.Declaration
    + Geradheit -> Pset_Swisspor_Technisch.Straightness
    + Wasserdichtheit -> Pset_Swisspor_Technisch.Watertightness
    + Höchstzugkraft_längs/quer -> TensileForceLongitudinal/Transverse
    + Höchstzugkraftdehnung_längs/quer -> ElongationAtMaxForce*
    + Sichtbare_Mängel -> VisibleDefects
    + Masshaltigkeit -> DimensionalStability
    + Widerstand_gegen_stossartige_Belastung -> ImpactResistance
    + Widerstand_gegen_statische_Belastung -> StaticLoadResistance
    + Widerstand_gegen_Durchwurzelung -> RootResistance
    + Bemessung_Nutzung_schwimmende_Estriche -> FloatingScreedCategory
    + Obere_Anwendungsgrenztemperatur_unbelastet -> MaximumServiceTemperatureUnloaded
    + Kriechverhalten_bei_Druckbeanspruchung -> CreepBehaviorUnderCompression
    + Kuenstliche_Alterung_bei_Dauerbeanspruchung -> ArtificialAgingDurability

  FIELD_MAP_GESAMT (OT010 Gesamtdachbauteil):
    * Produktbezeichnung -> ProductName (vorher: Product) – Konsistenzfix
    + Preisstand -> PriceDate                               [fehlte komplett]
    + Mindestdicke -> MinThickness                          [fehlte]
    + Gefälle -> Slope                                      [fehlte]
    + Rohdichte -> Density                                  [fehlte]
    + Spezifische_Wärmekapazität -> SpecificHeatCapacity    [fehlte]
    + Diffusionswiderstandszahl -> VapourDiffusionResistanceFactor [fehlte]
    + Alle neuen Technisch-Properties (identisch FIELD_MAP)
============================================================
"""

import ifcopenshell
import ifcopenshell.guid
import csv, io
import pandas as pd

# ── Pset-Namen ─────────────────────────────────────────────────────────────
PSET_IDENT     = "Pset_Swisspor_Identifikation"
PSET_PRODUKT   = "Pset_Swisspor_Produkt"
PSET_FUNKTION  = "Pset_Swisspor_Funktion"  # v6.2: BIMQ v15 definiert Material hier
PSET_TECHNISCH = "Pset_Swisspor_Technisch"
PSET_KOSTEN    = "Pset_Swisspor_Kosten"

# ── BIMQ: Function -> schichtspezifischer Pset-Name ────────────────────────
# Quelle: DC-BAT-FS26-Swisspor_Projektanforderungen_v14.xlsx (OT010)
FUNCTION_TO_PSET = {
    "01_Unterkonstruktion":      "Pset_Swisspor_Produkt",
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

# ── FIELD_MAP: CSV-Spalte -> (Pset, IFC-Property, IFC-Typ) ─────────────────
# Für direkte Einzelschicht-Elemente OT001–OT009.
# Alle Properties werden auf das jeweilige Element geschrieben;
# write_prop() überspringt leere/fehlende Werte automatisch.
#
# Quelle: DC-BAT-FS26-Swisspor_Projektanforderungen_v14.xlsx
#
FIELD_MAP = [
    # ── Produktdaten (Pset_Swisspor_Produkt) ────────────────────────────
    ("Produktbezeichnung",                      PSET_PRODUKT,   "ProductName",                          "IfcText"),
    ("Produkttyp",                              PSET_PRODUKT,   "ProductType",                          "IfcText"),
    ("Hersteller",                              PSET_PRODUKT,   "Manufacturer",                         "IfcText"),
    ("Artikelnummer",                           PSET_PRODUKT,   "ArticleNumber",                        "IfcLabel"),
    ("Lieferformat",                            PSET_PRODUKT,   "Format",                               "IfcText"),
    ("LeistungserklärungNr",                    PSET_PRODUKT,   "DeclarationOfPerformanceNumber",       "IfcLabel"),
    ("Bezeichnung_nach_SIA",                    PSET_PRODUKT,   "SIADesignation",                       "IfcText"),
    ("Anwendung_nach_SIA",                      PSET_PRODUKT,   "SIAApplication",                       "IfcText"),

    # ── Funktion (Pset_Swisspor_Funktion) ────────────────────────────────
    # Quelle: OD00401 — in Projektanforderungen für alle OT001-OT009 definiert
    ("Material",                                PSET_FUNKTION,  "Material",                             "IfcText"),  # v6.2: BIMQ v15 → Pset_Swisspor_Funktion

    # ── Technische Kennwerte (Pset_Swisspor_Technisch) ───────────────────
    ("Wärmeleitfähigkeit",                      PSET_TECHNISCH, "ThermalConductivity",                  "IfcReal"),
    ("Gefälle",                                 PSET_TECHNISCH, "Slope",                                "IfcReal"),
    ("Mindestdicke",                            PSET_TECHNISCH, "MinThickness",                         "IfcReal"),
    ("sd_Wert",                                 PSET_TECHNISCH, "SdValue",                              "IfcReal"),
    ("Flächenbezogene_Masse",                   PSET_TECHNISCH, "MassPerArea",                          "IfcReal"),
    ("Druckfestigkeit",                         PSET_TECHNISCH, "CompressiveStrength",                  "IfcReal"),
    ("Wasserspeichervermögen",                  PSET_TECHNISCH, "WaterStorageCapacity",                 "IfcReal"),
    ("Brandverhalten",                          PSET_TECHNISCH, "ReactionToFire",                       "IfcLabel"),
    ("Brandverhaltensgruppe",                   PSET_TECHNISCH, "FireClassificationCH",                 "IfcLabel"),
    ("Spezifische_Wärmekapazität",              PSET_TECHNISCH, "SpecificHeatCapacity",                 "IfcReal"),
    ("Rohdichte",                               PSET_TECHNISCH, "Density",                              "IfcReal"),
    ("Diffusionswiderstandszahl",               PSET_TECHNISCH, "VapourDiffusionResistanceFactor",      "IfcReal"),
    ("Wärmestandfestigkeit",                    PSET_TECHNISCH, "HeatResistance",                       "IfcReal"),
    ("Kaltbiegeverhalten",                      PSET_TECHNISCH, "FlexibilityAtLowTemperature",          "IfcReal"),
    ("Verbrauch",                               PSET_TECHNISCH, "ConsumptionPerArea",                   "IfcReal"),
    ("MaximaleBenutzerzahl",                    PSET_TECHNISCH, "MaximumUsers",                         "IfcInteger"),
    ("Überprüfungsintervall",                   PSET_TECHNISCH, "InspectionInterval",                   "IfcInteger"),
    ("MaximalePrüflast",                        PSET_TECHNISCH, "MaximumTestLoad",                      "IfcReal"),
    # Normative Prüfkennwerte (OD00420–OD00435)
    ("Deklaration",                             PSET_TECHNISCH, "Declaration",                          "IfcLabel"),
    ("Geradheit",                               PSET_TECHNISCH, "Straightness",                         "IfcLabel"),
    ("Wasserdichtheit",                         PSET_TECHNISCH, "Watertightness",                       "IfcLabel"),
    ("Höchstzugkraft_längs",                    PSET_TECHNISCH, "TensileForceLongitudinal",             "IfcReal"),
    ("Höchstzugkraft_quer",                     PSET_TECHNISCH, "TensileForceTransverse",               "IfcReal"),
    ("Höchstzugkraftdehnung_längs",             PSET_TECHNISCH, "ElongationAtMaxForceLongitudinal",     "IfcReal"),
    ("Sichtbare_Mängel",                        PSET_TECHNISCH, "VisibleDefects",                       "IfcText"),
    ("Widerstand_gegen_stossartige_Belastung",  PSET_TECHNISCH, "ImpactResistance",                     "IfcLabel"),
    ("Widerstand_gegen_statische_Belastung",    PSET_TECHNISCH, "StaticLoadResistance",                 "IfcLabel"),
    ("Widerstand_gegen_Durchwurzelung",         PSET_TECHNISCH, "RootResistance",                       "IfcLabel"),
    ("Masshaltigkeit",                          PSET_TECHNISCH, "DimensionalStability",                 "IfcLabel"),
    ("Bemessung_Nutzung_schwimmende_Estriche",  PSET_TECHNISCH, "FloatingScreedCategory",               "IfcLabel"),
    ("Obere_Anwendungsgrenztemperatur_unbelastet", PSET_TECHNISCH, "MaximumServiceTemperatureUnloaded", "IfcLabel"),
    ("Höchstzugkraftdehnung_quer",              PSET_TECHNISCH, "ElongationAtMaxForceTransverse",       "IfcReal"),
    ("Kriechverhalten_bei_Druckbeanspruchung",  PSET_TECHNISCH, "CreepBehaviorUnderCompression",        "IfcReal"),
    ("Kuenstliche_Alterung_bei_Dauerbeanspruchung", PSET_TECHNISCH, "ArtificialAgingDurability",        "IfcLabel"),

    # ── Kostendaten (Pset_Swisspor_Kosten) ───────────────────────────────
    ("Kostencode",                              PSET_KOSTEN,    "CostCode",                             "IfcLabel"),
    ("Kostenbezeichnung",                       PSET_KOSTEN,    "CostLabel",                            "IfcText"),
    ("Preisstand",                              PSET_KOSTEN,    "PriceDate",                            "IfcLabel"),  # ← NEU v6.0
    ("Preisbasis",                              PSET_KOSTEN,    "PriceUnit",                            "IfcLabel"),
    # Einheitspreis NICHT ins IFC -> kommt aus Swisspor_Preisdatenbank_2026.xlsx
]

# ── FIELD_MAP_GESAMT: CSV-Spalte -> (IFC-Property, IFC-Typ) ────────────────
# Für OT010 Gesamtdachbauteil — schreibt in schichtspezifische Psets
# (Pset-Name kommt aus FUNCTION_TO_PSET[function]).
# Identisch zu FIELD_MAP aber ohne Pset-Präfix, da Pset pro Schicht variiert.
#
# Quelle: DC-BAT-FS26-Swisspor_Projektanforderungen_v14.xlsx (OT010-Abschnitt)
#
FIELD_MAP_GESAMT = [
    # ── Produktdaten ─────────────────────────────────────────────────────
    ("Produktbezeichnung",                      "ProductName",                          "IfcText"),   # ← v6.0: Product->ProductName
    ("Produkttyp",                              "ProductType",                          "IfcText"),
    ("Hersteller",                              "Manufacturer",                         "IfcText"),
    ("Artikelnummer",                           "ArticleNumber",                        "IfcLabel"),
    ("Lieferformat",                            "Format",                               "IfcText"),
    ("LeistungserklärungNr",                    "DeclarationOfPerformanceNumber",       "IfcLabel"),
    ("Bezeichnung_nach_SIA",                    "SIADesignation",                       "IfcText"),
    ("Anwendung_nach_SIA",                      "SIAApplication",                       "IfcText"),

    # ── Technische Kennwerte ──────────────────────────────────────────────
    ("Material",                                "Material",                             "IfcText"),
    ("Wärmeleitfähigkeit",                      "ThermalConductivity",                  "IfcReal"),
    # v6.2: BIMQ v15 definiert CompressiveStressAt10PercentCompression (Dämmung).
    # Quelle-CSV hat nur Spalte "Druckspannung_bei_10_Stauchung" — wird gemappt falls vorhanden.
    ("Druckspannung_bei_10_Stauchung",          "CompressiveStressAt10PercentCompression", "IfcReal"),
    ("Gefälle",                                 "Slope",                                "IfcReal"),   # ← NEU v6.0
    ("Mindestdicke",                            "MinThickness",                         "IfcReal"),   # ← NEU v6.0
    ("sd_Wert",                                 "SdValue",                              "IfcReal"),
    ("Flächenbezogene_Masse",                   "MassPerArea",                          "IfcReal"),
    ("Druckfestigkeit",                         "CompressiveStrength",                  "IfcReal"),
    ("Wasserspeichervermögen",                  "WaterStorageCapacity",                 "IfcReal"),
    ("Brandverhalten",                          "ReactionToFire",                       "IfcLabel"),
    ("Brandverhaltensgruppe",                   "FireClassificationCH",                 "IfcLabel"),
    ("Spezifische_Wärmekapazität",              "SpecificHeatCapacity",                 "IfcReal"),   # ← NEU v6.0
    ("Rohdichte",                               "Density",                              "IfcReal"),   # ← NEU v6.0
    ("Diffusionswiderstandszahl",               "VapourDiffusionResistanceFactor",      "IfcReal"),   # ← NEU v6.0
    ("Wärmestandfestigkeit",                    "HeatResistance",                       "IfcReal"),
    ("Kaltbiegeverhalten",                      "FlexibilityAtLowTemperature",          "IfcReal"),
    ("Verbrauch",                               "ConsumptionPerArea",                   "IfcReal"),
    ("MaximaleBenutzerzahl",                    "MaximumUsers",                         "IfcInteger"),
    ("Überprüfungsintervall",                   "InspectionInterval",                   "IfcInteger"),
    ("MaximalePrüflast",                        "MaximumTestLoad",                      "IfcReal"),
    # Normative Prüfkennwerte
    ("Deklaration",                             "Declaration",                          "IfcLabel"),
    ("Geradheit",                               "Straightness",                         "IfcLabel"),
    ("Wasserdichtheit",                         "Watertightness",                       "IfcLabel"),
    ("Höchstzugkraft_längs",                    "TensileForceLongitudinal",             "IfcReal"),
    ("Höchstzugkraft_quer",                     "TensileForceTransverse",               "IfcReal"),
    ("Höchstzugkraftdehnung_längs",             "ElongationAtMaxForceLongitudinal",     "IfcReal"),
    ("Sichtbare_Mängel",                        "VisibleDefects",                       "IfcText"),
    ("Widerstand_gegen_stossartige_Belastung",  "ImpactResistance",                     "IfcLabel"),
    ("Widerstand_gegen_statische_Belastung",    "StaticLoadResistance",                 "IfcLabel"),
    ("Widerstand_gegen_Durchwurzelung",         "RootResistance",                       "IfcLabel"),
    ("Masshaltigkeit",                          "DimensionalStability",                 "IfcLabel"),
    ("Bemessung_Nutzung_schwimmende_Estriche",  "FloatingScreedCategory",               "IfcLabel"),
    ("Obere_Anwendungsgrenztemperatur_unbelastet", "MaximumServiceTemperatureUnloaded", "IfcLabel"),
    ("Höchstzugkraftdehnung_quer",              "ElongationAtMaxForceTransverse",       "IfcReal"),
    ("Kriechverhalten_bei_Druckbeanspruchung",  "CreepBehaviorUnderCompression",        "IfcReal"),
    ("Kuenstliche_Alterung_bei_Dauerbeanspruchung", "ArtificialAgingDurability",        "IfcLabel"),

    # ── Kostendaten ───────────────────────────────────────────────────────
    ("Kostencode",                              "CostCode",                             "IfcLabel"),
    ("Kostenbezeichnung",                       "CostLabel",                            "IfcText"),
    ("Preisstand",                              "PriceDate",                            "IfcLabel"),  # ← NEU v6.0
    ("Preisbasis",                              "PriceUnit",                            "IfcLabel"),
    # Einheitspreis NICHT ins IFC -> kommt aus Preisdatenbank
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
    Schreibt eine IFC-Property in das angegebene PropertySet.
    Überspringt leere, None oder Null-Werte.
    """
    if value is None or str(value).strip() in ("", "nan", "none", "None", "-", "0", "0.0"):
        return False

    # Typkonvertierung
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
    """Erkennt Trennzeichen anhand der ersten 10 Zeilen."""
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

        # Titelzeile überspringen
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

        logs.append(f"OK {csv_file.name}: {n} Schichten (Delimiter: '{delim}')")

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
    logs.append(f"OK IFC geladen: {ifc_file_raw.name}")

    # 2. CSVs laden
    lookup, csv_logs = load_csvs(csv_files_list)
    logs.extend(csv_logs)
    logs.append(f"INFO {len(lookup)} Schichten | SystemIDs: {sorted(set(k[0] for k in lookup))}")

    # 3. Mapping
    elements     = ifc.by_type("IfcElement")
    mapped_cnt   = missing_cnt = gesamt_cnt = skipped_cnt = 0
    mapped_set   = set()

    def get_swisspor_ident(el):
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

    # ── A) Direkte Einzelschichten (OT001–OT009) ──────────────────────────
    logs.append("\n─── A) Direkte Schichten ───")
    for el in elements:
        sid, function = get_swisspor_ident(el)
        if not (sid and function):
            skipped_cnt += 1
            continue
        if function in ("Gesamtaufbau", "00_Gesamtdachaufbau"):
            continue

        key = (sid, function)
        if key not in lookup:
            logs.append(f"  WARNUNG Kein CSV: ({sid}, {function}) [{el.Name or '?'}]")
            missing_cnt += 1
            continue

        d = lookup[key]
        count = sum(
            1 for col, pset, prop, typ in FIELD_MAP
            if write_prop(ifc, el, pset, prop, d.get(col), typ)
        )
        logs.append(f"  OK {sid} | {function:28} -> {d.get('Produktbezeichnung','?')} [{count} Props]")
        mapped_set.add(key)
        mapped_cnt += 1

    # ── B) Gesamtdachbauteil (OT010) ─────────────────────────────────────
    logs.append("\n─── B) Gesamtdachbauteil ───")
    for el in elements:
        sid, function = get_swisspor_ident(el)
        if not (sid and function):
            continue
        if function not in ("Gesamtaufbau", "00_Gesamtdachaufbau"):
            continue

        sys_rows = {k[1]: v for k, v in lookup.items() if k[0] == sid}
        if not sys_rows:
            logs.append(f"  WARNUNG Keine CSV für System {sid}")
            continue

        logs.append(f"\n  🏗️ [{el.Name or '?'}] System {sid}")
        aufbau  = []
        n_total = 0

        for func, row in sorted(sys_rows.items()):
            prod      = row.get("Produktbezeichnung", "").strip()
            pset_name = FUNCTION_TO_PSET.get(
                func,
                f"Pset_Swisspor_{func.split('_', 1)[-1] if '_' in func else func}"
            )
            if prod:
                aufbau.append(f"{func}: {prod}")

            n = sum(
                1 for col, prop, typ in FIELD_MAP_GESAMT
                if write_prop(ifc, el, pset_name, prop, row.get(col), typ)
            )
            already = " [auch direkt modelliert ✓]" if (sid, func) in mapped_set else ""
            logs.append(f"     INFO {func} -> {pset_name} [{n} Props]{already}")
            n_total   += n
            gesamt_cnt += 1

        if aufbau:
            write_prop(ifc, el, PSET_IDENT, "Schichtaufbau", " | ".join(aufbau), "IfcText")

        logs.append(f"     Total: {n_total} Props")

    # 4. Output
    ifc_bytes = ifc.to_string().encode("utf-8")

    summary_df = pd.DataFrame([
        {"Kategorie": "Schichten direkt gemappt",  "Anzahl": mapped_cnt},
        {"Kategorie": "🏗️ Gesamtaufbau-Schichten",    "Anzahl": gesamt_cnt},
        {"Kategorie": "Kein CSV-Eintrag",           "Anzahl": missing_cnt},
        {"Kategorie": "⏭ Übersprungen",               "Anzahl": skipped_cnt},
    ])

    return {
        "ifc_bytes":     ifc_bytes,
        "summary_table": summary_df,
        "log":           "\n".join(logs),
    }
