import streamlit as st
import swisspor_mapping
import swisspor_kosten

st.set_page_config(
    page_title="swissporBIM",
    page_icon="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 60 40'><text y='28' font-size='18' font-family='Helvetica Neue,sans-serif' font-weight='700' fill='%231a3a5c'>sw</text><circle cx='44' cy='22' r='5' fill='%231a3a5c'/><circle cx='52' cy='22' r='5' fill='%23f0a500'/><circle cx='60' cy='22' r='5' fill='%23e63946'/></svg>",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── NAV_PAD muss mit --nav-pad CSS-Variable übereinstimmen ───────────────────
NAV_PAD = "4rem"

st.markdown(f"""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">

<style>
:root {{
    --sw-navy:   #1a3a5c;
    --sw-red:    #e63946;
    --sw-orange: #f0a500;
    --sw-bg:     #f0f2f5;
    --sw-white:  #ffffff;
    --sw-border: #e0e0e0;
    --sw-text:   #1a1a1a;
    --sw-muted:  #6b7280;
    --font:      'Source Sans 3', 'Helvetica Neue', Helvetica, Arial, sans-serif;
    --mono:      'DM Mono', monospace;
    --nav-pad:   {NAV_PAD};
}}

/* ── Streamlit boilerplate ausblenden ── */
#MainMenu, footer, header, .stDeployButton,
[data-testid="collapsedControl"] {{ display: none !important; }}
section[data-testid="stSidebar"] {{ display: none !important; }}

html, body, .stApp {{
    background: var(--sw-bg) !important;
    font-family: var(--font) !important;
    color: var(--sw-text) !important;
}}

/* ── KERN: block-container bekommt die Seitenränder ── */
/* Das trifft ALLE Streamlit-Widgets: Uploader, Buttons, Columns usw. */
.block-container {{
    padding: 0 {NAV_PAD} 4rem {NAV_PAD} !important;
    max-width: 1280px !important;
    margin: 0 auto !important;
}}

/* ── Topnav: bricht aus dem block-container aus ── */
/* Negative Margins + volle Breite mit vw-Trick */
.sw-nav {{
    background: var(--sw-white);
    border-bottom: 3px solid var(--sw-red);
    /* Aus dem Container herausbrechen */
    margin-left:  calc(-1 * var(--nav-pad));
    margin-right: calc(-1 * var(--nav-pad));
    padding: 0 var(--nav-pad);
    width: calc(100% + 2 * var(--nav-pad));
    height: 60px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 200;
    margin-bottom: 1.5rem;
}}
.sw-logo {{ display: flex; align-items: center; }}
.sw-logo-text {{ font-size: 1.35rem; font-weight: 700; color: var(--sw-navy);
    letter-spacing: -0.01em; line-height: 1; }}
.sw-logo-dots {{ display: flex; gap: 4px; margin-left: 6px; margin-top: 2px; }}
.sw-logo-dots span {{ width: 9px; height: 9px; border-radius: 50%; display: inline-block; }}
.sw-nav-title {{ font-size: 0.95rem; color: var(--sw-navy); font-weight: 400; }}
.sw-nav-right {{ font-size: 0.78rem; color: var(--sw-muted); font-family: var(--mono);
    letter-spacing: 0.05em; }}

/* ── Tabs ── */
[data-testid="stTabs"] [role="tablist"] {{
    background: transparent !important; border-bottom: 2px solid var(--sw-border) !important;
    gap: 0 !important; padding: 0 !important; }}
[data-testid="stTabs"] [role="tab"] {{
    font-family: var(--font) !important; font-size: 0.78rem !important;
    font-weight: 700 !important; letter-spacing: 0.1em !important;
    text-transform: uppercase !important; color: var(--sw-muted) !important;
    padding: 0.7rem 1.6rem !important; border: none !important;
    border-bottom: 2px solid transparent !important; border-radius: 0 !important;
    background: transparent !important; margin-bottom: -2px !important; }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: var(--sw-navy) !important; border-bottom-color: var(--sw-red) !important;
    background: transparent !important; }}
[data-testid="stTabsContent"] {{ padding-top: 1.25rem !important; border: none !important; }}

/* ── Cards ── */
.sw-card {{
    background: var(--sw-white);
    border-radius: 6px;
    padding: 1.5rem 1.75rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}

/* ── Section headers ── */
.sw-section-header {{ display: flex; align-items: center; gap: 12px; margin-bottom: 0.9rem; }}
.sw-section-num {{ width: 26px; height: 26px; border-radius: 50%;
    background: var(--sw-navy); color: white; font-size: 0.72rem; font-weight: 700;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0; }}
.sw-section-title {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--sw-navy); }}
.sw-upload-hint {{ font-size: 0.82rem; color: var(--sw-muted);
    margin-bottom: 1rem; line-height: 1.4; }}

.sw-file-badge {{ display: inline-flex; align-items: center; gap: 5px;
    color: var(--sw-navy); border: 1px solid var(--sw-navy);
    font-size: 0.75rem; font-weight: 600; padding: 3px 10px;
    border-radius: 2px; margin: 3px 5px 3px 0; }}

/* ── Buttons ── */
.stButton > button {{
    background: var(--sw-white) !important; color: var(--sw-red) !important;
    border: 1.5px solid var(--sw-red) !important; border-radius: 2px !important;
    font-family: var(--font) !important; font-size: 0.88rem !important;
    font-weight: 600 !important; letter-spacing: 0.02em !important;
    padding: 0.55rem 1.8rem !important; transition: all 0.15s !important; }}
.stButton > button:hover {{ background: var(--sw-red) !important; color: white !important; }}
.stButton > button:disabled {{ opacity: 0.35 !important; }}

.stDownloadButton > button {{
    background: var(--sw-red) !important; color: white !important;
    border: none !important; border-radius: 2px !important;
    font-family: var(--font) !important; font-size: 0.88rem !important;
    font-weight: 600 !important; padding: 0.6rem 2rem !important; }}
.stDownloadButton > button:hover {{ opacity: 0.85 !important; }}

/* ── Stats strip ── */
.sw-stats {{ display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 1px; background: var(--sw-border);
    border: 1px solid var(--sw-border); border-radius: 6px;
    overflow: hidden; margin: 1.25rem 0; }}
.sw-stat {{ background: var(--sw-white); padding: 1.1rem 1.25rem; text-align: center; }}
.sw-stat-num {{ font-size: 2rem; font-weight: 700; color: var(--sw-navy);
    line-height: 1; margin-bottom: 0.25rem; }}
.sw-stat-num.green {{ color: #2d7d46; }}
.sw-stat-num.warn  {{ color: #c47f00; }}
.sw-stat-num.chf   {{ font-size: 1.2rem; color: var(--sw-navy); }}
.sw-stat-label {{ font-size: 0.67rem; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: var(--sw-muted); }}

/* ── Log ── */
.sw-log {{ background: #f9f9f9; border: 1px solid var(--sw-border);
    border-left: 3px solid var(--sw-navy); border-radius: 2px;
    font-family: var(--mono); font-size: 0.74rem; color: #444;
    padding: 1rem 1.2rem; line-height: 1.75; max-height: 280px;
    overflow-y: auto; white-space: pre-wrap; }}

/* ── File uploader ── */
[data-testid="stFileUploader"], [data-testid="stFileUploader"] section,
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploaderDropzone"] {{
    background: #fafafa !important; border: 1.5px dashed #c8d0da !important;
    border-radius: 3px !important; color: var(--sw-text) !important; }}
[data-testid="stFileUploaderDropzone"]:hover {{
    border-color: var(--sw-navy) !important; background: #f4f6f9 !important; }}
[data-testid="stFileUploaderDropzone"] svg {{
    color: var(--sw-navy) !important; fill: var(--sw-navy) !important; }}
[data-testid="stFileUploaderDropzoneInstructions"] > div > span,
[data-testid="stFileUploaderDropzoneInstructions"] span {{
    color: var(--sw-navy) !important; font-size: 0.88rem !important; font-weight: 600 !important; }}
[data-testid="stFileUploaderDropzoneInstructions"] small {{
    color: var(--sw-muted) !important; font-size: 0.75rem !important; }}
[data-testid="stFileUploaderDropzone"] button {{
    background: white !important; color: var(--sw-navy) !important;
    border: 1.5px solid var(--sw-navy) !important; border-radius: 2px !important;
    font-size: 0.8rem !important; font-weight: 600 !important; }}
[data-testid="stFileUploaderDropzone"] button:hover {{
    background: var(--sw-navy) !important; color: white !important; }}
[data-testid="stFileUploaderFile"], [data-testid="uploadedFileData"] {{
    background: white !important; border: 1px solid var(--sw-border) !important;
    border-radius: 2px !important; }}
[data-testid="stFileUploaderFileName"] {{ color: var(--sw-navy) !important; font-weight: 600 !important; }}

.stSpinner > div > div {{ border-top-color: var(--sw-red) !important; }}

.sw-success-bar {{ background: var(--sw-navy); color: white;
    padding: 0.65rem 1.25rem; font-size: 0.82rem; font-weight: 600;
    letter-spacing: 0.04em; border-radius: 4px; margin-bottom: 1.2rem;
    display: flex; align-items: center; gap: 8px; }}
.sw-success-bar::before {{ content: ''; width: 8px; height: 8px;
    border-radius: 50%; background: #5cb85c; flex-shrink: 0; }}

.sw-divider {{ border: none; border-top: 1px solid var(--sw-border); margin: 1rem 0; }}
.sw-chart-title {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: var(--sw-navy); margin-bottom: 0.75rem; }}
.sw-footer {{ text-align: center; font-size: 0.72rem; color: var(--sw-muted);
    padding: 2.5rem 0 1rem; letter-spacing: 0.04em; font-family: var(--mono); }}

.sw-section-bar {{ display: flex; align-items: center; gap: 12px; margin: 2.5rem 0 0.4rem; }}
.sw-section-bar-line {{ width: 3px; height: 20px; border-radius: 2px; flex-shrink: 0; }}
.sw-section-bar-title {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--sw-navy); }}
.sw-section-bar-sub {{ font-size: 0.82rem; color: #6b7280; margin: 0.1rem 0 1rem 15px; line-height: 1.5; }}
</style>
""", unsafe_allow_html=True)

# ─── TOPNAV (innerhalb block-container, bricht per CSS aus) ──────────────────
st.markdown("""
<nav class="sw-nav">
    <div class="sw-logo">
        <span class="sw-logo-text">swisspor</span>
        <div class="sw-logo-dots">
            <span style="background:#1a3a5c"></span>
            <span style="background:#f0a500"></span>
            <span style="background:#e63946"></span>
        </div>
    </div>
    <span class="sw-nav-title">swissporBIM - IFC-Attribuierung und Kostenermittlung</span>
    <span class="sw-nav-right">DC_BAT_FS26 - v5.0</span>
</nav>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs([
    "UC-01  -  IFC-Mapping",
    "UC-02  -  Mengen und Kosten"
])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 – MAPPING
# ════════════════════════════════════════════════════════════════════════════
with tab1:

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">1</div>
        <span class="sw-section-title">IFC-Modell hochladen</span>
    </div>
    <p class="sw-upload-hint">ArchiCAD / Revit Export mit <code>Pset_Swisspor_Identifikation</code>.</p>
    """, unsafe_allow_html=True)
    ifc_file = st.file_uploader("IFC", type=["ifc"], label_visibility="collapsed", key="ifc_map")
    if ifc_file:
        size = len(ifc_file.getvalue()) / 1024 / 1024
        st.markdown(f'<div class="sw-file-badge"> {ifc_file.name} - {size:.1f} MB</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">2</div>
        <span class="sw-section-title">System-Kataloge hochladen</span>
    </div>
    <p class="sw-upload-hint">System_01_GREENROOF.csv, System_02_KIES_SAFSYS.csv, …</p>
    """, unsafe_allow_html=True)
    csv_files = st.file_uploader("CSVs", type=["csv"], accept_multiple_files=True,
                                  label_visibility="collapsed", key="csv_map")
    if csv_files:
        st.markdown("".join(f'<div class="sw-file-badge"> {f.name}</div>' for f in csv_files),
                    unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">3</div>
        <span class="sw-section-title">Mapping ausführen</span>
    </div>
    """, unsafe_allow_html=True)
    if not (ifc_file and csv_files):
        st.markdown('<p class="sw-upload-hint">Bitte IFC-Modell und CSV-Kataloge hochladen.</p>', unsafe_allow_html=True)
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        run_map = st.button("Mapping starten", disabled=not (ifc_file and csv_files),
                             use_container_width=True, key="btn_map")
    st.markdown('</div>', unsafe_allow_html=True)

    if run_map and ifc_file and csv_files:
        with st.spinner("Mapping läuft …"):
            try:
                results = swisspor_mapping.process_mapping(ifc_file, csv_files)
            except Exception as e:
                import traceback
                st.error(f"Fehler: {e}")
                st.code(traceback.format_exc())
                st.stop()

        df = results["summary_table"]
        def _n(label):
            row = df[df["Kategorie"].str.contains(label, na=False)]
            return int(row["Anzahl"].values[0]) if len(row) else 0

        st.markdown(f'<div class="sw-success-bar">Mapping abgeschlossen — {len(csv_files)} Katalog(e) verarbeitet</div>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div class="sw-stats">
            <div class="sw-stat"><div class="sw-stat-num green">{_n("direkt")}</div><div class="sw-stat-label">Direkt gemappt</div></div>
            <div class="sw-stat"><div class="sw-stat-num">{_n("Gesamtaufbau")}</div><div class="sw-stat-label">Gesamtaufbau</div></div>
            <div class="sw-stat"><div class="sw-stat-num warn">{_n("Kein CSV")}</div><div class="sw-stat-label">Kein CSV</div></div>
            <div class="sw-stat"><div class="sw-stat-num" style="color:#999">{_n("bersprungen")}</div><div class="sw-stat-label">Übersprungen</div></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="sw-card">', unsafe_allow_html=True)
        st.markdown('<div class="sw-section-header"><div class="sw-section-num" style="background:#555">4</div><span class="sw-section-title">Protokoll und Download</span></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sw-log">{results["log"]}</div>', unsafe_allow_html=True)
        st.markdown('<hr class="sw-divider">', unsafe_allow_html=True)
        out_name = ifc_file.name.replace(".ifc", "_enriched.ifc")
        st.download_button(f"  {out_name} herunterladen", data=results["ifc_bytes"],
                           file_name=out_name, mime="application/octet-stream")
        st.markdown('</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# TAB 2 – MENGEN UND KOSTEN
# ════════════════════════════════════════════════════════════════════════════
with tab2:

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">1</div>
        <span class="sw-section-title">Enriched IFC hochladen</span>
    </div>
    <p class="sw-upload-hint">Das angereicherte IFC aus UC-01 – enthält alle Pset_Swisspor_* Schichten.</p>
    """, unsafe_allow_html=True)
    ifc_enriched = st.file_uploader("Enriched IFC", type=["ifc"],
                                     label_visibility="collapsed", key="ifc_kost")
    if ifc_enriched:
        size = len(ifc_enriched.getvalue()) / 1024 / 1024
        st.markdown(f'<div class="sw-file-badge"> {ifc_enriched.name} - {size:.1f} MB</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">2</div>
        <span class="sw-section-title">Preisdatenbank hochladen</span>
    </div>
    <p class="sw-upload-hint">
        Swisspor_Preisdatenbank_2026.xlsx – Sheets <code>Lookup_Kostenauswertung</code>
        (Art.-Nr. + EP_Einbau) und <code>Lookup_Drittprodukte</code> (CostLabel  CHF).
    </p>
    """, unsafe_allow_html=True)
    preisdb_file = st.file_uploader("Preisdatenbank", type=["xlsx"],
                                     label_visibility="collapsed", key="preisdb")
    if preisdb_file:
        st.markdown(f'<div class="sw-file-badge"> {preisdb_file.name}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sw-card">', unsafe_allow_html=True)
    st.markdown("""
    <div class="sw-section-header">
        <div class="sw-section-num">3</div>
        <span class="sw-section-title">Auswertung berechnen</span>
    </div>
    """, unsafe_allow_html=True)
    if not (ifc_enriched and preisdb_file):
        st.markdown('<p class="sw-upload-hint">Bitte enriched IFC und Preisdatenbank hochladen.</p>', unsafe_allow_html=True)
    col_btn2, _ = st.columns([1, 3])
    with col_btn2:
        run_kost = st.button("Kosten berechnen", disabled=not (ifc_enriched and preisdb_file),
                              use_container_width=True, key="btn_kost")
    st.markdown('</div>', unsafe_allow_html=True)

    if run_kost and ifc_enriched and preisdb_file:
        with st.spinner("Auswertung läuft …"):
            try:
                results_k = swisspor_kosten.process_kosten(ifc_enriched, preisdb_file)
            except Exception as e:
                import traceback
                st.error(f"Fehler: {e}")
                st.code(traceback.format_exc())
                st.stop()

        gt     = results_k["grand_total"]
        n_sw   = results_k.get("n_matched", 0)
        n_drit = results_k.get("n_drittprodukt", 0)
        n_sys  = results_k["n_systems"]

        st.markdown('<div class="sw-success-bar">Auswertung abgeschlossen</div>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div class="sw-stats">
            <div class="sw-stat"><div class="sw-stat-num green">{n_sw}</div><div class="sw-stat-label">swisspor Pos.</div></div>
            <div class="sw-stat"><div class="sw-stat-num warn">{n_drit}</div><div class="sw-stat-label">Drittprodukte</div></div>
            <div class="sw-stat"><div class="sw-stat-num">{n_sys}</div><div class="sw-stat-label">Systeme</div></div>
            <div class="sw-stat"><div class="sw-stat-num chf">CHF {gt:,.0f}</div><div class="sw-stat-label">Total exkl. MWST</div></div>
        </div>
        """, unsafe_allow_html=True)

        kosten_list = results_k.get("kosten_list", [])
        if kosten_list:
            import pandas as pd
            import plotly.graph_objects as go

            df_k = pd.DataFrame(kosten_list)
            df_k["func_label"] = (df_k["function"]
                                  .str.replace(r"^\d+_", "", regex=True)
                                  .str.strip())

            PALETTE = ["#1a3a5c", "#e63946", "#f0a500", "#2d7d46",
                       "#0077b6", "#8338ec", "#fb8500", "#6b7280"]
            systems = sorted(df_k["system_id"].unique())

            # ──────────────────────────────────────────────────────────────
            # A - MENGENAUSWERTUNG
            # ──────────────────────────────────────────────────────────────
            st.markdown("""
            <div class="sw-section-bar">
              <div class="sw-section-bar-line" style="background:#1a3a5c;"></div>
              <span class="sw-section-bar-title">A - Mengenauswertung nach Systemaufbau</span>
            </div>
            <p class="sw-section-bar-sub">Schichten je System mit Fläche, Bestellmengen und Einheitspreis.</p>
            """, unsafe_allow_html=True)

            for idx, sid in enumerate(systems):
                sys_df = df_k[df_k["system_id"] == sid].copy().sort_values("function")
                col_accent = PALETTE[idx % len(PALETTE)]
                n_schichten = len(sys_df)
                n_dritt_sys = int(sys_df["is_drittprodukt"].sum())

                base_rows = sys_df[sys_df["function"].str.startswith("01_")]
                sys_area = (base_rows["flaeche_m2"].sum()
                            if not base_rows.empty
                            else sys_df["flaeche_m2"].max())

                dritt_badge = (
                    f'<span style="font-size:0.68rem;background:rgba(255,255,255,0.18);'
                    f'color:white;padding:2px 8px;border-radius:3px;margin-left:10px;'
                    f'font-weight:600;">{n_dritt_sys} Drittprodukt(e)</span>'
                    if n_dritt_sys > 0 else ""
                )

                # Card-Kopf
                st.markdown(f"""
                <div style="background:white;border-radius:8px;margin-bottom:1.75rem;
                            box-shadow:0 1px 6px rgba(0,0,0,0.08);overflow:hidden;">
                  <div style="background:{col_accent};padding:0.9rem 1.5rem;
                               display:flex;align-items:center;justify-content:space-between;">
                    <div style="display:flex;align-items:center;gap:12px;">
                      <code style="background:rgba(255,255,255,0.18);border-radius:4px;
                                    padding:3px 10px;font-size:0.68rem;color:white;
                                    letter-spacing:0.1em;">SYSTEM {sid}</code>
                      <span style="font-size:0.9rem;font-weight:600;color:white;">
                        {n_schichten} Schichten{dritt_badge}
                      </span>
                    </div>
                    <div style="text-align:right;">
                      <div style="font-size:1.5rem;font-weight:700;color:white;line-height:1;">
                        {sys_area:,.2f} m²
                      </div>
                      <div style="font-size:0.62rem;color:rgba(255,255,255,0.65);
                                   text-transform:uppercase;letter-spacing:0.08em;">Systemfläche</div>
                    </div>
                  </div>
                  <div style="display:grid;grid-template-columns:8px 170px 1fr 100px 210px 120px;
                               gap:0 1rem;padding:0.45rem 1.5rem;
                               background:#f8f9fb;border-bottom:1.5px solid #eaecf0;">
                    <div></div>
                    <div style="font-size:0.61rem;font-weight:700;letter-spacing:0.09em;
                                 text-transform:uppercase;color:#9ca3af;">Schichtfunktion</div>
                    <div style="font-size:0.61rem;font-weight:700;letter-spacing:0.09em;
                                 text-transform:uppercase;color:#9ca3af;">Produkt / Art.-Nr.</div>
                    <div style="font-size:0.61rem;font-weight:700;letter-spacing:0.09em;
                                 text-transform:uppercase;color:#9ca3af;text-align:right;">Fläche</div>
                    <div style="font-size:0.61rem;font-weight:700;letter-spacing:0.09em;
                                 text-transform:uppercase;color:#9ca3af;">Bestellmenge</div>
                    <div style="font-size:0.61rem;font-weight:700;letter-spacing:0.09em;
                                 text-transform:uppercase;color:#9ca3af;text-align:right;">EP CHF</div>
                  </div>
                """, unsafe_allow_html=True)

                for i, (_, row) in enumerate(sys_df.iterrows()):
                    bg     = "#fafbfc" if i % 2 == 0 else "white"
                    is_d   = row["is_drittprodukt"]
                    p_clr  = "#b45309" if is_d else "#1a3a5c"
                    art    = row["article_nr"] if row["article_nr"] else "—"
                    d_tag  = ('<span style="font-size:0.58rem;background:#fef3c7;color:#92400e;'
                              'padding:1px 5px;border-radius:2px;margin-left:5px;font-weight:600;">Dritt</span>'
                              if is_d else "")

                    st.markdown(f"""
                    <div style="display:grid;grid-template-columns:8px 170px 1fr 100px 210px 120px;
                                 gap:0 1rem;padding:0.65rem 1.5rem;
                                 background:{bg};border-bottom:1px solid #f0f2f5;align-items:center;">
                      <div style="display:flex;justify-content:center;">
                        <div style="width:7px;height:7px;border-radius:50%;background:{col_accent};"></div>
                      </div>
                      <div style="font-size:0.77rem;font-weight:600;color:#374151;">{row['func_label']}</div>
                      <div>
                        <div style="font-size:0.87rem;font-weight:600;color:{p_clr};">
                          {row['product_name'] or '—'}{d_tag}
                        </div>
                        <div style="font-size:0.7rem;color:#9ca3af;font-family:'DM Mono',monospace;margin-top:1px;">
                          {art}
                        </div>
                      </div>
                      <div style="text-align:right;">
                        <div style="font-size:0.9rem;font-weight:700;color:#1a3a5c;">{"—" if (row['flaeche_m2'] or 0) == 0.0 else f"{row['flaeche_m2']:,.2f}"}</div>
                        <div style="font-size:0.6rem;color:#9ca3af;text-transform:uppercase;">{"lfm" if (row['flaeche_m2'] or 0) == 0.0 else "m²"}</div>
                      </div>
                      <div style="font-size:0.76rem;color:#4b5563;line-height:1.4;">{row['bestell_detail'] or '—'}</div>
                      <div style="text-align:right;">
                        <div style="font-size:0.8rem;font-weight:600;color:#374151;
                                     font-family:'DM Mono',monospace;">{row['ep_material']:,.2f}</div>
                        <div style="font-size:0.6rem;color:#9ca3af;">CHF/Einheit</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("</div>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────
            # B - KOSTENAUSWERTUNG
            # ──────────────────────────────────────────────────────────────
            st.markdown("""
            <div class="sw-section-bar" style="margin-top:3rem;">
              <div class="sw-section-bar-line" style="background:#e63946;"></div>
              <span class="sw-section-bar-title">B - Kostenauswertung nach System</span>
            </div>
            <p class="sw-section-bar-sub">Material- und Einbaukosten je System sowie CHF/m² Vergleich.</p>
            """, unsafe_allow_html=True)

            sys_records = []
            for idx, sid in enumerate(systems):
                sys_df = df_k[df_k["system_id"] == sid]
                base_rows = sys_df[sys_df["function"].str.startswith("01_")]
                sys_area = (base_rows["flaeche_m2"].sum()
                            if not base_rows.empty
                            else sys_df["flaeche_m2"].max())
                total  = sys_df["poskosten_total"].sum()
                mat    = sys_df["poskosten_mat"].sum()
                ein    = sys_df["poskosten_ein"].sum()
                chf_m2 = total / sys_area if sys_area > 0 else 0
                sys_records.append({
                    "sid": sid, "area": sys_area, "mat": mat, "ein": ein,
                    "total": total, "chf_m2": chf_m2, "n_pos": len(sys_df),
                    "color": PALETTE[idx % len(PALETTE)]
                })

            grand_total = sum(r["total"] for r in sys_records)

            # System-Cards (2 Spalten)
            n_cols = min(len(sys_records), 2)
            cols = st.columns(n_cols)
            for idx, rec in enumerate(sys_records):
                pct = (rec["total"] / grand_total * 100) if grand_total > 0 else 0
                with cols[idx % n_cols]:
                    st.markdown(f"""
                    <div style="background:white;border-radius:8px;margin-bottom:1rem;
                                box-shadow:0 1px 6px rgba(0,0,0,0.07);overflow:hidden;">
                      <div style="height:4px;background:{rec['color']};"></div>
                      <div style="padding:1.1rem 1.3rem;">
                        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0.65rem;">
                          <div>
                            <div style="font-size:0.63rem;font-weight:700;letter-spacing:0.12em;
                                         text-transform:uppercase;color:#9ca3af;margin-bottom:2px;">System {rec['sid']}</div>
                            <div style="font-size:1.65rem;font-weight:700;color:#1a3a5c;line-height:1.1;">
                              CHF {rec['total']:,.0f}
                            </div>
                          </div>
                          <div style="text-align:right;">
                            <div style="font-size:0.63rem;letter-spacing:0.08em;text-transform:uppercase;color:#9ca3af;">Anteil</div>
                            <div style="font-size:1.4rem;font-weight:700;color:{rec['color']};">{pct:.1f}%</div>
                          </div>
                        </div>
                        <div style="background:#f0f2f5;border-radius:4px;height:5px;margin-bottom:0.9rem;">
                          <div style="background:{rec['color']};height:5px;border-radius:4px;width:{min(pct,100):.1f}%;"></div>
                        </div>
                        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:0.4rem;margin-bottom:0.6rem;">
                          <div style="background:#f8f9fb;border-radius:5px;padding:0.55rem 0.7rem;text-align:center;">
                            <div style="font-size:0.88rem;font-weight:700;color:#1a3a5c;">{rec['area']:,.1f}</div>
                            <div style="font-size:0.58rem;color:#9ca3af;text-transform:uppercase;letter-spacing:0.06em;">m²</div>
                          </div>
                          <div style="background:#f8f9fb;border-radius:5px;padding:0.55rem 0.7rem;text-align:center;">
                            <div style="font-size:0.88rem;font-weight:700;color:{rec['color']};">{rec['chf_m2']:,.1f}</div>
                            <div style="font-size:0.58rem;color:#9ca3af;text-transform:uppercase;letter-spacing:0.06em;">CHF/m²</div>
                          </div>
                          <div style="background:#f8f9fb;border-radius:5px;padding:0.55rem 0.7rem;text-align:center;">
                            <div style="font-size:0.88rem;font-weight:700;color:#374151;">{rec['n_pos']}</div>
                            <div style="font-size:0.58rem;color:#9ca3af;text-transform:uppercase;letter-spacing:0.06em;">Pos.</div>
                          </div>
                        </div>
                        <div style="display:grid;grid-template-columns:1fr 1fr;gap:0.4rem;">
                          <div style="padding:0.4rem 0.7rem;border-left:3px solid #e5e7eb;">
                            <div style="font-size:0.65rem;color:#9ca3af;">Material</div>
                            <div style="font-size:0.82rem;font-weight:600;color:#374151;font-family:'DM Mono',monospace;">
                              CHF {rec['mat']:,.0f}
                            </div>
                          </div>
                          <div style="padding:0.4rem 0.7rem;border-left:3px solid #e5e7eb;">
                            <div style="font-size:0.65rem;color:#9ca3af;">Einbau</div>
                            <div style="font-size:0.82rem;font-weight:600;color:#374151;font-family:'DM Mono',monospace;">
                              CHF {rec['ein']:,.0f}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

            # Charts
            col_l, col_r = st.columns([3, 2])
            with col_l:
                st.markdown('<div class="sw-card">', unsafe_allow_html=True)
                st.markdown('<p class="sw-chart-title">CHF/m² nach System</p>', unsafe_allow_html=True)
                fig_m2 = go.Figure(go.Bar(
                    x=[f"System {r['sid']}" for r in sys_records],
                    y=[r["chf_m2"] for r in sys_records],
                    marker_color=[r["color"] for r in sys_records],
                    text=[f"CHF {r['chf_m2']:,.1f}" for r in sys_records],
                    textposition="outside",
                    textfont=dict(size=11, family="Arial"),
                    hovertemplate="<b>%{x}</b><br>CHF/m²: %{y:,.2f}<extra></extra>",
                ))
                fig_m2.update_layout(
                    margin=dict(l=10, r=10, t=20, b=10),
                    height=280,
                    xaxis=dict(showgrid=False, title=""),
                    yaxis=dict(showgrid=True, gridcolor="#f0f2f5", title="CHF/m²", tickformat=",.0f"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    font=dict(family="Arial", size=11, color="#1a1a1a"),
                    showlegend=False,
                )
                st.plotly_chart(fig_m2, use_container_width=True, config={"displayModeBar": False})
                st.markdown('</div>', unsafe_allow_html=True)

            with col_r:
                st.markdown('<div class="sw-card">', unsafe_allow_html=True)
                st.markdown('<p class="sw-chart-title">Kostenverteilung</p>', unsafe_allow_html=True)
                fig_pie = go.Figure(go.Pie(
                    labels=[f"Sys {r['sid']}" for r in sys_records],
                    values=[r["total"] for r in sys_records],
                    hole=0.52,
                    marker_colors=[r["color"] for r in sys_records],
                    textinfo="label+percent",
                    textfont=dict(size=11, family="Arial"),
                    hovertemplate="<b>%{label}</b><br>CHF %{value:,.0f} (%{percent})<extra></extra>",
                ))
                fig_pie.update_layout(
                    margin=dict(l=10, r=10, t=10, b=10), height=280,
                    showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                    font=dict(family="Arial", size=11, color="#1a1a1a"),
                    annotations=[dict(
                        text=f"<b>CHF {grand_total:,.0f}</b>",
                        x=0.5, y=0.5, font=dict(size=12, family="Arial", color="#1a3a5c"),
                        showarrow=False
                    )],
                )
                st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})
                st.markdown('</div>', unsafe_allow_html=True)

            # MWST Totalbox
            mwst_rate   = 0.081
            mwst_betrag = grand_total * mwst_rate
            total_inkl  = grand_total * (1 + mwst_rate)
            st.markdown(f"""
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:1px;
                         background:#e0e0e0;border-radius:6px;overflow:hidden;margin:0.5rem 0 1.5rem;">
              <div style="background:white;padding:0.9rem 1.25rem;">
                <div style="font-size:0.62rem;font-weight:700;letter-spacing:0.1em;
                             text-transform:uppercase;color:#9ca3af;margin-bottom:3px;">Total exkl. MWST</div>
                <div style="font-size:1.2rem;font-weight:700;color:#1a3a5c;font-family:'DM Mono',monospace;">
                  CHF {grand_total:,.2f}
                </div>
              </div>
              <div style="background:white;padding:0.9rem 1.25rem;">
                <div style="font-size:0.62rem;font-weight:700;letter-spacing:0.1em;
                             text-transform:uppercase;color:#9ca3af;margin-bottom:3px;">MWST 8.1%</div>
                <div style="font-size:1.2rem;font-weight:700;color:#6b7280;font-family:'DM Mono',monospace;">
                  CHF {mwst_betrag:,.2f}
                </div>
              </div>
              <div style="background:#1a3a5c;padding:0.9rem 1.25rem;">
                <div style="font-size:0.62rem;font-weight:700;letter-spacing:0.1em;
                             text-transform:uppercase;color:rgba(255,255,255,0.55);margin-bottom:3px;">Total inkl. MWST</div>
                <div style="font-size:1.2rem;font-weight:700;color:white;font-family:'DM Mono',monospace;">
                  CHF {total_inkl:,.2f}
                </div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            # Detailtabelle
            st.markdown('<div class="sw-card">', unsafe_allow_html=True)
            st.markdown('<p class="sw-chart-title">Alle Positionen – Detailansicht</p>', unsafe_allow_html=True)
            table_cols = [
                "system_id", "func_label", "product_name", "article_nr",
                "flaeche_m2", "laenge_m", "mengenbasis", "bestell_detail",
                "ep_material", "ep_einbau",
                "poskosten_mat", "poskosten_ein", "poskosten_total",
                "is_drittprodukt"
            ]
            for c in ["laenge_m", "mengenbasis"]:
                if c not in df_k.columns:
                    df_k[c] = None
            df_table = df_k[table_cols].copy()
            df_table.columns = [
                "System", "Funktion", "Produkt", "Art.-Nr.",
                "Fläche m²", "Länge m", "Mengenbasis", "Bestellmenge",
                "EP Mat.", "EP Ein.",
                "Material CHF", "Einbau CHF", "Total CHF", "Drittprodukt"
            ]
            st.dataframe(
                df_table, hide_index=True, use_container_width=True, height=360,
                column_config={
                    "Total CHF":    st.column_config.NumberColumn(format="CHF %.2f"),
                    "Material CHF": st.column_config.NumberColumn(format="CHF %.2f"),
                    "Einbau CHF":   st.column_config.NumberColumn(format="CHF %.2f"),
                    "EP Mat.":      st.column_config.NumberColumn(format="CHF %.2f"),
                    "EP Ein.":      st.column_config.NumberColumn(format="CHF %.2f"),
                    "Fläche m²":    st.column_config.NumberColumn(format="%.2f m²"),
                    "Länge m":      st.column_config.NumberColumn(format="%.2f m"),
                }
            )
            st.markdown("""
            <div style="font-size:0.71rem;color:#9ca3af;margin-top:5px;">
              Drittprodukt = True: Platzhalter CHF 0.00 – Preis manuell erfassen.
            </div>""", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Protokoll und Download
        st.markdown('<div class="sw-card">', unsafe_allow_html=True)
        st.markdown('<div class="sw-section-header"><div class="sw-section-num" style="background:#555">4</div><span class="sw-section-title">Protokoll und Download</span></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sw-log">{results_k["log"]}</div>', unsafe_allow_html=True)
        st.markdown('<hr class="sw-divider">', unsafe_allow_html=True)
        if results_k["excel_bytes"]:
            st.download_button(
                "  Kostenauswertung_Swisspor.xlsx herunterladen",
                data=results_k["excel_bytes"],
                file_name="Kostenauswertung_Swisspor.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Keine Kostenpositionen — enriched IFC und Preisdatenbank prüfen.")
        st.markdown('</div>', unsafe_allow_html=True)

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="sw-footer">
    swisspor AG - Bahnhofstrasse 50 - CH-6312 Steinhausen
     -  swissporBIM v5.0 - DC_BAT_FS26 - HSLU
</div>
""", unsafe_allow_html=True)
