Swisspor Gruppierungs-Fix

Geändert in swisspor_kosten.py:
- Positionen werden nicht mehr nur nach (SystemID, Function) gruppiert.
- Neuer Schlüssel: SystemID, Function, Artikelnummer, Produktname, Mindestdicke, Format.
- Dadurch werden mehrere Wärmedämmungen im gleichen System sauber getrennt bzw. bei identischem Artikel/Dicke/Format korrekt addiert.
- Mengenbasis wird als Einzelschicht oder Gesamtaufbau-Fallback weitergeführt.
- Absturzsicherung bleibt als Laufmeter-Logik erhalten.

Start:
streamlit run app.py
