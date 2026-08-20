# -*- coding: utf-8 -*-
"""
Datenbasis für 02_Massnahmenkatalog.
Jede Zeile: (id, kategorie, name, einheit, formula_kind, einheitspreis, nutzungsdauer,
             foerderkategorie, foerdersatz_text, pflicht_eh55, pflicht_eh40, klimaneutral,
             eigenleistung_moeglich, trigger, kommentar)

formula_kind steuert, welche Mengenformel auf dem Objektblatt erzeugt wird (siehe build_workbook.py):
  'dachflaeche'      -> Obj_Dachflaeche
  'fassadenflaeche'  -> Obj_Fassadenflaeche
  'fensterflaeche'   -> Obj_Fensterflaeche
  'anz_fenster'      -> Obj_AnzFenster
  'wohnflaeche'      -> Obj_Wohnflaeche
  'grundflaeche'     -> Obj_Grundflaeche
  'umfang'           -> Obj_Umfang
  'anz_heizkoerper'  -> Obj_AnzHeizkoerper
  'anz_tueren'       -> Obj_AnzTueren
  'anz_raeume'       -> Obj_AnzRaeume
  'pauschal1'        -> 1
  'pauschal0'        -> 0 (nur bei Bedarf manuell auf 1 setzen)
  'anz_baeder'       -> Obj_AnzBaeder
  'geschosse_minus1' -> MAX(0, Obj_Geschosse-1)
  'kwp'              -> PV_kWp (aus Klimaneutralität-Block)
  'speicher_kwh'     -> PV_Speicher_kWh

trigger: Bedingung (True/False als Python) ODER Excel-Formel-Text, die bestimmt ob Default-Override
         "automatik" auf "entfällt" stehen sollte (nur informativ, wird als Kommentar genutzt).
"""

CATALOG = [
    # --- HUELLE ---
    dict(id="H01", kat="Hülle", name="Steildach komplett neu (Abriss+Aufsparrendämmung U≤0,14+Eindeckung+Klempner)",
         einheit="€/m² Dachfläche", fk="dachflaeche", preis=520, nutzung=55, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False,
         kommentar="Nur wenn Dachform=Satteldach; Schwäbisch Hall 2026 nennt 400-600€/m² für eine "
                    "umfangreiche Dachsanierung, hier oberer Mittelwert wegen hoher Dämmanforderung "
                    "U≤0,14 angesetzt (Recherche-Update 20.08.2026, vorher 420€/m² zu niedrig laut "
                    "Nutzer-Praxistest; siehe auch reduco.ai/profirechner.de/klotz-bedachungen.de 2026: "
                    "200-350€/m² für Standard-Ausführung, Schwäbisch Hall als konservativere Bank-Quelle "
                    "für 'umfangreich' bevorzugt)"),
    dict(id="H02", kat="Hülle", name="Flachdachsanierung (PVC/FPO inkl. Gefälledämmung, Abriss)",
         einheit="€/m² Dachfläche", fk="dachflaeche", preis=250, nutzung=20, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False, kommentar="Nur wenn Dachform=Flachdach; dachdeckernrw.org 2026"),
    dict(id="H03", kat="Hülle", name="Dachbodendämmung oberste Geschossdecke (falls kein Dachausbau)",
         einheit="€/m²", fk="grundflaeche", preis=27, nutzung=40, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=True, kommentar="co2online 2026"),
    dict(id="H04", kat="Hülle", name="WDVS Außenwand (U≤0,20, 16cm) inkl. Gerüst+Putz",
         einheit="€/m² Fassade", fk="fassadenflaeche", preis=180, nutzung=40, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False, kommentar="co2online/Verbraucherzentrale 2026; BBSR-Neufassung 2025 setzt Nutzungsdauer >50J an, konservativ 40J angesetzt"),
    dict(id="H05", kat="Hülle", name="Fenster 3-fach-Verglasung inkl. Einbau",
         einheit="€/m² Fensterfläche", fk="fensterflaeche", preis=285, nutzung=37, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False, kommentar="my-hammer.de Preisradar 2026"),
    dict(id="H06", kat="Hülle", name="Haustür RC2/RC3 inkl. Einbau",
         einheit="€/Stück", fk="pauschal1", preis=2800, nutzung=45, foerder=True, satz="15%+5% iSFP",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="H07", kat="Hülle", name="Nebeneingangstür Stahl RC2",
         einheit="€/Stück", fk="geschosse_minus1", preis=1000, nutzung=45, foerder=True, satz="15%+5% iSFP",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="H08", kat="Hülle", name="Rollläden elektrisch nachrüsten",
         einheit="€/Fenster", fk="anz_fenster", preis=850, nutzung=30, foerder=True, satz="15%+5% iSFP",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="H09", kat="Hülle", name="Kellerdeckendämmung von unten",
         einheit="€/m²", fk="grundflaeche", preis=32, nutzung=40, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=True, kommentar="Nur wenn unbeheizter Keller; co2online 2026"),
    dict(id="H10", kat="Hülle", name="Perimeterdämmung + Kellerabdichtung außen",
         einheit="€/lfm Umfang", fk="umfang", preis=300, nutzung=40, foerder=True, satz="15%+5% iSFP",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur bei Vollkeller + Abdichtungsbedarf; daemmen-und-sanieren.de 2026"),
    dict(id="H11", kat="Hülle", name="Gerüst (Fassade+Dach, pauschal)",
         einheit="€ pauschal", fk="pauschal1", preis=2700, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="geruest-preis.de 2026"),
    dict(id="H12", kat="Hülle", name="Blower-Door-Test + Wärmebrückennachweis",
         einheit="€ pauschal", fk="pauschal1", preis=1140, nutzung=0, foerder=True, satz="Teil Fachplanung 50%",
         eh55=True, eh40=True, klima=False, el=False, kommentar="blowerdoormr.de/renewa.de 2026"),

    # --- ANLAGENTECHNIK ---
    dict(id="A01", kat="Anlagentechnik", name="Luft-Wasser-Wärmepumpe Gesamtsystem inkl. Demontage Altanlage",
         einheit="€ pauschal (skaliert)", fk="wohnflaeche_skaliert_wp", preis=22500, nutzung=15, foerder=True,
         satz="KfW458: 30%+16%+Einkommen, Deckel 28.000€", eh55=True, eh40=True, klima=False, el=False,
         kommentar="Referenz 22.500€ bei 130m² (co2online), linear mit Wohnfläche skaliert"),
    dict(id="A02", kat="Anlagentechnik", name="Alternativ: Sole-Wasser-WP mit Erdsonde inkl. Bohrung",
         einheit="€ pauschal (skaliert)", fk="wohnflaeche_skaliert_wp_erd", preis=28500, nutzung=27, foerder=True,
         satz="KfW458: 30%+16%+Einkommen, Deckel 28.000€", eh55=False, eh40=False, klima=False, el=False,
         kommentar="Alternative zu A01, nicht gleichzeitig ansetzen; my-hammer.de 2026"),
    dict(id="A03", kat="Anlagentechnik", name="Fußbodenheizung Nassestrich (Vollsanierung Bodenaufbau)",
         einheit="€/m²", fk="wohnflaeche_fbh", preis=65, nutzung=45, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False, kommentar="co2online 2026"),
    dict(id="A04", kat="Anlagentechnik", name="Austausch Heizkörper gegen NT-Heizkörper (wenn keine FBH)",
         einheit="€/Stück", fk="anz_heizkoerper", preis=830, nutzung=35, foerder=True, satz="15%+5% iSFP",
         eh55=False, eh40=False, klima=False, el=True, kommentar="Schwäbisch Hall 2026"),
    dict(id="A05", kat="Anlagentechnik", name="Hydraulischer Abgleich + Heizlastberechnung",
         einheit="€ pauschal", fk="pauschal1", preis=1125, nutzung=0, foerder=True, satz="15% Heizungsoptimierung",
         eh55=True, eh40=True, klima=False, el=False, kommentar="energie-experten.org 2026"),
    dict(id="A06", kat="Anlagentechnik", name="Zentrale Lüftungsanlage mit WRG inkl. Kanalnetz",
         einheit="€/m² WF", fk="wohnflaeche", preis=111, nutzung=25, foerder=True, satz="15%+5% iSFP",
         eh55=True, eh40=True, klima=False, el=False, kommentar="Schwäbisch Hall 2026, berechnet aus 16.660€/150m²"),
    dict(id="A07", kat="Anlagentechnik", name="Trinkwasserinstallation komplett neu",
         einheit="€ pauschal", fk="pauschal1", preis=3500, nutzung=39, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur wenn Leitungen >40 Jahre; byndl.de 2026"),
    dict(id="A08", kat="Anlagentechnik", name="Abwasserleitungen im Haus erneuern",
         einheit="€ pauschal", fk="pauschal1", preis=2700, nutzung=40, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur wenn >45 Jahre alt; byndl.de 2026"),

    # --- ELEKTRO / INNENAUSBAU ---
    dict(id="E01", kat="Elektro/Innenausbau", name="Elektroinstallation komplett neu inkl. Zählerschrank",
         einheit="€/m² WF", fk="wohnflaeche", preis=105, nutzung=40, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="E02", kat="Elektro/Innenausbau", name="Nur Zählerschrank/Unterverteilung erneuern",
         einheit="€ pauschal", fk="pauschal1", preis=900, nutzung=40, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Alternative zu E01; my-hammer.de 2026"),
    dict(id="E03", kat="Elektro/Innenausbau", name="Badsanierung komplett, Standard",
         einheit="€/Bad", fk="anz_baeder", preis=15000, nutzung=32, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de/Schwäbisch Hall 2026"),
    dict(id="E04", kat="Elektro/Innenausbau", name="Gäste-WC",
         einheit="€ pauschal", fk="pauschal1", preis=6000, nutzung=32, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Abgeleitet aus Bad-Kennwerten, unverifiziert"),
    dict(id="E05", kat="Elektro/Innenausbau", name="Innenputz/Malerarbeiten",
         einheit="€/m² Wandfläche", fk="fassadenflaeche_x2", preis=45, nutzung=20, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=True, kommentar="Schwäbisch Hall 2026"),
    dict(id="E06", kat="Elektro/Innenausbau", name="Innentüren ersetzen",
         einheit="€/Stück", fk="anz_tueren", preis=600, nutzung=55, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=True, kommentar="Schwäbisch Hall 2026"),
    dict(id="E07", kat="Elektro/Innenausbau", name="Bodenbeläge (Mix Parkett/Fliese)",
         einheit="€/m² WF", fk="wohnflaeche", preis=120, nutzung=35, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=True, kommentar="Schwäbisch Hall 2026, gemischte Nutzungsdauer"),
    dict(id="E08", kat="Elektro/Innenausbau", name="Treppe erneuern",
         einheit="€ pauschal", fk="pauschal1", preis=3200, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026, nur wenn mehrgeschossig+Bedarf"),
    dict(id="E09", kat="Elektro/Innenausbau", name="Wanddurchbruch tragend inkl. Statiker",
         einheit="€/Stück", fk="pauschal0", preis=2900, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026, nur bei Grundrissänderung"),
    dict(id="E10", kat="Elektro/Innenausbau", name="Entkernung komplett (nur Vollsanierung Baujahr <1960)",
         einheit="€/m² WF", fk="wohnflaeche", preis=95, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="dus2b.de 2026"),

    # --- RISIKO / SCHADSTOFFE ---
    dict(id="R01", kat="Risiko/Schadstoff", name="Asbest-Eternitplatten Dach: Demontage+Entsorgung",
         einheit="€/m² Dachfläche", fk="dachflaeche", preis=40, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur wenn Baujahr<1994 UND Dach=Eternit; my-hammer.de 2026"),
    dict(id="R02", kat="Risiko/Schadstoff", name="Vinyl-Asbest-Fliesen entfernen",
         einheit="€/m² WF", fk="wohnflaeche", preis=40, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur wenn Baujahr<1994; kostencheck.de 2026"),
    dict(id="R03", kat="Risiko/Schadstoff", name="Nachtspeicherofen entsorgen (asbestverdächtig)",
         einheit="€/Gerät", fk="pauschal0", preis=220, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="R04", kat="Risiko/Schadstoff", name="Öltank-Stilllegung + Entsorgung",
         einheit="€ pauschal", fk="pauschal0", preis=1500, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur wenn Ölheizung vorhanden; my-hammer.de 2026"),
    dict(id="R05", kat="Risiko/Schadstoff", name="Horizontalsperre gegen aufsteigende Feuchte",
         einheit="€/lfm Außenwand", fk="umfang", preis=300, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Nur bei Feuchteschaden Keller; vallovapor.de 2026"),
    dict(id="R06", kat="Risiko/Schadstoff", name="Laboranalyse Asbestprobe + Dokumentation (§5a GefStoffV)",
         einheit="€ pauschal", fk="pauschal1", preis=300, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Pflicht bei Baujahr<1994 vor Bauarbeiten"),

    # --- KLIMANEUTRALITAET ---
    dict(id="K01", kat="Klimaneutralität", name="PV-Anlage",
         einheit="€/kWp (gestaffelt)", fk="pv_kwp", preis=1200, nutzung=25, foerder=False, satz="0% MwSt.",
         eh55=False, eh40=False, klima=True, el=False, kommentar="42watt.de 2026, Staffel 1.500->950€/kWp, hier Mittelwert 10kWp-Klasse"),
    dict(id="K02", kat="Klimaneutralität", name="Batteriespeicher",
         einheit="€/kWh", fk="pv_speicher_kwh", preis=350, nutzung=18, foerder=False, satz="0% MwSt.",
         eh55=False, eh40=False, klima=True, el=False, kommentar="reduco.ai 2026"),
    dict(id="K03", kat="Klimaneutralität", name="Energiemanagementsystem/Wechselrichter-Aufpreis",
         einheit="€ pauschal", fk="pauschal1", preis=1500, nutzung=15, foerder=True, satz="15%+5% (digitale Systeme)",
         eh55=False, eh40=False, klima=True, el=False, kommentar="Schätzung, EMS-Einzelpreise nicht verifizierbar"),
    dict(id="K04", kat="Klimaneutralität", name="Notstrom-/Ersatzstromfähigkeit",
         einheit="€ pauschal", fk="pauschal0", preis=2500, nutzung=15, foerder=False, satz="-",
         eh55=False, eh40=False, klima=True, el=False, kommentar="reduco.ai 2026, optional"),

    # --- SMART HOME ---
    dict(id="S01", kat="Smart Home", name="Stufe 1 – Netzwerk-Basis (CAT7, Patchfeld, PoE-Switch, APs)",
         einheit="€/Dose", fk="anz_raeume_x1_5", preis=250, nutzung=25, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Stahlbergen.de 2026, Nachrüstung bewohnt"),
    dict(id="S02", kat="Smart Home", name="Stufe 2 – KNX Basis (Licht, Rollladen, Heizungsaktorik, Server)",
         einheit="€/m² WF", fk="pauschal0", preis=130, nutzung=25, foerder=False, satz="EMS-Anteil ggf. 15%+5%",
         eh55=False, eh40=False, klima=False, el=False, kommentar="smarthome-exklusiv.de 2026, untere Bandbreite"),
    dict(id="S03", kat="Smart Home", name="Stufe 3 – KNX Vollausbau (Aufschlag auf Stufe 2)",
         einheit="€/m² WF Aufschlag", fk="pauschal0", preis=100, nutzung=25, foerder=False, satz="EMS-Anteil ggf. förderfähig",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Schätzung, nur mit S02 zusammen sinnvoll"),

    # --- AUSSENANLAGEN ---
    dict(id="G01", kat="Außenanlagen", name="Zufahrt pflastern",
         einheit="€/m²", fk="pauschal0", preis=120, nutzung=36, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="G02", kat="Außenanlagen", name="Terrasse",
         einheit="€/m²", fk="pauschal0", preis=110, nutzung=30, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="G03", kat="Außenanlagen", name="Zaun",
         einheit="€/lfm", fk="pauschal0", preis=60, nutzung=25, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=True, kommentar="my-hammer.de 2026"),
    dict(id="G04", kat="Außenanlagen", name="Garage/Carport",
         einheit="€ pauschal", fk="pauschal0", preis=11000, nutzung=40, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),

    # --- WEICHE KOSTEN ---
    dict(id="W01", kat="Weiche Kosten", name="Energieberatung + iSFP",
         einheit="€ pauschal", fk="pauschal1", preis=1650, nutzung=0, foerder=True, satz="50%, Deckel 650€",
         eh55=True, eh40=True, klima=False, el=False, kommentar="BAFA/my-hammer.de 2026"),
    dict(id="W02", kat="Weiche Kosten", name="Fachplanung/Baubegleitung (EEE, Pflicht bei BEG-Antrag)",
         einheit="% der Sanierungssumme", fk="pauschal1_prozent3", preis=0, nutzung=0, foerder=True,
         satz="50%, Deckel 5.000€", eh55=True, eh40=True, klima=False, el=False,
         kommentar="3% der Hülle+Anlagentechnik-Summe, siehe Formel"),
    dict(id="W03", kat="Weiche Kosten", name="Statiker (nur bei Grundrissänderung/Dachausbau)",
         einheit="€ pauschal", fk="pauschal0", preis=1900, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="my-hammer.de 2026"),
    dict(id="W04", kat="Weiche Kosten", name="Bauantrag/Genehmigung",
         einheit="€ pauschal", fk="pauschal0", preis=1700, nutzung=0, foerder=False, satz="-",
         eh55=False, eh40=False, klima=False, el=False, kommentar="Schwäbisch Hall 2026, nur bei genehmigungspflichtigen Maßnahmen"),
]
