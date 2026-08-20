# AGENTS.md — Leitfaden für die Weiterentwicklung

Dieses Repo generiert ein Excel-Finanzmodell (`Hausbewertung_Altkreis_Luebbecke.xlsx`)
per Python/`openpyxl` aus `build/build_workbook.py`. Diese Datei ist der Leitfaden für
jede KI-Agentin/jeden Agenten (oder Mensch), der/die daran weiterarbeitet.

## Grundprinzip: die `.xlsx` wird nie von Hand editiert

**Ändere niemals die `.xlsx`-Datei direkt in Excel/LibreOffice und committe das.**
Jede Änderung passiert im Python-Skript (`build/build_workbook.py` bzw.
`build/catalog_data.py`), dann wird die Datei neu gebaut und neu berechnet. Grund:
Nachvollziehbarkeit im Git-Diff, und die Datei enthält ca. 2000 Formeln mit
durchgängig berechneten Zeilennummern (siehe unten) – manuelle Edits brechen das
Muster fast garantiert.

## Workflow für jede Änderung

```bash
cd build
python3 build_workbook.py                                    # 1. neu bauen
python3 recalc.py ../Hausbewertung_Altkreis_Luebbecke.xlsx 150   # 2. neu berechnen
```

`build/recalc.py` ist ein eigenständiges, im Repo mitgeliefertes Skript (kein
Anthropic-internes Sandbox-Werkzeug, funktioniert also auch außerhalb von Claude
Code) – es ruft `soffice --headless --convert-to xlsx` auf einer temporären Kopie
auf und meldet danach `#REF!`/`#NAME?`/… als JSON. Voraussetzung ist eine lokale
LibreOffice-Installation (`soffice` bzw. `libreoffice` auf dem PATH); ist keine
vorhanden, bricht das Skript mit einer klaren Fehlermeldung ab, statt sich in
irgendeiner Form zu verschlucken.

Danach **zwingend** prüfen:

1. Das JSON von `recalc.py` muss `"status": "success"` und `"total_errors": 0` zeigen
   (das ist der einzige Zähler in der Ausgabe – nicht mit einem nicht-existierenden
   Feld wie `"errors_found"` verwechseln, jeder Eintrag in `error_summary` zählt).
2. `99_Tests` in der neu berechneten Datei lesen (z. B. mit `openpyxl`,
   `load_workbook(..., data_only=True)`) und prüfen, dass Spalte F in jeder Testzeile
   `"OK"` zeigt und `→ Gesamtstatus` = `"ALLE TESTS OK"`.
3. Wenn eine Änderung eine der neun Prüfzeilen in `99_Tests` betrifft (z. B. eine neue
   Formel in `11_Kalkulation`), das Soll dort **von Hand nachrechnen** und aktualisieren
   – nicht einfach den Ist-Wert als neuen Soll-Wert übernehmen, das degradiert den Test
   zur Tautologie.

Ein Build, der nicht diese beiden Prüfungen besteht, gilt als kaputt, unabhängig
davon, ob das Skript ohne Python-Fehler durchläuft.

## Aufbau von `build_workbook.py`

Das Skript baut die Arbeitsmappe strikt sequenziell, ein Blatt nach dem anderen, in
der Reihenfolge, in der die Blätter auch in der Datei erscheinen sollen
(`00_Anleitung` → `00_Objektindex` → `01_Annahmen` … → `99_Tests`). Zeilennummern
werden **nie hartcodiert**, sondern über einen laufenden Python-Zähler `r` (oder
Sheet-lokale Varianten wie `r2`) plus kleine Schreibhelfer ermittelt und in
`ROW_*`-Variablen bzw. Dictionaries (`OBJ`, `CAT_ROW`, `CHK_ROW_BY_ID`, `COL`,
`CATROW`, `F`) abgelegt. Das macht das Skript robust gegen Einfügungen: wenn du eine
Zeile mitten in einem Block ergänzt, verschieben sich alle nachfolgenden `ROW_*`-Werte
automatisch beim nächsten Lauf.

**Wenn du eine neue Zeile/einen neuen Parameter einfügst, halte dich an dieses
Muster** (Beispiel aus `01_Annahmen`):

```python
param("Neuer Parameter", 42, "Einheit", "A_NeuerParameter", "Quelle/Kommentar")
```

statt einer Zeilennummer eine Variable zurückzugeben und weiterzureichen. Suche nach
dem passenden Helfer (`param`, `obj_field`, `geo_field`, `krow`, `frow`, `brow`,
`erow`, `zrow`, `corow` – je nach Blatt leicht unterschiedliche, lokal definierte
Varianten desselben Musters) und benutze ihn, statt Zellen manuell mit `ws.cell(...)`
zu setzen.

### Wichtigste Helfer

- `param(label, value, unit, name, comment, fmt=None, is_pct=False)` – Zeile auf
  `01_Annahmen`, erzeugt automatisch einen **workbook-globalen** benannten Bereich
  `name` (Konvention: Präfix `A_`).
- `obj_field(key, label, value, ...)` / `geo_field(key, label, formula, ...)` – Zeile
  auf `10_VORLAGE`, erzeugt einen **blattlokalen** benannten Bereich `Obj_<key>` UND
  trägt die Zeilennummer ins `OBJ`-Dict ein.
- `obj_ref(key)` – gibt `'10_VORLAGE'!$C$<row>` zurück, für Formeln auf **anderen**
  Blättern als `10_VORLAGE` (siehe nächster Abschnitt – wichtig!).
- `chk_cell(colkey, item_id)` / `chk_col_range(colkey)` / `sum_ref(kat, col)` /
  `heizung_sumifs(colkey)` – Zugriff auf einzelne Zellen/Spalten/Summen der
  Sanierungs-Checkliste auf `10_VORLAGE`, ebenfalls für Formeln auf anderen Blättern.

### Kritische Falle: `Obj_*`-Namen sind blattlokal, nicht global

`obj_field()` und `geo_field()` legen ihre benannten Bereiche (`Obj_Wohnflaeche`,
`Obj_Baujahr`, `Obj_Zielstandard`, …) als **sheet-scoped** (lokal auf `10_VORLAGE`)
an, nicht als workbook-global. Ein Grund: sobald `10_VORLAGE` als Blattgruppe kopiert
wird, braucht jede Kopie (`10_Rahden1` usw.) ihre eigenen, unabhängigen `Obj_*`-Werte
– bei global-workbook-weiten Namen gäbe es Namenskollisionen zwischen mehreren
Objekten.

**Konsequenz:** eine Formel `=Obj_Wohnflaeche*2`, geschrieben auf `11_Kalkulation`
(oder jedem anderen Blatt außer `10_VORLAGE`), ergibt `#NAME?` beim Neuberechnen –
der lokale Name existiert dort schlicht nicht. Dieser Fehler ist mir beim Bau von
`14_Betriebskosten`/`15_Zielpreis` mehrfach passiert und musste nachträglich gefixt
werden (127 `#NAME?`-Fehler in einem Durchgang). Verwende **immer** `obj_ref("key")`
(gibt die absolute Zelladresse `'10_VORLAGE'!$C$<row>` zurück), wenn du außerhalb von
`10_VORLAGE` auf Objektdaten zugreifen willst:

```python
# FALSCH auf 14_Betriebskosten:
formula = "=Obj_Wohnflaeche*A_NB_GEG"
# RICHTIG:
formula = f"={obj_ref('Wohnflaeche')}*A_NB_GEG"
```

Nach jeder neuen Formel, die `Obj_` referenziert, lohnt sich ein Grep-Check:

```bash
grep -n "Obj_" build_workbook.py | awk -F: '$1>NNNN'   # NNNN = Zeile vor deiner Änderung
```

und alle Treffer außerhalb des `10_VORLAGE`-Blocks (Zeilen zwischen den beiden
`# 10_VORLAGE` / `# 11_Kalkulation`-Kommentaren) müssen `obj_ref(...)` oder eine
vorab aufgelöste `OR_*`-Variable sein, nie ein nackter `Obj_`-Name.

### Objekt-Blattgruppe: Namenskonvention 10_VORLAGE bis 16_CO2_Betrieb

Die sieben Objekt-Blätter heißen `10_VORLAGE`, `11_Kalkulation`, `12_Finanzierung`,
`13_Bauzeit`, `14_Betriebskosten`, `15_Zielpreis`, `16_CO2_Betrieb` – **nicht**
`11_VORLAGE` bis `16_VORLAGE`. Nur `10_VORLAGE` trägt das Wort "VORLAGE" im Namen.
Beim manuellen Kopieren in Excel (siehe README) wird deshalb nicht der String
`_VORLAGE` ersetzt, sondern das Ziffernpräfix behalten und der Rest durch die
Objekt-ID ersetzt (`11_Kalkulation` → `11_Rahden1`). Der Titel-Klammertext auf jedem
dieser Blätter (`set_title(ws, "... [Beim Kopieren: ...]")`) muss zu dieser Regel
passen – wenn du ein neues Objekt-Blatt (17_..., 18_...) hinzufügst, formuliere den
Klammertext identisch zu den bestehenden sechs.

Diese Konvention ist auch die Grundlage für `00_Objektindex` und `20_Dashboard`:
beide bauen ihre `INDIRECT()`-Formeln als `"'11_"&<ID-Zelle>&"'!C"&<Zeilennummer>"`
zusammen. Wenn du eine neue Kennzahl auf einem Objektblatt ergänzt, die im Dashboard
sichtbar sein soll, ergänze in `20_Dashboard` einen weiteren `dind(...)`-Aufruf mit
der passenden `ROW_*`-Variable statt eine Zeile manuell zu verdrahten.

## Formel-Kompatibilität (LibreOffice-Neuberechnung)

- Nur Excel-2007-Funktionen ohne Präfix verwenden: `SUMIFS`, `INDEX`, `MATCH`,
  `IFERROR`, `SUMPRODUCT`, `LOOKUP`, `INDIRECT`, `HYPERLINK`.
- **Niemals** `XLOOKUP`, `XMATCH`, `SORT`, `FILTER`, `UNIQUE`, `SEQUENCE` – LibreOffice
  in dieser Umgebung wertet sie nicht robust aus (Spill-Formeln ohne Spill-Metadaten
  ergeben nur die erste Zelle, `recalc.py` meldet trotzdem `0 Fehler`).
- Nachträglich hinzugekommene Excel-Funktionen (`TEXTJOIN`, `CONCAT`, `IFS`, `SWITCH`,
  `MAXIFS`, `MINIFS`) nur mit `_xlfn.`-Präfix schreiben, sonst `#NAME?`.
- Zellinhalte, die als **Text** gemeint sind, dürfen nicht mit `=` beginnen (auch
  keine Labels, Kommentare, Freitextzellen!) – sonst versucht Excel/LibreOffice, sie
  als Formel zu parsen, und es entsteht `#VALUE!`. Ist mir zweimal passiert
  (Kommentartext `"= 3%p.a."`, Zeilenlabels `"= Kaufpreis inkl. Nebenkosten"`) – als
  Konvention für Pfeil-Labels `"→ "` statt `"= "` verwenden.

## `catalog_data.py`

Enthält `CATALOG`, eine Liste von Dicts – eine pro Sanierungsmaßnahme (Felder: `id`,
`kat`, `name`, `einheit`, `fk`, `preis`, `nutzung`, `foerder`, `satz`, `eh55`, `eh40`,
`klima`, `el`, `kommentar`). Das Feld `fk` (Formel-Kind) steuert, welche
Mengenermittlungsformel `build_workbook.py` beim Bau der Checkliste auf `10_VORLAGE`
generiert – Zuordnung in `GENERIC_FK` (generische Fälle wie `wohnflaeche`,
`dachflaeche`, `anz_heizkoerper`) und `CONDITIONAL_OVERRIDES` (maßnahmenspezifische
Sonderfälle, z. B. Dach nur wenn Satteldach). Eine neue Maßnahme hinzufügen:

1. Zeile in `CATALOG` ergänzen (eindeutige `id`, z. B. `H13`).
2. Falls die Mengenermittlung mit einem bestehenden `fk`-Typ funktioniert: passenden
   Wert eintragen. Falls nicht: neuen Eintrag in `GENERIC_FK` oder
   `CONDITIONAL_OVERRIDES` ergänzen.
3. Neu bauen und neu berechnen – die Checkliste auf `10_VORLAGE` bekommt automatisch
   eine neue Zeile, alle nachgelagerten Summenzeilen verschieben sich automatisch mit.

## `add_object.py` — wie das automatische Anlegen funktioniert

`build/add_object.py` automatisiert den manuellen "Blattgruppe kopieren"-Schritt aus
`00_Anleitung`. Es arbeitet auf der FERTIGEN `.xlsx`-Datei (nicht auf
`build_workbook.py`) und nutzt `wb.copy_worksheet()`, um die sieben Objekt-Blätter zu
duplizieren. Das ist bewusst ein separates Skript und keine Erweiterung von
`build_workbook.py` — Letzteres baut die ganze Mappe von Grund auf neu und würde damit
alle bereits angelegten Objekte eines Nutzers löschen.

**Wichtig: `openpyxl.copy_worksheet()` kopiert weniger, als man erwartet.** Empirisch
geprüft (siehe Testreihe, die zu diesem Skript geführt hat) NICHT enthalten in einer
Kopie: blattlokale benannte Bereiche (`Obj_*`), Datenüberprüfungen/Dropdowns, bedingte
Formatierung, Diagramme, `freeze_panes`, `auto_filter`. Mitkopiert werden dagegen:
Zellwerte/-formeln, Zellformatierung, verbundene Zellen, Spaltenbreiten/Zeilenhöhen,
Reiterfarbe. `add_object.py` rekonstruiert die fehlenden Teile explizit nach dem
Kopieren (`copy_object_group()`); wenn du openpyxl aktualisierst, prüfe mit einem
kurzen `copy_worksheet()`-Test, ob sich diese Liste geändert hat, bevor du das Skript
änderst.

**Ein bereits gefundener und gefixter Bug, der sich leicht wiederholen lässt:**
`argparse` liefert jeden `--<feld>`-Wert als `str`. Ein direktes `ws.cell(value=...)`
mit diesem String schreibt eine TEXT-Zelle statt einer Zahl in die Tabelle. Das fällt
nicht sofort auf, weil `C4*A_GrESt` (Multiplikation) Text automatisch in eine Zahl
umwandelt und ein plausibles Ergebnis liefert — aber `SUM(C4:C7)` ignoriert Textzellen
komplett und lässt sie beim Summieren einfach weg, ohne Fehler zu werfen. Ergebnis:
eine leise falsche Zwischensumme, die `recalc.py` nicht als Fehler meldet (`SUM` über
eine gemischte Zahlen/Text-Range ist gültig, nur eben falsch). Deshalb definiert
`FIELD_TYPES` für jedes CLI-Feld explizit den Zieltyp (`int`/`float`/`str`), und
`fill_initial_values()` konvertiert vor dem Schreiben. Wenn du dem Skript ein weiteres
`--feld` hinzufügst: IMMER einen Eintrag in `FIELD_TYPES` ergänzen und danach mit
`data_only=True` UND `cell.data_type` prüfen, dass die Zelle wirklich `'n'` (Zahl) ist,
nicht `'s'` (Text) — ein reiner Wertevergleich reicht nicht, da `'265000' == 265000`
in Python `False`, aber in vielen schnellen Sichtprüfungen leicht übersehen wird.

**Diagramme werden geklont, nicht neu gebaut.** `rewrite_chart_refs()` läuft per
`copy.deepcopy()` über das bestehende Diagramm-Objekt der Vorlage und ersetzt nur die
Blattnamen in den `numRef.f`/`strRef.f`-Formelstrings der Datenreihen (z. B.
`"11_Kalkulation!$C$52:$C$60"` → `"11_Rahden1!$C$52:$C$60"`). Das ist robuster als das
Nachbauen der Diagramme aus den in `build_workbook.py` verwendeten Zeilenkonstanten
(`PIE_FIRST` usw.), weil es nicht mit dem Generator-Skript synchron gehalten werden
muss — jede Änderung an einem Diagramm in `build_workbook.py` landet automatisch auch
in künftig kopierten Objekten, ohne `add_object.py` anfassen zu müssen.

**Rückwärtskompatibilität:** das manuelle Kopieren aus `00_Anleitung` funktioniert
unverändert weiter und liefert exakt dasselbe Ergebnis (Excels "Blatt verschieben/
kopieren" macht intern ohnehin, was `add_object.py` explizit nachbildet: Formeln
umschreiben, Namen/Validierungen mitziehen). Wenn du `add_object.py` änderst, teste
beide Wege stichprobenartig gegeneinander (gleiche Objekt-ID, gleiche Eingabewerte,
gleicher `99_Tests`-Gesamtstatus).

**`--data <json>`: strukturierte Eingabe für Stammdaten + Checkliste.** Neben den
acht `--<feld>`-CLI-Flags (nur für schnelle manuelle Nutzung gedacht) akzeptiert
`add_object.py` eine JSON-Datei mit den Top-Level-Schlüsseln `stammdaten` (Dict,
deckt alle 28 `obj_field()`-Felder aus `10_VORLAGE` ab, siehe `STAMMDATEN_FIELDS`)
und `checkliste` (Liste von Maßnahmen-Overrides, siehe `CHK_FIELD_HEADERS`). Das ist
der Mechanismus, über den eine KI ein hochgeladenes Exposé automatisch in ein neues
Objekt überführt — die exakten Extraktions- und Mapping-Regeln dafür stehen in
**[`docs/Objekt_aus_Expose_anlegen.md`](docs/Objekt_aus_Expose_anlegen.md)**, nicht
hier, weil sie sich an die aufrufende KI richten statt an Entwickler:innen dieses
Repos. Technisch wichtig für Weiterentwicklung:
- Die Checklisten-Spalten werden **nicht** über hartkodierte Spaltenbuchstaben
  angesprochen, sondern über `find_checkliste_layout()`, das die Kopfzeile
  (Spalte B = `"ID"`) sucht und die Zielspalten per exaktem Textabgleich gegen
  `CHK_FIELD_HEADERS` findet. Änderst du `CHK_HEADERS` in `build_workbook.py`,
  muss `CHK_FIELD_HEADERS` in `add_object.py` mit den exakten (inkl. `\n`)
  Strings synchron gehalten werden, sonst bricht `add_object.py` kontrolliert mit
  einer Fehlermeldung ab (kein stiller Fehlschreib in die falsche Spalte).
- Dropdown-Felder (`keller_typ`, `dachform`, `heizung_alt`, `zielstandard`,
  `override_zeitpunkt` usw.) werden gegen die exakt gleiche Werteliste geprüft, die
  auch als Excel-Dropdown hinterlegt ist — bei einer Erweiterung der Dropdown-Liste
  in `build_workbook.py` (`obj_field(..., dropdown=[...])`) die passende Liste in
  `STAMMDATEN_FIELDS`/`OVERRIDE_ZEITPUNKT_CHOICES` in `add_object.py` mitpflegen.
- Ungültige Katalog-IDs in `checkliste` werden gegen `catalog_data.CATALOG`
  geprüft (`VALID_CATALOG_IDS`) — ein Tippfehler wie `"H5"` statt `"H05"` führt zu
  einem klaren Abbruch statt einer stillschweigend ignorierten Zeile.
- Validierungsfehler brechen **vor** `wb.save()` ab (siehe `fail()` → `sys.exit(1)`
  ohne vorherigen Save-Aufruf) — ein fehlerhafter `--data`-Lauf hinterlässt die
  Originaldatei unverändert, auch wenn die sieben Blätter im Arbeitsspeicher der
  `Workbook`-Instanz schon kopiert wurden.

## `build_demo.py` — Demo-Datei für Git

`build/build_demo.py` ruft nur `build_workbook.py` und `add_object.py` als
Subprozesse auf (keine eigene Excel-Logik) und erzeugt `demo/Hausbewertung_DEMO.xlsx`
mit zwei fiktiven Beispielobjekten (`Beispiel_Luebbecke`, `Beispiel_Rahden`). Das ist
die einzige `.xlsx`-Datei, die laut `.gitignore` committet werden darf — die echte
Ergebnis-Datei im Wurzelverzeichnis enthält reale, private Objekt-/Finanzierungsdaten
und ist deshalb generell ausgeschlossen (`*.xlsx` mit expliziter Ausnahme für die
Demo-Datei).

Technisch relevant, falls du `build_workbook.py` änderst: der Ausgabepfad wird über
die Umgebungsvariable `HAUSMODELL_OUTPUT` überschrieben (Default bleibt
`../Hausbewertung_Altkreis_Luebbecke.xlsx`, wenn die Variable nicht gesetzt ist) —
`build_demo.py` nutzt das, um in `demo/` statt ins Wurzelverzeichnis zu schreiben,
ohne die echte Datei des Nutzers anzufassen. Änderst du `OUT` in `build_workbook.py`
grundlegend (z. B. auf ein CLI-Argument statt `os.environ`), muss `build_demo.py`
entsprechend mitgezogen werden.

Wenn du Stammdaten-Felder oder Checklisten-Spalten änderst (siehe Abschnitt oben),
lauf danach `python3 build_demo.py` einmal durch und committe die aktualisierte
Demo-Datei mit — sie soll immer zum aktuellen Stand von `build_workbook.py` passen.

## Behobener Bug: Fensterfixierung machte die Checkliste auf `10_<ID>` unsichtbar

`build_workbook.py` setzte früher `ws.freeze_panes = f"C{CHK_FIRST_ROW}"` auf
`10_VORLAGE` — die Absicht war, die Kopfzeile der Sanierungs-Checkliste beim
Scrollen sichtbar zu halten. Das Problem: Excel/LibreOffice frieren beim Fixieren
IMMER alles ab Zeile 1 ein, nicht nur einen Bereich mittendrin. Da `CHK_FIRST_ROW`
bei ca. 51 liegt (kompletter Objektstammdaten-Block davor), wurde effektiv der
gesamte Stammdaten-Bereich als "fixiert" behandelt — mehr Zeilen, als in ein
normales Fenster passen. Ergebnis: für die eigentliche, scrollbare Checkliste
blieb kein sichtbarer Platz mehr, sie wirkte komplett unerreichbar (per
Nutzer-Bugreport bestätigt: "ich kann nicht herunterscrollen"). Kein Datenverlust,
reiner Anzeigefehler.

**Fix:** `freeze_panes` wird jetzt erst am Ende des `10_VORLAGE`-Blocks gesetzt,
auf `"C2"` (nur Blatttitel-Zeile 1 + Spalte A/B fixiert, siehe Kommentar im Code an
der alten Stelle). Damit lässt sich die komplette Fläche normal durchscrollen; die
Checkliste-Kopfzeile bleibt beim Scrollen nicht mehr automatisch sichtbar — das ist
der bewusst in Kauf genommene Kompromiss, da eine "nur den unteren Bereich
fixieren"-Option in Excel/openpyxl technisch nicht existiert.

**Bereits bestehende Dateien mit älteren Objekten** (angelegt, bevor dieser Fix in
`build_workbook.py`/`add_object.py` landete) haben die fehlerhafte Fixierung
weiterhin eingebacken, da `copy_worksheet()` `freeze_panes` unverändert von der
Quelle übernimmt. Es gab dafür ein einmaliges Migrationsskript
(`build/fix_checklist_freeze.py`), das ausschließlich `freeze_panes` auf jedem
`10_*`-Blatt einer bestehenden Datei korrigiert hat, ohne Zellwerte/Formeln
anzufassen — nach Gebrauch wieder entfernt, da es kein dauerhaft benötigtes
Werkzeug ist. Falls du auf eine Datei triffst, die diesen Fix noch braucht: pro
`10_*`-Blatt in Excel/LibreOffice `Ansicht → Fenster fixieren → Fixierung
aufheben`, dann Cursor auf Zelle `C2` setzen und erneut `Ansicht → Fenster
fixieren → Fenster fixieren` klicken (entspricht `freeze_panes = "C2"`) — oder
das kleine Skript bei Bedarf aus diesem Abschnitt heraus neu schreiben, der
Code-Ausschnitt oben zeigt die komplette Logik.

## Bekannte, bewusste Design-Entscheidungen (nicht versehentlich "vergessen")

- **Kein Power Query.** `openpyxl` kann keine echten Datenverbindungen erzeugen; die
  ursprünglich im Bauplan (`docs/Bauplan_Hausbewertung_Excel.md`, Abschnitt 13.1)
  vorgesehene `KPI_<ID>`-Tabellen-Aggregation wurde durch `INDIRECT()`-Formeln auf
  `00_Objektindex`/`20_Dashboard` ersetzt.
- **Kein natives Excel-Datatable** auf `15_Zielpreis` (Sensitivitätstabelle) – stattdessen
  eine Matrix aus 12 einzeln geschriebenen Formeln, weil `TABLE()`-Arrayformeln bei
  `openpyxl`/LibreOffice-Neuberechnung nicht zuverlässig funktionieren.
- **Kein Blattschutz.** Würde das manuelle "Blattgruppe kopieren" beim Anlegen neuer
  Objekte stören/verkomplizieren.
- **Konstante Annuität** in `12_Finanzierung`, nur der Zinssatz wechselt nach
  Zinsbindungsablauf – vereinfachtes, aber transparentes Modell (Restschuld am
  Laufzeitende wird explizit ausgewiesen, siehe `ROW_RESTSCHULD_ENDLAUFZEIT`).
- **Bauzeitkosten (`13_Bauzeit`) fließen bewusst nicht automatisch** in die
  Finanzierungssumme auf `12_Finanzierung` ein, um eine Zirkelbeziehung
  (Bereitstellungszins hängt vom Kreditbetrag ab, der wiederum vom
  Finanzierungsbedarf abhängt, der die Bauzeitkosten enthalten würde) zu vermeiden.

Wenn du eine dieser Entscheidungen änderst, aktualisiere den entsprechenden Absatz in
`README.md` mit.

## Wo der fachliche Hintergrund steht

`docs/Bauplan_Hausbewertung_Excel.md` enthält die ursprüngliche Anforderungsklärung,
Marktrecherche (Förderprogramme, KfW-Konditionen, Baupreise, Bodenrichtwerte) und die
Formel-Herleitung pro Blatt, wie sie vor dem Bau der Datei erarbeitet wurde. Die
Datei ist über `.gitignore` bewusst nicht Teil dieses Repos (Recherche-Rohfassung,
nicht für Veröffentlichung gedacht) – falls sie in deiner lokalen Kopie fehlt, sind
die aufbereiteten Design-Entscheidungen trotzdem vollständig im Abschnitt oben und
in `README.md` festgehalten. Bei fachlichen Fragen (warum wird X so berechnet?),
sofern die Datei lokal vorhanden ist, dort zuerst nachsehen, bevor du die
Formel änderst – und `90_Quellen` in der Excel-Datei selbst für den aktuellen
Verlässlichkeits-Status jeder Annahme.

`docs/Objekt_aus_Expose_anlegen.md` enthält den Ablauf und das JSON-Schema, nach dem
eine KI ein hochgeladenes Exposé automatisiert in ein neues Objekt (`add_object.py
--data ...`) überführt — vollständige Feld-Mapping-Tabelle, Ableitungsregeln
("was darf aus einem Exposé übernommen werden, was nicht") und Regeln für den
Umgang mit fehlenden/unsicheren Angaben (nichts erfinden, lieber Default stehen
lassen und dem Nutzer melden).
