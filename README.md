# Hausbewertung NRW

Excel-Modell zur Bewertung gebrauchter Einfamilienhäuser vor dem Kauf, mit Fokus auf
Sanierung zu einem sehr energieeffizienten / möglichst klimaneutralen Standard
(Effizienzhaus 55 EE / 40 EE). Ausgelegt auf Objekte in Nordrhein-Westfalen (Annahmen
wie Grunderwerbsteuer 6,5 % und Regionalfaktor sind NRW-weit gültig; die mitgelieferten
Beispielobjekte liegen in Lübbecke und Rahden, lassen sich aber für jeden NRW-Standort
anpassen).

Das Modell hat zwei Ebenen, die bewusst unterschiedlich gepflegt werden:

- **Struktur** (Formeln, Blattaufbau, Maßnahmenkatalog, Annahmen – "die Engine")
  wird **nicht von Hand in Excel gepflegt**, sondern ausschließlich aus den
  Python-Skripten unter `build/` generiert; die `.xlsx`-Datei ist hier nur ein
  Build-Artefakt, das Skript ist die Quelle der Wahrheit. Grund: die Datei enthält
  ~2.000 durchnummerierte Formeln, Hand-Edits daran sind im Git-Diff nicht
  nachvollziehbar und brechen das Muster fast garantiert. Wenn du an Formeln,
  Blättern oder dem Katalog etwas ändern willst, passiert das im Skript, dann wird
  die Datei komplett neu gebaut (siehe "Arbeitsmappe neu bauen" unten).
- **Deine Objektdaten** (Adresse, Kaufpreis, Zustand der einzelnen
  Sanierungsmaßnahmen, Kommentare, …) sind genau umgekehrt: dafür sind die
  **gelben Zellen** auf jedem `10_<ID>`-Blatt gedacht (Farblegende auf
  `00_Anleitung`), und die trägst du ganz normal **von Hand direkt in
  Excel/LibreOffice** ein – egal ob das Objektblatt per Skript oder manuell
  angelegt wurde. `add_object.py --data` (Variante C unten) ist nur eine optionale
  Abkürzung, um diese gelben Zellen automatisiert vorzubefüllen; von Hand
  eintragen/ändern bleibt jederzeit möglich und ist für alles, was z. B. ein
  Exposé nicht hergibt, sogar der Normalfall.

Es gibt bewusst **keinen Blattschutz** (siehe "Bekannte Einschränkungen" unten), du
könntest also technisch auch Formelzellen überschreiben – vorgesehen ist das aber
nicht: weiße/grüne Zellen sind Formeln und sollten unangetastet bleiben, nur gelb
ist zum Editieren gedacht.

> **Hinweis (Stand 20.08.2026):** Dieses Modell enthält Annahmen zu KfW-/BEG-
> Fördersätzen, -konditionen und -deckeln (z. B. BEG-EM, KfW 261/458), Bauzins-
> Näherungswerten und Energiepreisen, die sich **nach der Erstellung dieses Repos
> ändern können** – Förderprogramme werden erfahrungsgemäß mehrfach jährlich
> angepasst oder auch ausgesetzt/neu aufgelegt. Alle diese Annahmen sind mit Quelle
> und Datum im Blatt `90_Quellen` (in der Excel-Datei) dokumentiert – vor jeder
> Verwendung dort den Stand prüfen und im Zweifel gegen die
> aktuellen Angaben von [kfw.de](https://www.kfw.de) / [foerderdatenbank.de](https://www.foerderdatenbank.de)
> gegenchecken. Dieses Repo ist kein Ersatz für eine Finanzierungs-, Steuer- oder
> Energieberatung.

## Voraussetzungen

- Python 3.9+ mit [`openpyxl`](https://pypi.org/project/openpyxl/) (`pip install -r build/requirements.txt`).
- Eine lokale LibreOffice-Installation (`soffice`/`libreoffice` auf dem PATH) für
  `build/recalc.py` – wird von `build_workbook.py`, `add_object.py` und
  `build_demo.py` zur Neuberechnung der Formeln aufgerufen. Ohne LibreOffice
  funktionieren die Skripte trotzdem (mit einem Hinweis in der Konsole), nur muss
  die erzeugte `.xlsx`-Datei dann einmal manuell in Excel/LibreOffice geöffnet und
  gespeichert werden, damit die Formeln berechnete Werte bekommen.

## Inhalt des Repos

```
Hausbewertung_NRW.xlsx                  # generierte Excel-Datei (Build-Artefakt, NICHT in Git)
README.md                               # diese Datei
AGENTS.md                               # Leitfaden für KI-Agenten / Entwickler:innen
LICENSE                                 # MIT-Lizenz
.gitignore                              # schließt Ergebnis-Datei(en)/Backups/Caches von Git aus
build/
  requirements.txt                      # Python-Abhängigkeiten (openpyxl)
  build_workbook.py                     # Hauptskript: baut die komplette Arbeitsmappe
  catalog_data.py                       # Stammdaten: 51 Sanierungsmaßnahmen mit Preisen etc.
  add_object.py                         # legt ein neues Objekt (Haus) automatisch an
  build_demo.py                         # baut demo/Hausbewertung_DEMO.xlsx (siehe unten)
  recalc.py                             # berechnet Formeln per LibreOffice headless neu
docs/
  Bauplan_Hausbewertung_Excel.md         # ursprünglicher Rechercheplan, NICHT in Git (siehe unten)
  Objekt_aus_Expose_anlegen.md           # Anleitung für die KI: Objekt automatisch aus Exposé anlegen
demo/
  Hausbewertung_DEMO.xlsx               # Demo-Datei mit 2 fiktiven Beispielobjekten (IST in Git)
```

Die echte Ergebnis-Datei im Wurzelverzeichnis enthält reale Kaufobjekt- und
Finanzierungsdaten und ist deshalb über `.gitignore` von Git ausgeschlossen (`*.xlsx`,
mit expliziter Ausnahme für `demo/Hausbewertung_DEMO.xlsx`). Wer das Repo klont, baut
sie sich lokal selbst (siehe "Arbeitsmappe neu bauen" unten). `docs/Bauplan_Hausbewertung_Excel.md`
ist ebenfalls über `.gitignore` ausgeschlossen – das ist die ungefilterte
Recherche-Rohfassung von vor dem Bau des Modells; die aufbereiteten Inhalte daraus
(Design-Entscheidungen, Quellenlage) stehen in `AGENTS.md`, in diesem README und im
Blatt `90_Quellen` der Excel-Datei.

### Demo-Datei erzeugen/aktualisieren

```bash
cd build
python3 build_demo.py
```

Ruft nacheinander `build_workbook.py` (mit Ausgabepfad `demo/Hausbewertung_DEMO.xlsx`
über die Umgebungsvariable `HAUSMODELL_OUTPUT`) und zweimal `add_object.py --data …`
für die fiktiven Beispielobjekte `Beispiel_Luebbecke` (EH 55 EE, Gasheizung) und
`Beispiel_Rahden` (EH 40 EE, Ölheizung, Asbestverdacht) auf. Ergebnis ist eine fertig
durchgerechnete Demo-Arbeitsmappe ohne echte Objektdaten, die problemlos ins
Git-Repo kann und auf GitHub direkt zeigt, wie ein befülltes Objekt aussieht.

## Was das Modell kann

- **01_Annahmen**: ca. 90 globale Parameter (Steuern, Zinsen, Baupreise, Energiepreise,
  CO2-Faktoren, Geometrie-Faustformeln) als benannte Bereiche (`A_...`).
- **02_Massnahmenkatalog / 03_Foerderprogramme / 04_Nutzungsdauern / 05_Energiekennwerte**:
  Stammdaten – 51 Sanierungsmaßnahmen (Hülle, Anlagentechnik, Elektro, Risiko/Schadstoff,
  Klimaneutralität, Smart Home/KNX, Außenanlagen, weiche Kosten), Fördersätze,
  Nutzungsdauern, Baualtersklassen.
- **10_VORLAGE bis 16_CO2_Betrieb** (Objekt-Blattgruppe, pro Haus einmal kopiert):
  Objektstammdaten, Sanierungs-Checkliste mit automatischer Fälligkeits- und
  Förderfähigkeits-Berechnung, Gesamtinvestition, Förderweg-Vergleich (Effizienzhaus vs.
  Einzelmaßnahmen), Finanzierung mit Jahres-Tilgungsplan über zwei Tranchen, Liquidität
  während der Bauzeit, Energiekosten vorher/nachher, Rückwärtsrechnung der
  Kaufpreisobergrenze (20 % günstiger als Neubau), CO2-Bilanz im Betrieb.
- **00_Objektindex / 20_Dashboard**: Übersicht aller angelegten Objekte mit
  automatischem Kennzahlenvergleich – siehe Abschnitt "Wie Mehrfach-Objekte funktionieren".
- **90_Quellen**: Herkunft und Verlässlichkeit aller unsicheren Annahmen.
- **99_Tests**: 9 unabhängig von Hand nachgerechnete Kontrollwerte (Kaufnebenkosten,
  Maßnahmen-Fälligkeit, Förderrechnung, Tilgungsplan, Zielpreis) – müssen nach jeder
  Änderung "OK" zeigen.

## Neue Objekte (Häuser) anlegen

### Variante A: per Skript (empfohlen)

```bash
cd build
python3 add_object.py Rahden1
```

Das Skript kopiert die sieben Objekt-Blätter, benennt sie um, schreibt alle internen
Querverweise und Diagramme um, rekonstruiert Dropdowns/bedingte Formatierung/lokale
Namen (siehe Abschnitt "`add_object.py` — wie das automatische Anlegen funktioniert"
in `AGENTS.md` für Details) und trägt die Objekt-ID auf `00_Objektindex` ein.
Optional gleich mit Startwerten:

```bash
python3 add_object.py Rahden1 --adresse "Bahnhofstr. 12" --plz 32369 --ort Rahden \
    --baujahr 1985 --kaufpreis 265000 --wohnflaeche 145 \
    --grundstuecksflaeche 600 --bodenrichtwert 95
```

Vor dem Speichern wird automatisch ein Backup der Datei angelegt
(`Hausbewertung_NRW.xlsx.bak-<Zeitstempel>`), danach läuft standardmäßig
die LibreOffice-Neuberechnung (`--no-recalc` überspringt das, `--no-backup` das Backup).
`python3 add_object.py --help` zeigt alle Optionen. Statt einzelner `--<feld>`-Flags
können auch alle Stammdaten + Sanierungs-Checkliste auf einmal aus einer JSON-Datei
kommen (`--data <datei>.json`) – siehe Variante C. Restliche Objektdaten danach auf
`10_<ID>` von Hand ausfüllen.

### Variante B: manuell (z. B. ohne Python)

Die Kurzfassung steht auch auf dem Blatt `00_Anleitung` in der Datei selbst:

1. Objekt-ID vergeben (kurz, keine Sonderzeichen, z. B. `Rahden1`).
2. Die sieben Blätter `10_VORLAGE, 11_Kalkulation, 12_Finanzierung, 13_Bauzeit,
   14_Betriebskosten, 15_Zielpreis, 16_CO2_Betrieb` gemeinsam markieren (Strg+Klick auf
   die Reiter) → Rechtsklick → "Blatt verschieben/kopieren…" → "Kopie erstellen".
   Excel passt dabei automatisch alle internen Querverweise zwischen den kopierten
   Blättern an.
3. Die sieben neuen Blätter umbenennen: Ziffernpräfix behalten, Rest durch die Objekt-ID
   ersetzen (`11_Kalkulation (2)` → `11_Rahden1` usw.).
4. Objektdaten auf `10_Rahden1` eintragen (gelbe Zellen), Sanierungs-Checkliste
   durchgehen.
5. Auf `00_Objektindex` in der nächsten freien Zeile **nur die Objekt-ID** eintragen.

### Variante C: Exposé hochladen und von einer KI anlegen lassen

Ein Immobilien-Exposé (PDF/Bild/Text) direkt in einen Claude-Chat laden und bitten,
daraus ein neues Objekt anzulegen. Die KI liest dann `docs/Objekt_aus_Expose_anlegen.md`
(dort auch aus `AGENTS.md` verlinkt), extrahiert die im Exposé vorhandenen Werte in
eine JSON-Datei nach dem dort beschriebenen Schema und ruft:

```bash
python3 add_object.py <ID> --data <ID>.json
```

Ein Exposé liefert selten alle Felder – nicht befüllte Werte bleiben bewusst auf dem
Default aus `10_VORLAGE` stehen, statt geraten zu werden. Die KI meldet am Ende, was
automatisch gesetzt wurde und was noch von Hand zu ergänzen ist (v. a.
Bodenrichtwert, Energieausweis-Wert und die Sanierungs-Checkliste).

## Arbeitsmappe neu bauen

```bash
cd build
python3 build_workbook.py
```

Das Skript schreibt `../Hausbewertung_NRW.xlsx` neu. `openpyxl` schreibt
Formeln nur als Text, ohne berechnete Werte – deshalb danach **immer** neu berechnen
(ruft im Hintergrund eine lokale LibreOffice-Installation headless auf, siehe
Voraussetzung unten):

```bash
python3 recalc.py ../Hausbewertung_NRW.xlsx 150
```

Das Ergebnis muss `"status": "success"` und `"total_errors": 0` zeigen. Danach
`99_Tests` in der Datei öffnen (oder per `openpyxl`/`data_only=True` auslesen) und
prüfen, dass `→ Gesamtstatus` = `ALLE TESTS OK` ist.

**Voraussetzung für `recalc.py` und für `add_object.py` (das `recalc.py`
automatisch mitaufruft):** eine lokale LibreOffice-Installation (`soffice` bzw.
`libreoffice` muss auf dem PATH liegen) – z. B. `apt install libreoffice-calc`
(Debian/Ubuntu), `brew install --cask libreoffice` (macOS) oder der Installer von
[libreoffice.org](https://www.libreoffice.org/download/) (Windows). Ohne
LibreOffice bricht `recalc.py` mit einer klaren Fehlermeldung ab; Abhilfe ist dann,
die generierte Datei einmal manuell in Excel oder LibreOffice zu öffnen und zu
speichern – das berechnet alle Formeln ebenfalls einmalig durch.

### Checkliste auf `10_<ID>` nicht sichtbar / lässt sich nicht herunterscrollen?

Das war ein bekannter, mittlerweile behobener Bug: `build_workbook.py` fror
versehentlich den kompletten Objektstammdaten-Block ein, wodurch für die
Sanierungs-Checkliste darunter kein sichtbarer Platz mehr blieb (siehe
`AGENTS.md`, Abschnitt "Behobener Bug: Fensterfixierung machte die Checkliste
auf `10_<ID>` unsichtbar"). Neu angelegte Objekte (per Skript oder manuellem
Kopieren aus einer aktuell gebauten Datei) haben dieses Problem nicht mehr. Bei
einer älteren Datei mit betroffenen Objekten: pro `10_*`-Blatt in Excel
`Ansicht → Fenster fixieren → Fixierung aufheben`, Cursor auf Zelle `C2`,
dann erneut `Fenster fixieren`.

Details, Stolpersteine und Konventionen für Weiterentwicklung: siehe `AGENTS.md`.

## Bekannte Einschränkungen (bewusste Design-Entscheidungen)

- **Kein Blattschutz.** Damit "Blattgruppe kopieren" beim Anlegen neuer Objekte nicht
  versehentlich blockiert wird, sind keine Zellen gesperrt. Farblegende auf
  `00_Anleitung` beachten (nur gelbe Zellen editieren).
- **Kein VBA/Makro.** Bewusste Entscheidung für Portabilität (Excel Windows/Mac/Online
  und LibreOffice) – Objekte werden manuell per "Blattgruppe kopieren" angelegt.
- **Konstante Annuität in der Finanzierung.** `12_Finanzierung` hält die Rate über die
  gesamte Laufzeit konstant und wechselt nur den Zinssatz nach Ablauf der Zinsbindung.
  Das kann bei Zinsanstieg zu einer Restschuld am Laufzeitende führen (wird als eigene
  Kennzahl ausgewiesen) – realistischer wäre eine Neuberechnung der Rate bei
  Anschlussfinanzierung, das würde aber den Rahmen eines Kalkulationsmodells sprengen.
- **Sensitivitätstabelle auf `15_Zielpreis`** ist eine fest berechnete Matrix, kein
  natives Excel-Datatable (`TABLE()`-Arrayformeln sind bei `openpyxl`/LibreOffice-
  Neuberechnung unzuverlässig).
- Mehrere Annahmen sind unverifiziert oder grobe Näherungen (KfW-Zinssätze,
  Fernwärmepreise, Regionalfaktor OWL, Haushaltsstrombedarf) – siehe `90_Quellen` für
  den vollständigen Prüfhinweis-Katalog vor einer echten Kaufentscheidung.

## Lizenz / Nutzung

Veröffentlicht unter der [MIT-Lizenz](LICENSE). Keine Rechts- oder Steuerberatung –
alle Annahmen vor einer echten Finanzierungs- oder Kaufentscheidung mit Bank,
Steuerberater:in bzw. Energieberater:in gegenprüfen.
