# Neues Objekt aus einem Exposé anlegen (Anleitung für die KI)

Diese Datei richtet sich an eine KI (z. B. Claude), der/die in einem Chat ein
Immobilien-Exposé (PDF, Bild oder Text) hochgeladen bekommt und daraus automatisch
ein neues Objekt in `Hausbewertung_Altkreis_Luebbecke.xlsx` anlegen soll. Sie ist aus
`AGENTS.md` verlinkt und ergänzt dessen Abschnitt "`add_object.py` — wie das
automatische Anlegen funktioniert" um die Exposé-spezifischen Schritte.

**Grundprinzip: lieber ehrlich lückenhaft als falsch vollständig.** Ein Exposé
liefert selten alle ~28 Stammdaten-Felder und praktisch nie belastbare Werte für die
51 Sanierungsmaßnahmen der Checkliste. Die KI darf **keine Zahl erfinden**, um ein
Feld "vollständig" aussehen zu lassen. Fehlt ein Wert im Exposé, bleibt das Feld
unausgefüllt (Default aus `10_VORLAGE` bleibt stehen) und wird am Ende in der
Zusammenfassung an den Nutzer als "manuell zu prüfen" gemeldet.

## Exakter Ablauf

1. **Exposé lesen und Rohdaten extrahieren.** Alle Zahlen, Adressen und
   Beschreibungen aus dem Dokument sammeln, bevor irgendetwas geschrieben wird.
   Bei eingescannten/Bild-PDFs den Text so sorgfältig wie möglich transkribieren;
   bei Unleserlichkeit lieber das Feld weglassen als raten.
2. **Werte auf das JSON-Schema unten abbilden** (siehe Tabellen "Stammdaten
   (`--data` → `stammdaten`)" und "Checkliste (`--data` → `checkliste`)"). Nur
   Felder aufnehmen, für die im Exposé eine belastbare Angabe steht oder die sich
   eindeutig herleiten lassen (siehe Ableitungsregeln unten).
3. **Objekt-ID vergeben**: kurz, keine Sonderzeichen (`/ \ ? * [ ] :`), maximal 28
   Zeichen. Empfehlung: `<Ort><laufende Nummer>`, z. B. `Rahden1`, `Espelkamp3`.
   Vor der Vergabe prüfen, ob die ID nicht schon in der Datei existiert
   (`grep -o '"1[0-6]_[A-Za-z0-9]*"' ...` oder einfach die Blattnamen der
   Arbeitsmappe auflisten).
4. **JSON-Datei schreiben**, z. B. nach `build/<obj_id>.json` (Datei wird nicht
   committed, ist nur ein Übergabeformat für diesen einen Lauf).
5. **Skript ausführen**:
   ```bash
   cd build
   python3 add_object.py <obj_id> --data <obj_id>.json
   ```
   Das Skript legt automatisch ein Backup an, kopiert die sieben Objekt-Blätter,
   befüllt Stammdaten + Checkliste aus dem JSON, trägt die ID in
   `00_Objektindex` ein und ruft danach die LibreOffice-Neuberechnung auf.
6. **Ergebnis prüfen, bevor es an den Nutzer gemeldet wird:**
   - Konsolenausgabe von `add_object.py` auf `FEHLER:`-Zeilen prüfen (Exit-Code
     ungleich 0 bedeutet: nichts wurde gespeichert, siehe unten "Fehlerverhalten").
   - Die recalc-Ausgabe am Ende muss `"status": "success"` und
     `"total_errors": 0` zeigen. Falls nicht: nicht raten, sondern die
     gemeldeten Zellen mit `openpyxl` (`data_only=True`) ansehen.
   - `99_Tests` prüfen (Zelle mit Label "→ Gesamtstatus" muss `ALLE TESTS OK`
     zeigen) — diese Tests hängen nur an `10_VORLAGE`, ein neues Objekt kann sie
     also nicht kaputt machen; schlagen sie trotzdem fehl, ist etwas grundsätzlich
     schiefgelaufen (z. B. eine Formel versehentlich verändert) und muss vor der
     Auslieferung geklärt werden.
   - Auf `20_Dashboard` die neue Zeile stichprobenartig gegen das Exposé
     gegenlesen (Adresse, Kaufpreis, Wohnfläche).
7. **Dem Nutzer klar zurückmelden:**
   - Welche Felder aus dem Exposé automatisch gesetzt wurden (kurze Liste).
   - Welche relevanten Felder **nicht** aus dem Exposé hervorgingen und daher auf
     dem Default stehen geblieben sind (z. B. Bodenrichtwert, Energieausweis-Wert,
     Zustand der Checkliste) — das sind die Punkte, die der Nutzer als Nächstes
     von Hand auf `10_<ID>` nachträgt.
   - Etwaige Annahmen, die die KI selbst treffen musste (z. B. "Baujahr auf
     Checkliste = Baujahr des Hauses, da keine Angabe zu 'letzte Erneuerung'
     im Exposé stand").
   - Die Datei ist danach unverändert weiter über die üblichen Wege
     (`SendUserFile` + Commit ins verbundene Gerätefolder) auszuliefern.

## JSON-Schema

```json
{
  "stammdaten": {
    "adresse": "Bahnhofstraße 12",
    "plz": 32369,
    "ort": "Rahden",
    "baujahr": 1985,
    "kaufpreis": 265000,
    "wohnflaeche": 145,
    "grundstuecksflaeche": 612,
    "geschosse": 2,
    "keller_typ": "Vollkeller",
    "dachform": "Satteldach",
    "heizung_alt": "Gasheizung",
    "anz_baeder": 2,
    "makler_ja_nein": "Ja"
  },
  "checkliste": [
    {
      "id": "H05",
      "baujahr": 1985,
      "zustand": 3,
      "kommentar": "Laut Exposé Doppelverglasung, kein Baujahr-Update bekannt"
    }
  ]
}
```

Beide Top-Level-Schlüssel sind optional (nur `stammdaten`, nur `checkliste`, oder
beides). Jedes Feld innerhalb von `stammdaten` ist optional — nur belegte Felder
werden geschrieben, alles andere bleibt auf dem `10_VORLAGE`-Default. Unbekannte
Schlüssel lässt `add_object.py` nicht stillschweigend durch, sondern bricht mit
einer klaren Fehlermeldung ab (Tippfehler-Schutz).

### Stammdaten (`--data` → `stammdaten`)

| JSON-Feld | Excel-Label (Spalte B auf `10_<ID>`) | Typ | Erlaubte Werte / Format | Typische Exposé-Formulierung |
|---|---|---|---|---|
| `adresse` | Adresse | Text | frei | Straße + Hausnummer aus Kopfzeile/Anschrift |
| `plz` | PLZ | Ganzzahl | 5-stellig | aus Anschrift |
| `ort` | Ort | Text | frei | aus Anschrift |
| `baujahr` | Baujahr | Ganzzahl | z. B. 1900–2026 | "Baujahr", "erbaut", "Bezugsfertig" |
| `kaufpreis` | Kaufpreis (Angebot) | Zahl | € | "Kaufpreis", "Angebotspreis", "VB" |
| `wohnflaeche` | Wohnfläche | Zahl | m² | "Wohnfläche ca. X m²" |
| `grundstuecksflaeche` | Grundstücksfläche | Zahl | m² | "Grundstücksfläche", "Grundstück ca." |
| `bodenrichtwert` | Bodenrichtwert | Zahl | €/m² | steht so gut wie nie im Exposé — **nicht** aus dem Kaufpreis zurückrechnen, sondern weglassen und den Nutzer auf BORIS-NRW verweisen |
| `anpassungsfaktor` | Anpassungsfaktor Grundstücksgröße | Zahl | Faktor, Default 1.00 | praktisch nie im Exposé — weglassen |
| `geschosse` | Geschosse | Ganzzahl | z. B. 1–3 | "Vollgeschosse", "1,5-geschossig" (dann 2 ansetzen, in Meldung an Nutzer erwähnen) |
| `keller_typ` | Keller | Text | `Kein Keller` \| `Teilkeller` \| `Vollkeller` | "Vollkeller", "unterkellert", "kein Keller" |
| `dachform` | Dachform | Text | `Satteldach` \| `Flachdach` \| `Walmdach` | "Satteldach", "Flachdach", "Walmdach" — bei anderen Formen (Krüppelwalm, Pultdach) die ähnlichste Kategorie wählen und das in der Rückmeldung an den Nutzer explizit vermerken |
| `dach_eternit` | Dacheindeckung Eternit/Asbestverdacht | Text | `Ja` \| `Nein` | nur setzen, wenn explizit erwähnt ("Faserzement", "Asbest", "Eternit"); sonst weglassen, nicht "Nein" raten |
| `energieausweis` | Energieausweis-Wert (0 = kein Ausweis) | Zahl | kWh/m²a | Endenergiebedarf/-verbrauch aus dem Energieausweis-Feld des Exposés |
| `heizung_alt` | Heizungsart aktuell | Text | `Gasheizung` \| `Ölheizung` \| `Nachtspeicher` \| `Fernwärme` \| `Wärmepumpe` \| `Sonstige` | "Gasheizung", "Ölheizung", "Nachtspeicheröfen", "Fernwärme", "Wärmepumpe" |
| `anz_nachtspeicher` | Anzahl Nachtspeicheröfen | Ganzzahl | Stück | nur wenn explizit gezählt/erwähnt |
| `feuchte_keller` | Feuchteschaden im Keller | Text | `Ja` \| `Nein` | nur setzen, wenn im Exposé/Zustandsbericht ausdrücklich erwähnt |
| `zielstandard` | Zielstandard | Text | `EH 55 EE` \| `EH 40 EE` \| `Nur Einzelmaßnahmen` | steht nie im Exposé (das ist eine Entscheidung des Nutzers) — **nicht setzen**, Default `EH 55 EE` bleibt stehen |
| `isfp_ja_nein` | iSFP-Bonus nutzen | Text | `Ja` \| `Nein` | Nutzerentscheidung — nicht aus dem Exposé ableitbar, weglassen |
| `foerderweg_wahl` | Förderweg-Wahl | Text | `Automatik (Empfehlung)` \| `Weg A (Effizienzhaus)` \| `Weg B (Einzelmaßnahmen)` | Nutzerentscheidung — weglassen |
| `makler_ja_nein` | Makler beteiligt | Text | `Ja` \| `Nein` | ergibt sich meist daraus, ob das Exposé von einem Maklerbüro stammt bzw. eine Provisionsangabe enthält |
| `anz_baeder` | Anzahl Bäder | Ganzzahl | Stück | "Bäder", "Bad/WC" |
| `personen_override` | Personen im Haushalt (0=Default) | Ganzzahl | Anzahl | nicht aus dem Exposé ableitbar — weglassen |
| `sanierungsdauer_monate` | Sanierungsdauer | Ganzzahl | Monate | Planungsannahme des Nutzers — weglassen |
| `einzug_nach_sanierung` | Einzug erst nach Sanierung | Text | `Ja` \| `Nein` | Nutzerentscheidung — weglassen |
| `parallelmiete_monat` | Miete während Sanierung | Zahl | €/Monat | Nutzerangabe — weglassen |
| `pv_kwp_override` | PV Zielgröße (0 = automatisch) | Zahl | kWp | nur setzen, wenn eine PV-Anlage mit bekannter Leistung bereits vorhanden ist und übernommen werden soll |
| `pv_ausrichtung` | PV Ausrichtung | Text | `Süd` \| `Ost-West` | nur wenn Dachausrichtung im Exposé beschrieben ist |

Zahlen dürfen im JSON als echte JSON-Zahl (`265000`) oder als String mit Punkt
oder Komma (`"265000"`, `"265.000,00"` funktioniert **nicht** — deutsche
Tausenderpunkte vor dem Schreiben ins JSON entfernen, nur Dezimalkomma/-punkt ist
erlaubt) übergeben werden. Am saubersten: im JSON immer echte Zahlen ohne
Tausendertrennzeichen verwenden.

### Checkliste (`--data` → `checkliste`)

Jeder Eintrag bezieht sich auf **eine Zeile** der Sanierungs-Checkliste
(Maßnahme aus `02_Massnahmenkatalog` / `build/catalog_data.py`) und braucht ein
`id`-Feld (z. B. `"H05"`). Gültige IDs: `H01`–`H12` (Hülle), `A01`–`A08`
(Anlagentechnik), `E01`–`E10` (Elektro/Innenausbau), `R01`–`R06`
(Risiko/Schadstoff), `K01`–`K04` (Klimaneutralität), `S01`–`S03` (Smart Home),
`G01`–`G04` (Außenanlagen), `W01`–`W04` (weiche Kosten) — die genaue Liste mit
Beschreibung steht in `build/catalog_data.py`.

| JSON-Feld | Spalte in der Checkliste | Typ | Bedeutung |
|---|---|---|---|
| `id` | ID | Text (Pflicht) | Katalog-ID der Maßnahme |
| `menge_override` | Menge Override | Zahl | überschreibt die automatisch berechnete Menge (z. B. Dachfläche in m²); nur setzen, wenn das Exposé eine genauere Zahl liefert als die Geometrie-Faustformel, oder um eine Maßnahme faktisch auf 0 zu setzen (z. B. "kein Nachtspeicherofen vorhanden") |
| `baujahr` | Baujahr/letzte Erneuerung | Ganzzahl | Baujahr der jeweiligen Komponente, falls abweichend vom Baujahr des Hauses (z. B. "Fenster 2015 erneuert", "Heizung 2018 neu") — steuert die Restlebensdauer-Berechnung |
| `zustand` | Zustandsnote (1-5) | Ganzzahl 1–5 | 1 = sehr gut/neuwertig, 5 = sehr schlecht/erneuerungsbedürftig; nur setzen, wenn das Exposé einen belastbaren Hinweis liefert (z. B. "neue Fenster", "sanierungsbedürftiges Dach"); ohne konkreten Hinweis lieber weglassen (Default 3 bleibt stehen) als raten |
| `override_zeitpunkt` | Override Zeitpunkt | Text | `Automatik` \| `Sofort` \| `In Horizont` \| `Später` \| `Entfällt`; nur setzen, wenn das Exposé eine explizite Dringlichkeit nahelegt (z. B. "Heizung muss zeitnah ersetzt werden" → `Sofort`) oder eine Maßnahme laut Exposé eindeutig nicht zutrifft (z. B. kein Keller vorhanden → betroffene Kellermaßnahme `Entfällt`) |
| `eigenleistung_pct` | Eigenleistung % | Zahl 0–1 | Anteil in Eigenleistung (0,2 = 20 %); praktisch nie aus dem Exposé ableitbar — weglassen, außer der Nutzer hat das vorab in der Konversation festgelegt |
| `kommentar` | Kommentar | Text | Freitext-Begründung, warum der Wert gesetzt wurde (z. B. Zitat/Paraphrase aus dem Exposé) — **immer** mitgeben, wenn eines der obigen Felder gesetzt wird, damit der Nutzer die Herkunft nachvollziehen kann |

**Welche Maßnahmen sind aus einem Exposé überhaupt sinnvoll befüllbar?**
In der Praxis liefert ein Exposé kaum je Zustandsdetails auf Maßnahmenebene.
Realistisch ableitbar sind meist nur:
- `R03` (Nachtspeicherofen) und `R04` (Öltank) auf `menge_override=0` bzw.
  `override_zeitpunkt="Entfällt"`, wenn das Exposé explizit eine andere
  Heizungsart nennt.
- `H05` (Fenster) oder `A01`/`A02` (Wärmepumpe), wenn im Exposé ausdrücklich
  "neue Fenster [Jahr]" oder "bereits Wärmepumpe vorhanden" steht — dann eher
  `baujahr`/`zustand` als aktuell statt als 3/Default setzen.
- Kellermaßnahmen (`H09`, `H10`, `R05`), wenn das Exposé "kein Keller" angibt
  → `override_zeitpunkt="Entfällt"` mit Begründung im Kommentar.

Für alle übrigen Maßnahmen gilt: **nichts eintragen.** Die Checkliste bleibt auf
den Defaults aus `10_VORLAGE` stehen und wird vom Nutzer beim Vor-Ort-Termin bzw.
nach Sichtung des Energieausweises von Hand ergänzt — das ist der Normalfall,
kein Fehler des Skripts.

## Fehlerverhalten

`add_object.py` validiert `--data` strikt und bricht mit `FEHLER: ...` und
Exit-Code 1 ab, **bevor** irgendetwas in die Datei geschrieben wird, wenn:
- die JSON-Datei kein gültiges JSON ist,
- ein unbekannter Top-Level-Schlüssel (etwas anderes als `stammdaten`/
  `checkliste`) vorkommt,
- ein unbekanntes Stammdaten-Feld verwendet wird,
- ein Dropdown-Feld (z. B. `dachform`, `heizung_alt`) einen nicht erlaubten Wert
  enthält,
- eine Checklisten-`id` nicht im Katalog existiert,
- `zustand` außerhalb von 1–5 oder `eigenleistung_pct` außerhalb von 0–1 liegt,
- ein Zahlenfeld einen nicht parsbaren Wert enthält (z. B. Text mit
  Tausenderpunkt).

Die KI sollte bei einem solchen Fehler **nicht** versuchen, den fehlerhaften Wert
stillschweigend zu reparieren und stumpf erneut auszuführen, sondern die genaue
Fehlermeldung lesen, den betroffenen Wert im Exposé noch einmal prüfen und
entweder korrigieren oder bewusst weglassen.

## Siehe auch

- `AGENTS.md`, Abschnitt "`add_object.py` — wie das automatische Anlegen
  funktioniert" (technischer Hintergrund: was `copy_worksheet()` nicht mitkopiert,
  die argparse-Text-vs-Zahl-Falle, Diagramm-Klonen).
- `README.md`, Abschnitt "Neue Objekte (Häuser) anlegen" (Variante A/B für den
  manuellen bzw. CLI-Gebrauch ohne JSON).
- `build/catalog_data.py` (vollständige Liste der 51 Sanierungsmaßnahmen mit IDs).
