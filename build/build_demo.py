# -*- coding: utf-8 -*-
"""
build_demo.py — baut eine Demo-Arbeitsmappe mit zwei Beispielobjekten.

Ruft ausschließlich die vorhandenen Skripte auf (build_workbook.py, add_object.py) -
enthält selbst keine Excel-Logik. Ergebnis: demo/Hausbewertung_DEMO.xlsx mit den
Beispielobjekten "Beispiel_Luebbecke" (EH 55 EE, Gasheizung) und "Beispiel_Rahden"
(EH 40 EE, Ölheizung, umfangreichere Sanierung).

Die Demo-Datei ist bewusst dazu da, ins Git-Repo zu wandern (siehe .gitignore -
sie ist von der generellen Ausnahme für *.xlsx ausgenommen), damit auf GitHub direkt
sichtbar ist, wie ein befülltes Objekt aussieht, ohne echte private Kaufobjektdaten
preiszugeben.

Nutzung:
  python3 build_demo.py
"""
import json
import os
import subprocess
import sys
import tempfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DEMO_DIR = os.path.join(REPO_ROOT, "demo")
DEMO_XLSX = os.path.join(DEMO_DIR, "Hausbewertung_DEMO.xlsx")

# Zwei fiktive Beispielobjekte - keine echten Kaufobjekte, nur zur Illustration des
# Modells. Feldnamen/Schema wie in docs/Objekt_aus_Expose_anlegen.md beschrieben.
DEMO_OBJECTS = [
    (
        "Beispiel_Luebbecke",
        {
            "stammdaten": {
                "adresse": "Musterstraße 12",
                "plz": 32312,
                "ort": "Lübbecke",
                "baujahr": 1982,
                "kaufpreis": 279000,
                "wohnflaeche": 138,
                "grundstuecksflaeche": 680,
                "bodenrichtwert": 105,
                "geschosse": 2,
                "keller_typ": "Vollkeller",
                "dachform": "Satteldach",
                "dach_eternit": "Nein",
                "energieausweis": 165,
                "heizung_alt": "Gasheizung",
                "zielstandard": "EH 55 EE",
                "anz_baeder": 2,
                "makler_ja_nein": "Ja",
            },
            "checkliste": [
                {"id": "H05", "zustand": 3, "kommentar": "Demo: Fenster laut Exposé noch original, ca. 1982"},
                {"id": "A01", "override_zeitpunkt": "Sofort",
                 "kommentar": "Demo: Gasheizung soll direkt durch Wärmepumpe ersetzt werden"},
                {"id": "R03", "menge_override": 0, "kommentar": "Demo: kein Nachtspeicherofen vorhanden"},
            ],
        },
    ),
    (
        "Beispiel_Rahden",
        {
            "stammdaten": {
                "adresse": "Bahnhofstraße 4",
                "plz": 32369,
                "ort": "Rahden",
                "baujahr": 1968,
                "kaufpreis": 195000,
                "wohnflaeche": 122,
                "grundstuecksflaeche": 590,
                "bodenrichtwert": 85,
                "geschosse": 1,
                "keller_typ": "Teilkeller",
                "dachform": "Walmdach",
                "dach_eternit": "Ja",
                "energieausweis": 210,
                "heizung_alt": "Ölheizung",
                "zielstandard": "EH 40 EE",
                "anz_baeder": 1,
                "makler_ja_nein": "Nein",
            },
            "checkliste": [
                {"id": "R01", "override_zeitpunkt": "Sofort",
                 "kommentar": "Demo: Baujahr <1994 + Eternit-Dach laut Exposé, Asbestverdacht"},
                {"id": "R04", "override_zeitpunkt": "Sofort", "kommentar": "Demo: Ölheizung, Tank stilllegen"},
                {"id": "H04", "zustand": 4, "kommentar": "Demo: Fassade ungedämmt laut Exposé-Fotos"},
                {"id": "H10", "override_zeitpunkt": "Entfällt", "kommentar": "Demo: nur Teilkeller vorhanden"},
            ],
        },
    ),
]


def run(cmd, **kwargs):
    print(f"  $ {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=SCRIPT_DIR, **kwargs)
    if result.returncode != 0:
        print(f"FEHLER: Befehl ist mit Exit-Code {result.returncode} fehlgeschlagen: {' '.join(cmd)}",
              file=sys.stderr)
        sys.exit(result.returncode)


def main():
    os.makedirs(DEMO_DIR, exist_ok=True)
    if os.path.exists(DEMO_XLSX):
        os.remove(DEMO_XLSX)

    print(f"1. Baue leere Arbeitsmappe (build_workbook.py) nach {DEMO_XLSX} ...")
    env = dict(os.environ)
    env["HAUSMODELL_OUTPUT"] = DEMO_XLSX
    run([sys.executable, "build_workbook.py"], env=env)

    tmp_files = []
    try:
        for i, (obj_id, data) in enumerate(DEMO_OBJECTS):
            is_last = i == len(DEMO_OBJECTS) - 1
            fd, tmp_path = tempfile.mkstemp(suffix=f"_{obj_id}.json", dir=SCRIPT_DIR)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            tmp_files.append(tmp_path)

            print(f"\n2.{i + 1} Lege Beispielobjekt '{obj_id}' an ...")
            cmd = [sys.executable, "add_object.py", obj_id,
                   "--file", DEMO_XLSX, "--data", tmp_path, "--no-backup"]
            if not is_last:
                # Neuberechnung erst nach dem letzten Objekt - spart Zeit, das Ergebnis
                # ist danach ohnehin für die ganze Datei identisch aktuell.
                cmd.append("--no-recalc")
            run(cmd)
    finally:
        for p in tmp_files:
            if os.path.exists(p):
                os.remove(p)

    print(f"\nFertig: {DEMO_XLSX}")
    print("Diese Datei ist für Git gedacht (siehe .gitignore) - sie enthält nur die "
          "beiden fiktiven Beispielobjekte oben, keine echten Kaufobjektdaten.")


if __name__ == "__main__":
    main()
