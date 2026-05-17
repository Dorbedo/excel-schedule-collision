# excel-schedule-collision

Excel VBA Tool für die Kollisionsanalyse in Abhängigkeit von Layout und Terminen.

## Struktur

```
excel-schedule-collision.xlsm   ← Haupt-Excel-Datei (makroaktiviert)
src/
  Main.bas                      ← Einstiegspunkt (Button/Ribbon)
  CollisionChecker.bas          ← Kernlogik Kollisionserkennung
docs/                           ← Dokumentation
```

## Verwendung

1. `excel-schedule-collision.xlsm` in Excel öffnen
2. Makros aktivieren
3. Kollisionsprüfung über den vorgesehenen Button starten

## VBA-Module importieren

Um `.bas`-Dateien aus `src/` in die Excel-Datei zu importieren:

1. VBA-Editor öffnen (`Alt + F11`)
2. Datei → Datei importieren → `.bas`-Datei auswählen

## Lizenz

MIT — freie Nutzung, Modifikation und Weitergabe. Namensnennung erforderlich.
