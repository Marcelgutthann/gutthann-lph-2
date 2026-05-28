# LPH 2 - Vorentwurfsplanung (Vorentwurf) - Plugin fuer Claude Code

> Automatische Erstellung von LPH 2 Berichten nach HOAI Paragraph 33 (Gebaeude und Innenraeume) - generisch fuer Architektur- und Planungsbueros.

## Was dieses Plugin macht

Wenn du in einem Projektordner `/lph2` eintippst, durchsucht Claude alle Projekt-Dateien (PDFs, DOCX, XLSX, MSG) und erstellt daraus einen kompletten LPH 2 Bericht als HTML.

## Installation

```
/plugin marketplace add github:Marcelgutthann/gutthann-claude-marketplace
/plugin install lph-2
```

Nach der Installation: Claude Code einmal neu starten.

## Nutzung

1. Im Projektordner Claude Code oeffnen
2. `CLAUDE.md` anlegen mit Projektdaten (Projektnummer, AG, Ort)
3. Pflicht-Dateien ablegen (siehe `templates/LPH2_Informationsliste.md` falls vorhanden)
4. `/lph2` eintippen
5. Claude erzeugt `LPH2_Report.html` im Projektordner

## Templates

| Datei | Zweck |
|---|---|
| `templates/LPH2_Template.html` | HTML-Template (druckfest, mit Platzhaltern) |
| `templates/LPH2_Checkliste.md` | Pflicht-Checkliste fuer LPH 2 |

## Eingebaute Lessons-Learned (verbindlich)

- **Druckfest**: Print-CSS mit `print-color-adjust:exact` fuer ALLE dunklen Elemente
- **Quellen-Belege**: QUELLEN-Block unter jeder Checkliste mit Datei + Datum + Herkunft
- **Goal-Cards druckbar**: schwarzer Text auf hellem BG mit Akzentstreifen
- **Kurze Bindestriche** in Ueberschriften
- **Bauherr vs. Projektsteuerer** klar getrennt
- **Pre-Flight Checklist** vor jedem fertig
