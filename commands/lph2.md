---
name: lph2
description: Erstellt einen vollständigen LPH 2 Vorentwurfsbericht (HOAI § 33, Gebäude und Innenräume) für ein beliebiges Architekturprojekt. Durchsucht alle Dateien im Projektordner, extrahiert Daten und befüllt das standardisierte HTML-Template.
user-invocable: true
---

# LPH 2 Vorentwurfsbericht — Automatische Generierung

## OUTPUT-KONVENTION (verbindlich)

**WICHTIGSTE REGEL: Es wird ein EINZIGER Ordner namens `Claude/` im Projekt-Root erzeugt. ALLE Outputs landen dort drin. NIRGENDS sonst im Projekt.**

```
<Projekt-Root>/                           <- bleibt unveraendert
├── (die existierenden Projektordner)       (welche Struktur auch immer)
│
└── Claude/                               <- EINZIGER Ordner den wir anlegen
    ├── Projekt_Historie.md                 <- kumulative Wissensbasis
    ├── LPH_1/LPH1_Report_YYYY-MM-DD.html
    ├── LPH_2/LPH2_Report_YYYY-MM-DD.html
    ├── ... bis LPH_8/
    ├── Dashboard/Projekt_Dashboard_YYYY-MM-DD.html
    └── Fachplaner-Analyse/Fachplaner_Analyse_YYYY-MM-DD.md
```

**Dein konkreter Output-Pfad fuer diesen Skill:**
```
<Projekt-Root>/Claude/LPH_2/LPH2_Report_YYYY-MM-DD.html
```

**Pflichten beim Schreiben:**
1. Wenn `<Projekt-Root>/Claude/` nicht existiert: anlegen (mkdir)
2. Wenn `<Projekt-Root>/Claude/LPH_2` nicht existiert: anlegen
3. Datums-Format im Dateinamen: `YYYY-MM-DD` (z.B. 2026-05-27)
4. Bei jedem Aufruf: NEUE Datei mit aktuellem Datum (keine Ueberschreibung)

**NIEMALS schreibst du:**
- In andere Ordner des Projekts (00 Aktennotiz/, 11 Doku/, etc.) - keine Verschmutzung der Projekt-Struktur
- In den Projekt-Root direkt - nur in `Claude/` Unterstruktur
- In `~/.claude/plugins/` - das ist Read-only Plugin-Code
- In versteckte Ordner (.claude/) - nutze ausschliesslich `Claude/` (sichtbar)

---

## WISSENSBASIS NUTZEN

Falls `<Projekt-Root>/Claude/Projekt_Historie.md` existiert:
- ZUERST diese Datei lesen - sie enthaelt chronologisch alle E-Mails, Aktennotizen, Beschluesse
- Spart Zeit + Tokens (statt das ganze Projekt jedes Mal neu zu scannen)

Falls die Datei NICHT existiert:
- Empfehle dem User: "Tipp: Fuer aktuelle Daten zuerst /projekt-historie ausfuehren"
- Trotzdem das Projekt direkt scannen (Fallback)

---


## Template
Suche das Template relativ zum aktuellen Projektordner — KEIN hardcodierter Pfad:
```powershell
# Template finden:
Get-ChildItem -Path "." -Recurse -Filter "LPH2_Template.html" | Select-Object -First 1 -ExpandProperty FullName
```
Das Template liegt üblicherweise in `Claude_DOKU\LPH2_Template.html` (aus dem _Claude_DOKU_Starter).
Kopiere es als `Claude_DOKU\LPH2_Report.html` und befülle alle `{{PLATZHALTER}}`.
Falls nicht gefunden: Benutzer informieren dass das Template fehlt.

## Ablauf — 5 Schritte

### Schritt 1: Projektordner durchsuchen
Durchsuche den GESAMTEN Projektordner rekursiv:
- **PDFs**: Lies JEDE Seite jeder PDF (Glob `**/*.pdf`, dann Read)
- **DOCX**: Lies per PowerShell Word COM (`New-Object -ComObject Word.Application`)
- **XLSX**: Lies per PowerShell Excel COM (`New-Object -ComObject Excel.Application`)
- **Bilder**: Finde Logo (GHIW/Gutthann/Logo im Namen) und Titelbild (Perspektive/Visu/Lageplan)
- **Jour Fixe PDFs**: Suche nach `*JourFixe*`, `*Jour*Fixe*`, `*PS-JF*`, `*FPJF*`
- **Beteiligtenliste**: Suche nach `*Beteiligte*` oder `*Projektbeteiligte*`

WICHTIG: DOCX und XLSX können NICHT mit dem Read-Tool gelesen werden. Verwende IMMER PowerShell COM.

### Schritt 2: Daten extrahieren
Extrahiere aus allen gelesenen Dateien:
- **Beteiligte**: Namen, Firmen, Kontakte (NUR aus Beteiligtenliste, NICHTS erfinden)
- **Flächen**: BGF, BRI, NUF je Bauteil
- **Raumprogramm**: Soll/Ist-Vergleich je Raum
- **Kosten**: Kostenrahmen, Kostenschätzung DIN 276, KGR 200-700, Kostentreiber
- **Termine**: Rahmenterminplan, Meilensteine, Bauabschnitte
- **Brandschutz**: Konzept, offene Punkte
- **Barrierefreiheit**: Anforderungen, Maßnahmen
- **TGA**: Lüftung, Heizung, Elektro, Bauphysik
- **B-Plan**: GRZ, WH, Vollgeschosse, Bauweise
- **Beschlüsse aus Jour Fixe**: Entscheidungen, offene Punkte, Action Items

### Schritt 3: Template kopieren und befüllen
1. Kopiere `LPH2_Template.html` in den Projektordner als `LPH2_Report.html`
2. Ersetze ALLE `{{PLATZHALTER}}` mit den extrahierten Daten
3. Erstelle HTML-Tabellen für Tabellen-Platzhalter (Raumprogramm, Kosten, etc.)
4. Befülle die Sidebar-Dateiliste mit allen tatsächlich gelesenen Dateien

### Schritt 4: Zielkonflikte analysieren
Analysiere ALLE gelesenen Dokumente und erstelle Zielkonflikte-Tabelle:
- Format: Kategorie | Zielvorgabe | Konflikt ("xxx vs. xxx") | Relevanz | Risiko | Empfehlung
- Mindestens 8-10 Konflikte identifizieren
- Risiko-Klassen: `class="risk-high"` (rot), `class="risk-mid"` (orange), `class="risk-low"` (grau)

### Schritt 5: Qualitätsprüfung
- Keine `{{PLATZHALTER}}` mehr im Output
- Alle Zahlen gegen Quelldokumente prüfen
- Schreibstil: Sachlich, technisch, Passiv, unpersönlich ("ist einzuhalten", "wird geprüft")
- Max. 5-8 Sätze pro Absatz
- Beteiligte nur indirekt ("seitens AG", "durch Fachplaner")

## Design-Regeln (Lessons-Learned aus (Projekt-Nr.) — verbindlich)
- Farbe: SCHWARZ `#1a1a1a` — kein Grün, kein Teal, kein Blau
- **Schwarze Banner-Balken (ch-label, ch-num, ch-title) ALLE einheitlich `#1a1a1a`** — niemals `#333` oder andere Graustufen. Trennung der Felder über weiße 1px-Borderlines.
- **Print-CSS PFLICHT für ALLE dunklen Elemente:**
  `-webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;`
  Gilt für: `.chapter-banner`, `.ch-label`, `.ch-num`, `.ch-title`, `.toc-header`, `th`, `.risk-tag` (alle Status), `.lph-pip.active`, `.goal-card-header`
- **Goal-Card-Header (9 Zielkarten 01-09): NIEMALS weiße Schrift auf schwarzem BG.**
  PFLICHT: schwarzer Text auf hellem BG (`#fafafa`/`#f5f5f5`/`#f0f0f0`) mit:
  - 5px schwarzem Akzentstreifen links (`border-left:5px solid #1a1a1a`)
  - 2px Border-bottom (`border-bottom:2px solid #1a1a1a`)
- **Lange Gedankenstriche (`&mdash;` / —) in ÜBERSCHRIFTEN durch kurze Bindestriche `-` ersetzen.**
  Gilt für: proj-title, proj-sub, toc-header, sec-heading, ch-title, goal-card-header.
  In Fließtext und Tabellenzellen sind `&mdash;` OK (grammatikalisch korrekt).
- Alle Felder `contenteditable` für nachträgliche Bearbeitung
- PDF-Druck: A4 Querformat, randlos, keine Schatten, keine Checklisten
- Zielvorgaben als 3x3 Karten-Grid (`.goal-grid` / `.goal-card`)

## Inhalts-Regeln (Lessons-Learned aus (Projekt-Nr.) — verbindlich)
- **Pendelliste / Jour Fixe** (z.B. `Aktennotiz_Pendelliste.xlsx`) ist die WICHTIGSTE Inhaltsquelle.
  Lies sie zwingend per PowerShell Excel COM aus. Daraus stammen die aktuellsten Erkenntnisse seit Vertragsabschluss.
- **Bestandsplan-Quelle EXAKT klären**: kommen Pläne vom Bauherr (z.B. Landratsamt / Stadt / Gemeinde) oder vom Projektsteuerer? Im Zweifel beim User nachfragen — NIEMALS raten.
- **Projektsteuerer ist NICHT automatisch Plan-Lieferant** — Rolle sauber trennen.

## Struktur-Regeln (Lessons-Learned aus (Projekt-Nr.) — verbindlich)
- **Tabelle 2.8 Meilensteine: KEINE Spalte "Nr."** — nur Meilenstein/Phase, Termin, Status.
- **QUELLEN-Block unter JEDER Checkliste** in der rechten Sidebar (`.side-col`):
  ```html
  <div class="source-block">
    <strong>QUELLEN</strong>
    • [Dateiname + Datum + Herkunft]
    • [Quelle 2 ...]
  </div>
  ```
  CSS:
  ```css
  .source-block { font-size:7pt; color:#444; margin-top:10px; padding-top:6px; border-top:1px dashed #999; line-height:1.5; }
  .source-block strong { display:block; margin-bottom:3px; color:#1a1a1a; letter-spacing:.3px; font-size:7.5pt; }
  ```
  Jede Quelle muss enthalten: Dateiname / Plan-Nr + Datum + Herkunft.
- **Risiken**: IMMER als `<span class="risk-tag risk-XXX">label</span>` — nicht als reine Textfarbe.

## Pre-Flight Checklist (vor "fertig" abhaken)
- [ ] Alle 9 Goal-Cards lesbar (schwarzer Text auf hellem BG mit Akzentstreifen)?
- [ ] Schwarze Banner-Balken einheitlich `#1a1a1a` (kein `#333`)?
- [ ] Print-CSS hat `print-color-adjust:exact !important` für alle dunklen Elemente?
- [ ] Tabelle 2.8 Meilensteine OHNE Spalte "Nr."?
- [ ] Quellen-Block (`.source-block`) unter jeder Checkliste vorhanden?
- [ ] Lange Gedankenstriche (—) in Überschriften durch "-" ersetzt?
- [ ] Stand-Datum aktuell?
- [ ] Bestandsplan-Quelle korrekt (Bauherr vs. Projektsteuerer geklärt)?
- [ ] Jour Fixe / Pendelliste-Inhalte eingebaut?
- [ ] Beteiligtenliste vollständig + nur dort aufgeführte Firmen verwendet?

## Platzhalter-Referenz

| Platzhalter | Beschreibung |
|---|---|
| `{{PROJEKT_TITEL}}` | z.B. "NEUBAU GRUNDSCHULE MIT SPORTHALLE UND KITA" |
| `{{PROJEKT_NR}}` | z.B. "(Projekt-Nr.)-25 STKR" |
| `{{PROJEKT_ORT}}` | z.B. "{ORTSNAME}" |
| `{{DATUM}}` | Aktuelles Datum TT.MM.JJJJ |
| `{{TITELBILD_PFAD}}` | Relativer Pfad zum Titelbild |
| `{{AG_NAME}}` | Name des Auftraggebers |
| `{{2_1_1_VORGABEN_LPH1}}` | Fließtext Vorgaben aus LPH 1 |
| `{{2_1_2_GRUNDLAGENANALYSE_TEXT}}` | Einleitungstext |
| `{{2_1_2_GRUNDLAGEN_TABELLE}}` | HTML `<table>` mit Grundlagen |
| `{{2_1_3_RAUMPROGRAMM_TEXT}}` | Einleitungstext Raumprogramm |
| `{{2_1_3_RAUMPROGRAMM_TABELLE}}` | HTML `<table>` Soll/Ist |
| `{{2_1_4_BGF_BRI_TABELLE}}` | HTML `<table>` BGF/BRI |
| `{{ZIEL_01_GESTALTERISCH}}` bis `{{ZIEL_09_TERMINE}}` | Bullet-Points je Zielkarte |
| `{{2_2_2_ZIELKONFLIKTE_ZEILEN}}` | HTML `<tr>` Zeilen für Konflikte |
| `{{2_3_1_BETEILIGTE_TABELLE}}` | HTML `<table class="proj-table">` |
| `{{2_3_2_NUTZUNGSKONZEPT}}` | Fließtext je Bauteil |
| `{{2_3_3_RAUMPROGRAMM_TABELLEN}}` | Detaillierte Raumtabellen |
| `{{2_4_1_STAEDTEBAU}}` bis `{{2_4_3_TGA}}` | Fließtext je Abschnitt |
| `{{2_5_ARBEITSERGEBNISSE_TEXT}}` | Einleitungstext |
| `{{2_5_PLANVERZEICHNIS_TABELLE}}` | HTML `<table>` Planverzeichnis |
| `{{2_6_GENEHMIGUNG}}` | Fließtext Genehmigungsfähigkeit |
| `{{2_7_KOSTEN_TEXT}}` | Einleitungstext Kosten |
| `{{2_7_KOSTEN_TABELLEN}}` | HTML Kostentabellen DIN 276 |
| `{{2_8_TERMINPLAN_TEXT}}` | Einleitungstext |
| `{{2_8_TERMINPLAN_TABELLEN}}` | HTML Terminplan-Tabellen |
| `{{2_9_1_OFFENE_BEAUFTRAGUNGEN_TABELLE}}` | HTML Beauftragungstabelle |
| `{{2_9_2_OFFENE_PUNKTE_TABELLE}}` | HTML Offene-Punkte-Tabelle |
| `{{ABSCHLUSS_TEXT}}` | Abschlusstext LPH 2 |
| `{{DATEILISTE}}` | `<li>` Elemente für Sidebar |
| `{{CHECKLISTE_2_X}}` | Checkbox-Items je Kapitel |
