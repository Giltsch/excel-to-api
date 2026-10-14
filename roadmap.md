# Roadmap — Everything to API

## Vision

Jede strukturierte oder semi-strukturierte Datenquelle soll mit minimalem Aufwand in eine vollständige, produktionsreife REST-API verwandelt werden können — inklusive Datenmodell, Validierung, CRUD-Endpunkte und Dokumentation.

---

## Phase 1 — Excel to API ✅ (aktueller Stand)

- [x] Excel-Upload über Streamlit UI
- [x] Regelbasierte Analyse (Typ-Erkennung, Constraints, Beziehungen)
- [x] Optionale LLM-Verifikation & -Verbesserung via OpenRouter
- [x] Vollständige FastAPI-Projektgenerierung (Models, Schemas, CRUD, DB, Seed-Daten)
- [x] Live-Preview des generierten API-Servers
- [x] Download als ZIP-Projekt
- [x] Docker-Support
- [x] PK-Erkennung für verschiedene Namenskonventionen (`category_id`, `categoryid`, `CategoryID`)
- [x] Englische & deutsche Singular/Plural-Erkennung (`categories` ↔ `category`, `Produkte` ↔ `Produkt`)
- [x] FK-Spalten referenzieren den echten PK-Spaltennamen der Zieltabelle
- [x] LLM-Fehler werden sichtbar in der UI angezeigt (nicht mehr still verschluckt)

### Offene Verbesserungen (Phase 1)

- [ ] Bessere Erkennung von zusammengesetzten Primary Keys
- [ ] Support für mehrere Excel-Dateien gleichzeitig
- [ ] Konfigurierbare Naming-Conventions (camelCase, snake_case etc.)
- [ ] Authentifizierung / API-Key-Generierung als optionales Feature
- [ ] Paginierung & Filter-Endpunkte in generierten APIs

---

## Phase 2 — Robuste Excel-Extraktion & Tabellarische Formate

Ziel: Auch "unsaubere" Excel-Dateien zuverlässig verarbeiten + alle gängigen tabellarischen Formate unterstützen.

### Excel-Bereinigung & Sonderfälle

- [ ] **Merged Cells** — Zusammengeführte Zellen erkennen und automatisch auffüllen
- [ ] **Header-Erkennung** — Header nicht immer in Zeile 1 (z.B. nach Titel / Leerzeilen / Logos)
- [ ] **Mehrere Tabellen pro Sheet** — Erkennung von Sub-Tabellen innerhalb eines Blattes
- [ ] **Diagramme & Bilder ignorieren** — Charts, eingebettete Bilder und Zeichnungen sauber überspringen
- [ ] **Leere Zeilen/Spalten** — Trailing/Leading-Leerzeilen und komplett leere Spalten entfernen
- [ ] **Formeln auflösen** — Berechnete Zellen als Werte übernehmen (nicht als Formel-String)
- [ ] **Datentyp-Konflikte** — Spalten mit gemischten Typen (z.B. Zahlen + Text) intelligent behandeln
- [ ] **Named Ranges / Tabellen** — Excel-Tabellen (`ListObject`) und benannte Bereiche als Datenquelle nutzen
- [ ] **Kommentar-Zeilen** — Zeilen mit Anmerkungen / Fußnoten erkennen und ausschließen
- [ ] **Encoding & Sonderzeichen** — Umlaute, Unicode, kaputte Encoding-Artefakte bereinigen
- [ ] **Multi-Sheet-Vorschau** — Im UI pro Sheet anzeigen: "Sieht das korrekt aus?" mit Korrekturmöglichkeit

### Weitere tabellarische Formate

- [ ] **CSV** — Typ-Inferenz aus Werten, Delimiter-Erkennung
- [ ] **TSV** — analog zu CSV
- [ ] **JSON** (flach/tabellarisch) — Array of Objects → API
- [ ] **YAML** — analog zu JSON
- [ ] **Google Sheets** — Integration via Google Sheets API

### Technische Aufgaben

- [ ] Gemeinsames Parser-Interface (`DataSource`) für alle Formate
- [ ] Abstraktion des Analyzers von Excel-spezifischer Logik
- [ ] Format-Erkennung (Auto-Detect anhand Extension & Inhalt)
- [ ] Einheitliches Fehler-Handling & Validierungs-Feedback im UI
- [ ] LLM-gestützte Bereinigung: "Wo fängt die eigentliche Tabelle an?"

---

## Phase 3 — Dokument-basierte Quellen

Ziel: Semi-strukturierte Daten aus Dokumenten extrahieren.

- [ ] **Word (.docx)** — Tabellen aus Word-Dokumenten extrahieren
- [ ] **PowerPoint (.pptx)** — Tabellen & strukturierte Daten aus Slides
- [ ] **PDF** — Tabellenerkennung (z.B. via `pdfplumber` oder `camelot`)
- [ ] **HTML-Tabellen** — Scraping von Web-Tabellen via URL

### Technische Aufgaben

- [ ] LLM-gestützte Strukturerkennung für mehrdeutige Formate
- [ ] Preview / manuelle Korrektur der erkannten Struktur im UI
- [ ] Vertrauensscore pro extrahierter Tabelle / Spalte

---

## Phase 4 — Datenbank-Quellen

Ziel: Bestehende Datenbanken als API exponieren.

- [ ] **SQLite** — Direkte Schema-Analyse
- [ ] **PostgreSQL** — Connection-String-basiert
- [ ] **MySQL / MariaDB** — Connection-String-basiert
- [ ] **MongoDB** — Schema-Inference aus Dokumenten

### Technische Aufgaben

- [ ] DB-Schema-Introspection → `AnalysisResult`-Mapping
- [ ] Automatische Relationship-Erkennung aus Foreign Keys
- [ ] Optionale Read-Only-API-Generierung (kein Write auf Prod-DBs)
- [ ] Connection-Pooling & Sicherheitskonfiguration

---

## Phase 5 — Knowledge Graph für komplexe Daten

Ziel: Unstrukturierte und komplex vernetzte Daten als Graph-API bereitstellen.

- [ ] **Text → Knowledge Graph** — LLM-gestützte Entitäts- & Relationsextraktion
- [ ] **Graph-Datenmodell** — Nodes & Edges statt nur flache Tabellen
- [ ] **GraphQL-API-Generierung** — als Alternative zu REST für vernetzte Daten
- [ ] **Neo4j / ArangoDB Export** — optionaler Graph-DB-Export

### Anwendungsfälle

- Verträge, Berichte, Protokolle → vernetzte API
- Organigramme, Beziehungsnetzwerke
- Komplexe Multi-Sheet-Excel-Dateien mit vielen Verknüpfungen

---

## Querschnittsthemen (laufend)

- [ ] **UI/UX** — Drag & Drop Upload, Fortschrittsanzeige, interaktiver Schema-Editor
- [ ] **Testing** — Unit-Tests für Analyzer, Generator und alle Parser
- [ ] **CI/CD** — GitHub Actions Pipeline für Tests und Docker-Build
- [ ] **Dokumentation** — Nutzer-Doku, API-Doku, Architektur-Doku
- [ ] **Deployment** — One-Click-Deploy auf Railway / Fly.io / Render
- [ ] **Plugin-System** — Erweiterbare Parser & Generatoren für Community-Beiträge
- [ ] **Multi-Output** — Neben FastAPI auch Django REST, Express.js, Go Fiber etc.