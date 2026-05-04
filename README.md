# CompExcel - Java Excel Comparison Tool

Dieses Tool vergleicht zwei Excel-Dateien miteinander und markiert die Unterschiede in einer neuen Datei namens `diff.xlsx`.

## 🚀 Features

* **Vergleich von Inhalten:** Findet abweichende Werte zwischen zwei Excel-Tabellen.
* **Automatischer Export:** Generiert eine übersichtliche Ergebnisdatei (`diff.xlsx`).
* **Java-basiert:** Läuft plattformunabhängig (Windows, macOS, Linux).

## 🛠 Voraussetzungen

* **Java JDK 11** oder höher.
* **Apache POI Library** (5.5.1).
* Ein Build-Tool wie **Maven** oder **Gradle** (empfohlen).

## 📋 Installation & Vorbereitung

1. Klone das Repository:
   ```bash
   git clone https://github.com
   cd compexcel
   ```

2. Stelle sicher, dass die benötigten Abhängigkeiten (Apache POI) in deiner `pom.xml` oder `build.gradle` vorhanden sind.

## 💻 Verwendung

1. Lege die beiden zu vergleichenden Dateien in das Projektverzeichnis.
2. Starte das Programm über deine IDE oder die Konsole:
   
   Führe die Klasse "ExcelComparator" aus.
   ```
3. Nach Abschluss des Vorgangs findest du die Datei `diff.xlsx` im Hauptordner.

## 📂 Projektstruktur

* `src/`: Quellcode des Java-Programms.
* `diff.xlsx`: Die generierte Datei mit den Unterschieden (nach Ausführung).

---
Erstellt von [brpwalter](https://github.com)
