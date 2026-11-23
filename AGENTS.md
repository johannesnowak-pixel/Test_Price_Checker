# Projekt: Preis-Finder Agent

## Ziel
Dieses Projekt sucht automatisch nach Kaufangeboten im Internet für alle Suchbegriffe aus der Datei `produkte.txt` und speichert sie in einer Excel-Datei.

## Regeln für den Agenten
- Nutze Python und die Bibliotheken `pandas` und `openpyxl` für die Excel-Erstellung.
- Lese die Suchbegriffe immer dynamisch aus der Datei `produkte.txt` aus (schreibe niemals Produktnamen fest in den Code).
- Speichere das Ergebnis immer als `ergebnisse/angebote.xlsx` mit genau diesen Spalten: Suchbegriff, Produkt-Info, Preis, Verkäufer, Link, Datum.
- Erstelle den Ordner `ergebnisse` automatisch, falls er noch nicht existiert.
- Führe nach jeder Änderung am Code einen Testlauf durch, um zu prüfen, ob die Excel-Datei fehlerfrei erstellt wird.
