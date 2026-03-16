# Python Programm zur Bestimmung der Optionspreisentwicklung mit Daten aus der IB Trader Workstation 

* Du holst Dir die notwendigen Daten wie z.B. aktueller Kurs der Aktie, Volatilität, Zinssatz aus der Schnittstelle zur IB Trader Workstation
* Nutze TKinker zur Darstellung
* Erstelle eine venv, lade die notwendigen Bibliotheken 
* Dokumentiere den Sourcecode
* Zähle die Versionsnummer des Programms mit jedem Änderungsdurchlauf um 0.01 hoch, zeige die Versionsnummer oben in der Titelleiste
* Speicher das Programm am Ende einer Session bei Github, lege hierzu beim ersten Mal ein Projekt mit dem Namen IB_OptionspreisBerechner an
    Nutze hier die Datei: Github_Token.txt, Benutzername peterh69
* Erstelle eine Readme.txt wo beschrieben wird wie man das Progamm einrichtet und startet. Speicher die Datei auch auf Github 

1) Lade die Tickersymbole aus der Watchlist EUR und erstelle ein Auswahlfeld zur Auswahl der weiter zu untersuchenden Aktie 

2) Von der ausgewählten Aktie
	* Lade den aktuellen Kurs
	* Lade die aktuelle imp. Vol. 
	* Lade den aktuellen Zinsatz

3) Erstelle eine Tabelle mit den Optionspreisen für einen Short Put mit einer Laufzeit von bis zu 60 Tagen, Schrittweite wöchentlich. Benutze das richtige Datum, ausgehend vom aktuellen Datum der Rechneruhr. 
    * Strike Preis: Ausgehen vom aktuellen Kurs bis zu -10% vom aktuellen Kurs. Benutze die gleichen Strike Werte wie in der IB Trader Workstation zu dem gewälten Ticker vorgegeben. Vergleiche in der Tabelle den berechneten Wert mit dem bei IB angegebenen Optionspreis. 
