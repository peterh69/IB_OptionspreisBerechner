ALV Short-Put Optionsrechner
============================

Dieses Programm berechnet theoretische Short-Put-Optionspreise fuer die Allianz-Aktie
(ALV, XETRA) auf Basis des Black-Scholes-Merton-Modells. Der aktuelle Kurs wird
optional von der Interactive Brokers Trader Workstation (TWS) abgerufen.


VORAUSSETZUNGEN
---------------
- Python 3.10 oder hoeher
- Interactive Brokers TWS oder IB Gateway (fuer Live-Kurse)


EINRICHTUNG
-----------
1. Virtuelle Umgebung erstellen und aktivieren:

   Linux / macOS:
       python3 -m venv .venv
       source .venv/bin/activate

   Windows:
       python -m venv .venv
       .venv\Scripts\activate

2. Abhaengigkeiten installieren:

       pip install -r requirements.txt


IB TWS KONFIGURATION (fuer Live-Kursabruf)
-------------------------------------------
1. IB TWS oder IB Gateway starten und einloggen.
2. API-Verbindung aktivieren:
       Datei → Globale Konfiguration → API → Einstellungen
       → "ActiveX und Socket Clients aktivieren" anhakenHaekchen setzen
3. Erlaubte IP-Adressen: 127.0.0.1 eintragen (falls nicht vorhanden)
4. TWS neu starten, damit die Aenderungen wirksam werden.

Ports:
    7496 = TWS Live-Trading     (Standard)
    7497 = TWS Paper-Trading
    4001 = IB Gateway Live
    4002 = IB Gateway Paper


PROGRAMM STARTEN
----------------
Virtuelle Umgebung aktivieren (falls noch nicht geschehen), dann:

    python optionsrechner.py


BEDIENUNG
---------
1. "Kurs von IB laden" klicken, um den aktuellen ALV-Kurs von der TWS abzurufen.
   (TWS muss dafuer gestartet und eingeloggt sein.)

2. Alternativ: Volatilitaet und Zinssatz manuell anpassen, dann
   "Tabelle neu berechnen" klicken.

3. Die Tabelle zeigt fuer jeden Eurex-Verfallstermin (3. Freitag des Monats)
   innerhalb der naechsten 60 Tage die theoretischen Put-Preise fuer:
       - ATM  (At the Money, gerundet auf naechsten ganzen Euro)
       - -1%  (1% unter dem aktuellen Kurs)
       - -3%  (3% unter dem aktuellen Kurs)
       - -5%  (5% unter dem aktuellen Kurs)

Eingabeparameter:
    Impl. Vola (%)  : Implizite Volatilitaet in Prozent (Standard: 25 %)
    Zinssatz (%)    : Risikofreier Zinssatz in Prozent (Standard: 2.5 %)


HINWEISE
--------
- Alle berechneten Preise sind theoretische Werte nach Black-Scholes-Merton.
- Preise gelten pro Aktie. Ein Eurex-Kontrakt umfasst i.d.R. 100 Aktien.
- Ohne TWS-Verbindung koennen Vola und Zinssatz manuell eingetragen und
  die Tabelle mit einem beliebigen Kurs berechnet werden (Kurs muss dann
  zuvor ueber "Kurs von IB laden" oder nach einem erfolgreichen Abruf
  gesetzt worden sein).
