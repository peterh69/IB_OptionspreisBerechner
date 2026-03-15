"""
IB Optionspreis-Rechner – ALV Short Put
========================================

Dieses Programm verbindet sich mit der Interactive Brokers Trader Workstation (TWS)
oder dem IB Gateway und berechnet theoretische Short-Put-Preise für die Allianz-Aktie
(ALV, XETRA) auf Basis des Black-Scholes-Merton-Modells.

Angezeigt wird eine Tabelle mit:
- Verfallstagen bis zu 60 Tage (monatliche Eurex-Verfallstermine = 3. Freitag)
- Vier Strike-Stufen pro Verfall: ATM, -1% OTM, -3% OTM, -5% OTM

Der aktuelle ALV-Kurs wird von der IB TWS abgerufen. Implizite Volatilität und
Zinssatz sind als editierbare Felder in der GUI einstellbar.

Voraussetzungen
---------------
1. IB TWS oder IB Gateway muss gestartet und die API-Verbindung aktiviert sein:
   File → Global Configuration → API → Settings → „Enable ActiveX and Socket Clients"
2. Python-Pakete: ib_insync (siehe requirements.txt)

Installation
------------
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt

Verwendung
----------
    python optionsrechner.py

Konfiguration
-------------
Die Verbindungsparameter stehen im Abschnitt „Konfiguration" unten.

Ports:
    7496 = TWS Live-Trading
    7497 = TWS Paper-Trading
    4001 = IB Gateway Live
    4002 = IB Gateway Paper
"""

import asyncio
import logging
import math
import threading
import tkinter as tk
from datetime import date, timedelta
from tkinter import ttk

# Python 3.10+ erstellt keine Event Loop mehr automatisch.
# Vor dem Import von ib_insync muss eine neue Loop angelegt werden.
asyncio.set_event_loop(asyncio.new_event_loop())

from ib_insync import IB, Stock  # noqa: E402  (Import nach asyncio-Setup)


# ---------------------------------------------------------------------------
# Log-Filter: Harmlose IB-Fehlercodes unterdrücken
# ---------------------------------------------------------------------------

_BENIGN_LOG_CODES = {321, 354, 2104, 2106, 2107, 2108, 2158, 10090, 10167}


class _IBLogFilter(logging.Filter):
    """Unterdrückt bekannte, harmlose ib_insync-Fehlermeldungen im Log."""

    def filter(self, record):
        msg = record.getMessage()
        return not any(
            f'Error {c},' in msg or f'{c},' in msg[:20]
            for c in _BENIGN_LOG_CODES
        )


for _log_name in ('ib_insync.wrapper', 'ib_insync.client', 'ib_insync.ib'):
    logging.getLogger(_log_name).addFilter(_IBLogFilter())


# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

APP_VERSION = '1.00'       # Wird bei jeder Code-Änderung um 0.01 erhöht

TWS_HOST = '127.0.0.1'    # Hostname oder IP der TWS/Gateway-Instanz
TWS_PORT = 7496            # 7496=TWS Live, 7497=TWS Paper, 4001=Gateway Live
CLIENT_ID = 11             # Eindeutige Client-ID (IB_Ausstiegsrechner belegt 10)

MARKET_DATA_WAIT = 3       # Sekunden auf eingehende Marktdaten warten
MAX_DTE = 60               # Maximale Laufzeit in Tagen

DEFAULT_IV   = 25.0        # Standard-Volatilität in Prozent
DEFAULT_RATE = 2.5         # Standard-Zinssatz in Prozent (EZB-Leitzins)


# ---------------------------------------------------------------------------
# Black-Scholes-Merton – Put-Preisformel
# ---------------------------------------------------------------------------

def _norm_cdf(x: float) -> float:
    """Berechnet die kumulative Standardnormalverteilung N(x).

    Verwendet die Näherungsformel von Abramowitz & Stegun (26.2.17),
    die eine Genauigkeit von 7 Dezimalstellen liefert.

    Args:
        x: Argument der Verteilungsfunktion.

    Returns:
        Wert N(x) im Intervall [0, 1].
    """
    a = (0.319381530, -0.356563782, 1.781477937, -1.821255978, 1.330274429)
    t = 1.0 / (1.0 + 0.2316419 * abs(x))
    poly = t * (a[0] + t * (a[1] + t * (a[2] + t * (a[3] + t * a[4]))))
    pdf = math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)
    cdf = 1.0 - pdf * poly
    return cdf if x >= 0 else 1.0 - cdf


def bs_put_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """Berechnet den theoretischen Preis eines europäischen Puts (Black-Scholes-Merton).

    Formel: P = K·e^(-rT)·N(-d2) - S·N(-d1)
    mit d1 = [ln(S/K) + (r + σ²/2)·T] / (σ·√T)
         d2 = d1 - σ·√T

    Args:
        S:     Aktueller Aktienkurs (Spot).
        K:     Strike-Preis der Option.
        T:     Laufzeit in Jahren (z.B. 33/365 für 33 Tage).
        r:     Risikofreier Zinssatz, annualisiert (z.B. 0.025 für 2,5 %).
        sigma: Implizite Volatilität, annualisiert (z.B. 0.25 für 25 %).

    Returns:
        Theoretischer Put-Preis pro Aktie in EUR.
        Bei abgelaufener Option wird der innere Wert max(K-S, 0) zurückgegeben.
    """
    if T <= 0.0:
        return max(K - S, 0.0)
    if sigma <= 0.0 or S <= 0.0:
        return max(K - S, 0.0)
    d1 = (math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    return max(K * math.exp(-r * T) * _norm_cdf(-d2) - S * _norm_cdf(-d1), 0.0)


# ---------------------------------------------------------------------------
# Datumsberechnung – Eurex-Verfallstermine
# ---------------------------------------------------------------------------

def get_expiry_dates(max_days: int = MAX_DTE) -> list[date]:
    """Gibt monatliche Eurex-Verfallstermine (3. Freitag) bis max_days zurück.

    ALV-Optionen werden an der Eurex gehandelt. Der monatliche Standard-
    Verfallstermin ist jeweils der 3. Freitag des Monats (Settlement-Tag).

    Args:
        max_days: Maximale Laufzeit in Tagen ab heute.

    Returns:
        Sortierte Liste von Verfallsdaten im Fenster (heute, heute+max_days].
    """
    today = date.today()
    max_date = today + timedelta(days=max_days)
    expiries: list[date] = []

    # Drei Monate prüfen (genug für ein 60-Tage-Fenster)
    for month_offset in range(3):
        month = today.month + month_offset
        year = today.year + (month - 1) // 12
        month = ((month - 1) % 12) + 1

        first_of_month = date(year, month, 1)
        # Erster Freitag des Monats (weekday 4 = Freitag)
        days_to_friday = (4 - first_of_month.weekday()) % 7
        first_friday = first_of_month + timedelta(days=days_to_friday)
        third_friday = first_friday + timedelta(weeks=2)

        if today < third_friday <= max_date:
            expiries.append(third_friday)

    return sorted(expiries)


def days_to_expiry(expiry: date) -> int:
    """Berechnet die verbleibenden Kalendertage bis zum Verfallsdatum.

    Args:
        expiry: Verfallsdatum.

    Returns:
        Anzahl der verbleibenden Tage (≥ 0).
    """
    return max((expiry - date.today()).days, 0)


# ---------------------------------------------------------------------------
# IB-Marktdaten – ALV Kursabfrage
# ---------------------------------------------------------------------------

def get_price(ticker) -> float | None:
    """Ermittelt den besten verfügbaren Marktpreis aus einem ib_insync-Ticker.

    Prioritätsreihenfolge: Last Price → Close → Mitte aus Bid/Ask.
    Gibt None zurück, wenn kein Preis verfügbar ist.

    Args:
        ticker: ib_insync Ticker-Objekt.

    Returns:
        Preis als float, oder None wenn nicht verfügbar.
    """
    if ticker is None:
        return None
    last  = ticker.last
    close = ticker.close
    bid   = ticker.bid
    ask   = ticker.ask

    if last and last > 0:
        return last
    if close and close > 0:
        return close
    if bid and ask and bid > 0 and ask > 0:
        return (bid + ask) / 2
    return None


def fetch_alv_price_from_ib(host: str, port: int, client_id: int) -> float | None:
    """Verbindet sich mit IB TWS und liest den aktuellen ALV-Kurs.

    Läuft in einem separaten Thread. Nach dem Abruf wird die Verbindung
    sofort wieder getrennt.

    Args:
        host:      TWS-Hostname (z.B. '127.0.0.1').
        port:      API-Port (z.B. 7496).
        client_id: Client-ID für die IB-Verbindung.

    Returns:
        ALV-Kurs als float, oder None bei Fehler.

    Raises:
        ConnectionRefusedError: Wenn TWS nicht erreichbar ist.
        Exception:              Bei anderen IB-API-Fehlern.
    """
    # Jeder Thread braucht eine eigene asyncio Event-Loop (ib_insync-Anforderung)
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    ib = IB()
    try:
        ib.connect(host, port, clientId=client_id)

        # ALV an der XETRA (IBIS = IB-interner Exchange-Code für XETRA)
        contract = Stock('ALV', 'IBIS', 'EUR')
        ib.qualifyContracts(contract)

        # Marktdaten anfordern und auf Daten warten
        ticker = ib.reqMktData(contract, '', snapshot=False, regulatorySnapshot=False)
        ib.sleep(MARKET_DATA_WAIT)

        return get_price(ticker)

    finally:
        if ib.isConnected():
            ib.disconnect()


# ---------------------------------------------------------------------------
# Strike-Berechnung
# ---------------------------------------------------------------------------

def compute_strikes(spot: float) -> dict[str, float]:
    """Berechnet die vier Strike-Stufen für einen gegebenen Aktienkurs.

    Eurex-Optionen auf ALV haben 1-EUR-Strike-Abstände; daher wird auf
    den nächsten ganzen Euro gerundet.

    Strike-Stufen:
        ATM: Kurs gerundet auf 1 EUR
        -1%: 1 % unter dem Kurs
        -3%: 3 % unter dem Kurs
        -5%: 5 % unter dem Kurs

    Args:
        spot: Aktueller Aktienkurs in EUR.

    Returns:
        Dict mit Schlüsseln 'ATM', '-1%', '-3%', '-5%' und Strike-Werten.
    """
    return {
        'ATM': round(spot),
        '-1%': round(spot * 0.99),
        '-3%': round(spot * 0.97),
        '-5%': round(spot * 0.95),
    }


# ---------------------------------------------------------------------------
# Tkinter-Anwendung
# ---------------------------------------------------------------------------

class OptionsrechnerApp(tk.Tk):
    """Hauptfenster des ALV Short-Put Optionsrechners.

    Zeigt eine Tabelle mit theoretischen Short-Put-Preisen für verschiedene
    Laufzeiten (bis 60 Tage) und Strike-Stufen (ATM, -1%, -3%, -5% OTM).
    Der aktuelle ALV-Kurs wird optional von der IB TWS abgerufen.
    """

    # Spalten der Tabelle
    _COLUMNS = (
        'Verfallstag', 'DTE',
        'ATM Strike', 'ATM Preis',
        '-1% Strike', '-1% Preis',
        '-3% Strike', '-3% Preis',
        '-5% Strike', '-5% Preis',
    )

    # Spaltenbreiten in Pixeln
    _COL_WIDTHS = (110, 45, 90, 90, 90, 90, 90, 90, 90, 90)

    def __init__(self):
        super().__init__()
        self.title(f'ALV Short-Put Optionsrechner  v{APP_VERSION}')
        self.resizable(True, True)

        # Zustandsvariablen
        self._spot_price: float | None = None
        self._fetch_thread: threading.Thread | None = None

        # Tkinter-Variablen (an UI-Widgets gebunden)
        self._iv_var   = tk.StringVar(value=str(DEFAULT_IV))
        self._rate_var = tk.StringVar(value=str(DEFAULT_RATE))
        self._status_var = tk.StringVar(value='Bereit')

        self._build_widgets()
        self._update_table()   # Tabelle mit Platzhalter befüllen

    # ------------------------------------------------------------------
    # Widget-Aufbau
    # ------------------------------------------------------------------

    def _build_widgets(self):
        """Erstellt alle UI-Elemente des Hauptfensters."""
        self._build_controls()
        self._build_table()
        self._build_status_bar()

    def _build_controls(self):
        """Erstellt das obere Steuerleisten-Panel."""
        frame = ttk.LabelFrame(self, text='Steuerung', padding=6)
        frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        # Kurs-Anzeige
        ttk.Label(frame, text='ALV Kurs (€):').grid(row=0, column=0, sticky=tk.W)
        self._price_label = ttk.Label(frame, text='–', width=10,
                                      font=('TkDefaultFont', 10, 'bold'))
        self._price_label.grid(row=0, column=1, padx=(4, 20), sticky=tk.W)

        # Impl. Volatilität
        ttk.Label(frame, text='Impl. Vola (%):').grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(frame, textvariable=self._iv_var, width=7).grid(
            row=0, column=3, padx=(4, 20))

        # Zinssatz
        ttk.Label(frame, text='Zinssatz (%):').grid(row=0, column=4, sticky=tk.W)
        ttk.Entry(frame, textvariable=self._rate_var, width=7).grid(
            row=0, column=5, padx=(4, 20))

        # Schaltflächen
        self._fetch_btn = ttk.Button(
            frame, text='Kurs von IB laden', command=self._on_fetch_click)
        self._fetch_btn.grid(row=0, column=6, padx=(0, 6))

        ttk.Button(frame, text='Tabelle neu berechnen',
                   command=self._update_table).grid(row=0, column=7)

        # Hinweis auf Parameter
        ttk.Label(frame,
                  text='Volatilität und Zinssatz manuell anpassbar; Kurs wahlweise von IB abrufbar.',
                  foreground='gray').grid(row=1, column=0, columnspan=8,
                                          sticky=tk.W, pady=(4, 0))

    def _build_table(self):
        """Erstellt das Treeview-Widget für die Optionspreistabelle."""
        frame = ttk.LabelFrame(
            self,
            text='Short-Put Optionspreise – Black-Scholes-Merton (Preis pro Aktie in €)',
            padding=6,
        )
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self._tree = ttk.Treeview(
            frame, columns=self._COLUMNS, show='headings', height=12)

        for col, width in zip(self._COLUMNS, self._COL_WIDTHS):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=width, anchor=tk.CENTER, stretch=False)

        # Vertikaler Scrollbalken
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)

        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_status_bar(self):
        """Erstellt die untere Statusleiste."""
        frame = ttk.Frame(self)
        frame.pack(fill=tk.X, padx=10, pady=(0, 6))

        self._status_label = ttk.Label(
            frame, textvariable=self._status_var, foreground='gray')
        self._status_label.pack(side=tk.LEFT)

        ttk.Label(
            frame,
            text=(
                'Formel: Black-Scholes-Merton  |  '
                'ATM = aktueller Kurs (gerundet)  |  '
                '-1%/-3%/-5% = OTM Strikes'
            ),
            foreground='gray',
        ).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Ereignishandler
    # ------------------------------------------------------------------

    def _on_fetch_click(self):
        """Startet den IB-Kursabruf in einem Hintergrundthread."""
        if self._fetch_thread and self._fetch_thread.is_alive():
            return  # Bereits laufender Abruf – ignorieren

        self._fetch_btn.configure(state=tk.DISABLED)
        self._set_status('Verbinde mit IB TWS …')

        self._fetch_thread = threading.Thread(
            target=self._run_fetch, daemon=True)
        self._fetch_thread.start()

    def _run_fetch(self):
        """Hintergrundthread: Ruft den ALV-Kurs von IB TWS ab.

        Verwendet root.after() für den thread-sicheren Rückruf in den
        Tkinter-Hauptthread.
        """
        try:
            price = fetch_alv_price_from_ib(TWS_HOST, TWS_PORT, CLIENT_ID)
            if price is not None:
                # Thread-sicherer Aufruf zurück in den Hauptthread
                self.after(0, lambda p=price: self._on_price_received(p))
            else:
                self.after(0, lambda: self._set_status(
                    'Kein Kurs verfügbar – Kurs manuell eintragen oder '
                    'Marktdaten-Abonnement prüfen.',
                    error=True,
                ))
        except ConnectionRefusedError:
            self.after(0, lambda: self._set_status(
                f'Verbindung zu {TWS_HOST}:{TWS_PORT} fehlgeschlagen – '
                'TWS oder IB Gateway gestartet?',
                error=True,
            ))
        except Exception as exc:
            self.after(0, lambda e=exc: self._set_status(
                f'Fehler beim Abrufen: {e}', error=True))
        finally:
            self.after(0, lambda: self._fetch_btn.configure(state=tk.NORMAL))

    def _on_price_received(self, price: float):
        """Wird im Hauptthread aufgerufen, wenn ein neuer Kurs vorliegt."""
        self._spot_price = price
        self._price_label.configure(text=f'{price:.2f}')
        self._set_status(f'ALV-Kurs erfolgreich abgerufen: {price:.2f} €')
        self._update_table()

    # ------------------------------------------------------------------
    # Tabellenaktualisierung
    # ------------------------------------------------------------------

    def _update_table(self):
        """Berechnet und befüllt die Optionspreistabelle neu.

        Liest Volatilität und Zinssatz aus den Eingabefeldern.
        Falls kein Kurs vorliegt, wird eine Hinweiszeile angezeigt.
        """
        # Alte Einträge löschen
        for row in self._tree.get_children():
            self._tree.delete(row)

        if self._spot_price is None:
            self._tree.insert('', tk.END, values=(
                'Kurs noch nicht geladen', '', '', '', '', '', '', '', '', ''))
            self._set_status(
                'Kurs noch nicht geladen. „Kurs von IB laden" klicken '
                'oder TWS-Verbindung prüfen.')
            return

        # Parameter aus GUI lesen
        try:
            iv = float(self._iv_var.get()) / 100.0
            r  = float(self._rate_var.get()) / 100.0
        except ValueError:
            self._set_status('Ungültige Eingabe – Vola und Zinssatz müssen Zahlen sein.',
                             error=True)
            return

        S = self._spot_price
        strikes = compute_strikes(S)
        expiries = get_expiry_dates(MAX_DTE)

        if not expiries:
            self._tree.insert('', tk.END, values=(
                'Keine Verfallstermine im 60-Tage-Fenster', '', '', '',
                '', '', '', '', '', ''))
            return

        # Zeilen befüllen
        for exp in expiries:
            dte_days = days_to_expiry(exp)
            T = dte_days / 365.0

            row: list[str] = [exp.strftime('%d.%m.%Y'), str(dte_days)]
            for label in ('ATM', '-1%', '-3%', '-5%'):
                K = strikes[label]
                price = bs_put_price(S, K, T, r, iv)
                row.extend([f'{K:.0f}', f'{price:.2f}'])

            self._tree.insert('', tk.END, values=row)

        self._set_status(
            f'ALV {S:.2f} €  |  IV {iv*100:.1f} %  |  r {r*100:.2f} %  |  '
            f'{len(expiries)} Verfallstermine angezeigt')

    # ------------------------------------------------------------------
    # Hilfsmethoden
    # ------------------------------------------------------------------

    def _set_status(self, message: str, error: bool = False):
        """Aktualisiert die Statuszeile.

        Args:
            message: Anzuzeigende Nachricht.
            error:   True = rote Schrift, False = graue Schrift.
        """
        self._status_var.set(message)
        self._status_label.configure(foreground='red' if error else 'gray')


# ---------------------------------------------------------------------------
# Einstiegspunkt
# ---------------------------------------------------------------------------

def main():
    """Startet die Tkinter-Anwendung."""
    app = OptionsrechnerApp()
    app.mainloop()


if __name__ == '__main__':
    main()
