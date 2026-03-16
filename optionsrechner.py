"""
IB Optionspreis-Rechner
=======================

Dieses Programm verbindet sich mit der Interactive Brokers Trader Workstation (TWS)
und berechnet theoretische Short-Put-Preise nach Black-Scholes-Merton.

Funktionen:
- Liest Watchlisten direkt aus der TWS-Konfigurationsdatei
- Auswahl der zu analysierenden Aktie per Dropdown
- Automatischer Abruf von Kurs und ATM-IV von IB
- Lädt echte Strikes und Optionspreise aus der IB-Optionskette (EUREX)
- Tabelle: BSM-Preis vs. IB-Marktpreis, wöchentliche Laufzeiten bis 60 Tage

Voraussetzungen
---------------
1. IB TWS oder IB Gateway muss gestartet sein, API aktivieren:
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

Ports:
    7496 = TWS Live-Trading
    7497 = TWS Paper-Trading
    4001 = IB Gateway Live
    4002 = IB Gateway Paper
"""

import asyncio
import glob
import logging
import math
import os
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from tkinter import ttk

# Python 3.14 erstellt keine Event Loop mehr automatisch – vor ib_insync-Import setzen.
_loop = asyncio.new_event_loop()
asyncio.set_event_loop(_loop)

from ib_insync import IB, Contract, Option, Stock  # noqa: E402


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

APP_VERSION = '1.10'

TWS_HOST  = '127.0.0.1'
TWS_PORT  = 7496
CLIENT_ID = 11

MARKET_DATA_WAIT   = 5    # Sekunden auf Kurs/IV warten
OPTION_DATA_WAIT   = 8    # Sekunden auf Optionspreise warten (Batch)
MAX_DTE            = 60   # Maximale Laufzeit in Tagen

DEFAULT_IV   = 25.0       # Standard-Volatilität in Prozent
DEFAULT_RATE =  2.5       # Standard-Zinssatz in Prozent

_TWS_XML_GLOB = os.path.join(os.path.expanduser('~'), 'Jts', '*', 'tws.xml')


# ---------------------------------------------------------------------------
# TWS-Watchlisten aus XML lesen
# ---------------------------------------------------------------------------

def find_tws_xml() -> str | None:
    """Sucht die aktuellste TWS-Konfigurationsdatei."""
    files = sorted(glob.glob(_TWS_XML_GLOB), key=os.path.getmtime, reverse=True)
    return files[0] if files else None


def load_watchlists_from_xml() -> dict[str, list[int]]:
    """Liest alle Watchlisten (Name → conid-Liste) aus tws.xml."""
    xml_path = find_tws_xml()
    if not xml_path:
        return {}
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError:
        return {}

    watchlists: dict[str, list[int]] = {}
    for wl in tree.getroot().iter('Watchlist'):
        qmc = wl.find('.//QuoteMatrixContent')
        if qmc is None:
            continue
        name = qmc.attrib.get('name', 'Unbekannt')
        conids: list[int] = []
        for entry in wl.iter('TickerEntry'):
            try:
                conid = int(entry.attrib.get('conid', '0'))
            except ValueError:
                continue
            exchange = entry.attrib.get('exchange', '')
            if conid > 0 and exchange not in ('IDEALPRO', '.', ''):
                conids.append(conid)
        if conids:
            watchlists[name] = conids
    return watchlists


# ---------------------------------------------------------------------------
# IB: conids → Aktiensymbole auflösen
# ---------------------------------------------------------------------------

def resolve_conids_to_stocks(
    host: str, port: int, client_id: int, conids: list[int]
) -> list[dict]:
    """Löst conids über die IB API zu Aktien-Symbolen auf (nur secType=STK)."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    ib = IB()
    stocks: list[dict] = []
    try:
        ib.connect(host, port, clientId=client_id)
        for conid in conids:
            try:
                details = ib.reqContractDetails(Contract(conId=conid))
                if details:
                    c = details[0].contract
                    if c.secType == 'STK':
                        stocks.append({
                            'symbol':   c.symbol,
                            'exchange': c.primaryExchange or c.exchange,
                            'currency': c.currency,
                            'conid':    conid,
                        })
            except Exception:
                pass
    finally:
        if ib.isConnected():
            ib.disconnect()
    return stocks


# ---------------------------------------------------------------------------
# IB: Marktpreise aus Ticker lesen
# ---------------------------------------------------------------------------

def get_price(ticker) -> float | None:
    """Bester verfügbarer Preis aus einem ib_insync-Ticker.

    Priorität: Last → Close → Bid/Ask-Mitte.
    """
    if ticker is None:
        return None
    last  = ticker.last
    close = ticker.close
    bid   = ticker.bid
    ask   = ticker.ask
    if last  and last  > 0:                return last
    if close and close > 0:                return close
    if bid and ask and bid > 0 and ask > 0: return (bid + ask) / 2
    return None


# ---------------------------------------------------------------------------
# IB: Vollständiger Datenabruf (Kurs + IV + Optionskette)
# ---------------------------------------------------------------------------

def fetch_full_data_from_ib(
    host: str, port: int, client_id: int,
    symbol: str, exchange: str, currency: str, conid: int,
    status_cb=None,
) -> dict:
    """Ruft Kurs, ATM-IV und die vollständige Optionskette von IB ab.

    Ablauf:
    1. Kurs (DataType 2 = Frozen)
    2. ATM-IV (OPTION_IMPLIED_VOLATILITY, 5-Tage-Verlauf)
    3. Optionsparameter (Strikes und Verfälle von EUREX)
    4. Marktpreise aller Put-Optionen (ATM bis -10%, wöchentlich bis 60 Tage)

    Args:
        host, port, client_id: IB-Verbindungsparameter.
        symbol, exchange, currency: Aktien-Kennzeichnung.
        conid:     Bekannte Contract-ID (beschleunigt qualifyContracts).
        status_cb: Optionaler Callback(str) für Statusmeldungen während des Ladevorgangs.

    Returns:
        Dict mit Schlüsseln:
            'price'  : float | None
            'iv'     : float | None  (in Prozent)
            'chain'  : list[dict]    – je Eintrag: expiry, strike, ib_price
    """

    def status(msg: str):
        if status_cb:
            status_cb(msg)

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    ib = IB()
    result: dict = {'price': None, 'iv': None, 'chain': []}

    try:
        ib.connect(host, port, clientId=client_id)

        # Aktien-Contract qualifizieren
        contract = Stock(symbol, exchange, currency)
        contract.conId = conid
        ib.qualifyContracts(contract)

        # --- 1. Kurs ---
        status(f'{symbol}: Lade Kurs …')
        ib.reqMarketDataType(2)   # Frozen: Live oder letzter bekannter Kurs
        ticker = ib.reqMktData(contract, '', snapshot=True, regulatorySnapshot=False)
        ib.sleep(MARKET_DATA_WAIT)
        result['price'] = get_price(ticker)

        # --- 2. ATM Implizite Volatilität ---
        status(f'{symbol}: Lade impl. Volatilität …')
        bars = ib.reqHistoricalData(
            contract,
            endDateTime='',
            durationStr='5 D',
            barSizeSetting='1 day',
            whatToShow='OPTION_IMPLIED_VOLATILITY',
            useRTH=True,
            formatDate=1,
        )
        if bars:
            result['iv'] = round(bars[-1].close * 100, 2)

        if result['price'] is None:
            status(f'{symbol}: Kein Kurs verfügbar.')
            return result

        spot = result['price']

        # --- 3. Optionsparameter (Strikes & Verfälle) ---
        status(f'{symbol}: Lade Optionskette …')
        chains = ib.reqSecDefOptParams(symbol, '', 'STK', contract.conId)

        # EUREX mit tradingClass=symbol bevorzugen (hat alle wöchentlichen Termine)
        opt_chain = next(
            (c for c in chains if c.exchange == 'EUREX' and c.tradingClass == symbol),
            None,
        )
        if opt_chain is None:
            opt_chain = next((c for c in chains if c.exchange == 'EUREX'), None)
        if opt_chain is None and chains:
            opt_chain = chains[0]
        if opt_chain is None:
            status(f'{symbol}: Keine Optionskette gefunden.')
            return result

        # Verfälle: alle verfügbaren Termine innerhalb 60 Tage (wöchentlich)
        today    = date.today()
        max_date = today + timedelta(days=MAX_DTE)
        valid_expiries: list[tuple[str, date]] = []
        for exp_str in sorted(opt_chain.expirations):
            try:
                exp_date = date(int(exp_str[:4]), int(exp_str[4:6]), int(exp_str[6:8]))
            except ValueError:
                continue
            if today < exp_date <= max_date:
                valid_expiries.append((exp_str, exp_date))

        # Strikes: ATM (gerundet) bis -10%
        min_strike = round(spot * 0.90)
        max_strike = round(spot * 1.005)
        valid_strikes = sorted(
            s for s in opt_chain.strikes if min_strike <= s <= max_strike
        )

        if not valid_expiries or not valid_strikes:
            status(f'{symbol}: Keine passenden Strikes/Verfälle gefunden.')
            return result

        n_total = len(valid_expiries) * len(valid_strikes)
        status(
            f'{symbol}: Lade {n_total} Optionspreise '
            f'({len(valid_expiries)} Verfälle × {len(valid_strikes)} Strikes) …'
        )

        # --- 4. Option-Contracts erzeugen und batch-qualifizieren ---
        opt_meta: list[tuple[Option, str, date, float]] = []
        for exp_str, exp_date in valid_expiries:
            for strike in valid_strikes:
                opt = Option(
                    symbol, exp_str, strike, 'P',
                    opt_chain.exchange,
                    currency=currency,
                    tradingClass=opt_chain.tradingClass,  # verhindert ambiguous contract
                )
                opt_meta.append((opt, exp_str, exp_date, strike))

        ib.qualifyContracts(*[o for o, *_ in opt_meta])

        # Marktdaten für alle qualifizierten Optionen anfordern (streaming für Greeks)
        ib.reqMarketDataType(2)
        opt_tickers: dict[tuple[str, float], tuple] = {}
        for opt, exp_str, exp_date, strike in opt_meta:
            if opt.conId:
                t = ib.reqMktData(opt, '', snapshot=False, regulatorySnapshot=False)
                opt_tickers[(exp_str, strike)] = (t, exp_date, opt)

        ib.sleep(OPTION_DATA_WAIT)

        # Ergebnisse einsammeln + Subscriptions abbestellen
        chain_rows: list[dict] = []
        for (exp_str, strike), (t, exp_date, opt) in opt_tickers.items():
            mg = t.modelGreeks
            iv_option = (
                round(mg.impliedVol * 100, 2)
                if mg and mg.impliedVol and mg.impliedVol > 0
                else None
            )
            chain_rows.append({
                'expiry':    exp_date,
                'strike':    strike,
                'ib_price':  get_price(t),
                'iv_option': iv_option,
            })
            try:
                ib.cancelMktData(opt)
            except Exception:
                pass

        # Sortierung: Verfallsdatum aufsteigend, Strike absteigend (ATM zuerst)
        result['chain'] = sorted(
            chain_rows, key=lambda x: (x['expiry'], -x['strike'])
        )

    finally:
        if ib.isConnected():
            ib.disconnect()

    return result


# ---------------------------------------------------------------------------
# Black-Scholes-Merton – Put-Preisformel
# ---------------------------------------------------------------------------

def _norm_cdf(x: float) -> float:
    """Kumulative Standardnormalverteilung N(x) (Abramowitz & Stegun 26.2.17)."""
    a = (0.319381530, -0.356563782, 1.781477937, -1.821255978, 1.330274429)
    t = 1.0 / (1.0 + 0.2316419 * abs(x))
    poly = t * (a[0] + t * (a[1] + t * (a[2] + t * (a[3] + t * a[4]))))
    pdf  = math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)
    cdf  = 1.0 - pdf * poly
    return cdf if x >= 0 else 1.0 - cdf


def bs_put_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """Theoretischer Put-Preis nach Black-Scholes-Merton.

    Args:
        S:     Aktienkurs (Spot).
        K:     Strike-Preis.
        T:     Laufzeit in Jahren.
        r:     Risikofreier Zinssatz, annualisiert (z.B. 0.025).
        sigma: Impl. Volatilität, annualisiert (z.B. 0.25).

    Returns:
        Theoretischer Put-Preis pro Aktie in EUR.
    """
    if T <= 0.0 or sigma <= 0.0 or S <= 0.0:
        return max(K - S, 0.0)
    d1 = (math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    return max(K * math.exp(-r * T) * _norm_cdf(-d2) - S * _norm_cdf(-d1), 0.0)


# ---------------------------------------------------------------------------
# Tkinter-Anwendung
# ---------------------------------------------------------------------------

class OptionsrechnerApp(tk.Tk):
    """Hauptfenster des IB Optionsrechners.

    Lädt Watchlisten aus TWS, ermöglicht Aktienauswahl und zeigt eine
    Short-Put-Tabelle mit BSM-Preis vs. IB-Marktpreis.
    """

    _COLUMNS   = ('Verfallstag', 'DTE', 'Strike', 'OTM %', 'Option IV %',
                  'BSM Preis (€)', 'IB Preis (€)', 'Differenz (€)')
    _COL_WIDTHS = (110, 45, 70, 65, 90, 100, 100, 95)

    def __init__(self):
        super().__init__()
        self.title(f'IB Optionsrechner  v{APP_VERSION}')
        self.resizable(True, True)

        self._watchlists:  dict[str, list[int]] = {}
        self._stocks:      list[dict]           = []
        self._chain_data:  list[dict]           = []
        self._fetch_thread: threading.Thread | None = None

        self._watchlist_var = tk.StringVar()
        self._stock_var     = tk.StringVar()
        self._spot_var      = tk.StringVar()
        self._iv_var        = tk.StringVar(value=str(DEFAULT_IV))
        self._rate_var      = tk.StringVar(value=str(DEFAULT_RATE))
        self._status_var    = tk.StringVar(value='Bereit')

        self._build_widgets()
        self._load_watchlists_from_xml()

    # ------------------------------------------------------------------
    # Widget-Aufbau
    # ------------------------------------------------------------------

    def _build_widgets(self):
        self._build_watchlist_frame()
        self._build_params_frame()
        self._build_table()
        self._build_status_bar()

    def _build_watchlist_frame(self):
        """Panel für Watchlist- und Aktienauswahl."""
        frame = ttk.LabelFrame(self, text='Watchlist & Aktienauswahl', padding=6)
        frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        ttk.Label(frame, text='Watchlist:').grid(row=0, column=0, sticky=tk.W)
        self._wl_combo = ttk.Combobox(
            frame, textvariable=self._watchlist_var, width=12, state='readonly')
        self._wl_combo.grid(row=0, column=1, padx=(4, 6))
        self._load_wl_btn = ttk.Button(
            frame, text='Symbole laden', command=self._on_load_watchlist)
        self._load_wl_btn.grid(row=0, column=2, padx=(0, 20))

        ttk.Label(frame, text='Aktie:').grid(row=0, column=3, sticky=tk.W)
        self._stock_combo = ttk.Combobox(
            frame, textvariable=self._stock_var, width=8, state='readonly')
        self._stock_combo.grid(row=0, column=4, padx=(4, 6))
        self._load_data_btn = ttk.Button(
            frame, text='Daten & Optionskette laden', command=self._on_load_data)
        self._load_data_btn.grid(row=0, column=5)

    def _build_params_frame(self):
        """Panel für Kurs, Vola, Zinssatz und Neuberechnung."""
        frame = ttk.LabelFrame(
            self, text='Parameter (automatisch befüllt, manuell anpassbar)', padding=6)
        frame.pack(fill=tk.X, padx=10, pady=4)

        ttk.Label(frame, text='Kurs (€):').grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self._spot_var, width=10).grid(
            row=0, column=1, padx=(4, 20))

        ttk.Label(frame, text='ATM IV (%):').grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(frame, textvariable=self._iv_var, width=7).grid(
            row=0, column=3, padx=(4, 20))

        ttk.Label(frame, text='Zinssatz (%):').grid(row=0, column=4, sticky=tk.W)
        ttk.Entry(frame, textvariable=self._rate_var, width=7).grid(
            row=0, column=5, padx=(4, 20))

        ttk.Button(frame, text='BSM neu berechnen',
                   command=self._update_table).grid(row=0, column=6)

    def _build_table(self):
        """Treeview-Tabelle für BSM- vs. IB-Optionspreise."""
        frame = ttk.LabelFrame(
            self,
            text='Short-Put Optionspreise: BSM-Modell vs. IB-Marktpreis',
            padding=6,
        )
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self._tree = ttk.Treeview(
            frame, columns=self._COLUMNS, show='headings', height=14)

        for col, width in zip(self._COLUMNS, self._COL_WIDTHS):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=width, anchor=tk.CENTER, stretch=False)

        # Abwechselnde Zeilenfarben nach Verfallsgruppe
        self._tree.tag_configure('grp_a', background='#f0f4ff')
        self._tree.tag_configure('grp_b', background='#ffffff')

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_status_bar(self):
        """Untere Statusleiste."""
        frame = ttk.Frame(self)
        frame.pack(fill=tk.X, padx=10, pady=(0, 6))
        self._status_label = ttk.Label(
            frame, textvariable=self._status_var, foreground='gray')
        self._status_label.pack(side=tk.LEFT)
        ttk.Label(
            frame,
            text='Formel: Black-Scholes-Merton  |  Strikes & Preise: EUREX via IB',
            foreground='gray',
        ).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Watchlisten aus XML laden (offline)
    # ------------------------------------------------------------------

    def _load_watchlists_from_xml(self):
        self._watchlists = load_watchlists_from_xml()
        names = list(self._watchlists.keys())
        self._wl_combo['values'] = names
        if names:
            self._watchlist_var.set(names[0])
            self._set_status(f'Watchlisten aus TWS geladen: {", ".join(names)}')
        else:
            self._set_status('Keine TWS-Watchlisten gefunden.', error=True)

    # ------------------------------------------------------------------
    # Symbole der Watchlist per IB auflösen
    # ------------------------------------------------------------------

    def _on_load_watchlist(self):
        if self._fetch_thread and self._fetch_thread.is_alive():
            return
        wl_name = self._watchlist_var.get()
        if not wl_name or wl_name not in self._watchlists:
            self._set_status('Keine Watchlist ausgewählt.', error=True)
            return
        self._load_wl_btn.configure(state=tk.DISABLED)
        self._set_status('Verbinde mit IB TWS und löse Symbole auf …')
        conids = self._watchlists[wl_name]
        self._fetch_thread = threading.Thread(
            target=self._run_resolve_conids, args=(conids,), daemon=True)
        self._fetch_thread.start()

    def _run_resolve_conids(self, conids: list[int]):
        try:
            stocks = resolve_conids_to_stocks(TWS_HOST, TWS_PORT, CLIENT_ID, conids)
            self.after(0, lambda s=stocks: self._on_stocks_resolved(s))
        except ConnectionRefusedError:
            self.after(0, lambda: self._set_status(
                f'Verbindung zu {TWS_HOST}:{TWS_PORT} fehlgeschlagen.', error=True))
        except Exception as exc:
            self.after(0, lambda e=exc: self._set_status(
                f'Fehler: {e}', error=True))
        finally:
            self.after(0, lambda: self._load_wl_btn.configure(state=tk.NORMAL))

    def _on_stocks_resolved(self, stocks: list[dict]):
        self._stocks = stocks
        symbols = [s['symbol'] for s in stocks]
        self._stock_combo['values'] = symbols
        if symbols:
            self._stock_var.set(symbols[0])
            self._set_status(f'{len(symbols)} Aktien geladen: {", ".join(symbols)}')
        else:
            self._set_status('Keine Aktien in dieser Watchlist.', error=True)

    # ------------------------------------------------------------------
    # Marktdaten + Optionskette laden
    # ------------------------------------------------------------------

    def _on_load_data(self):
        if self._fetch_thread and self._fetch_thread.is_alive():
            return
        symbol = self._stock_var.get()
        if not symbol:
            self._set_status('Keine Aktie ausgewählt.', error=True)
            return
        stock = next((s for s in self._stocks if s['symbol'] == symbol), None)
        if stock is None:
            self._set_status('Aktie nicht gefunden – Watchlist neu laden.', error=True)
            return
        self._load_data_btn.configure(state=tk.DISABLED)
        self._set_status(f'Lade Daten für {symbol} …')
        self._fetch_thread = threading.Thread(
            target=self._run_fetch_data, args=(stock,), daemon=True)
        self._fetch_thread.start()

    def _run_fetch_data(self, stock: dict):
        def on_status(msg: str):
            self.after(0, lambda m=msg: self._set_status(m))
        try:
            data = fetch_full_data_from_ib(
                TWS_HOST, TWS_PORT, CLIENT_ID,
                stock['symbol'], stock['exchange'], stock['currency'], stock['conid'],
                status_cb=on_status,
            )
            self.after(0, lambda d=data, s=stock: self._on_data_received(d, s))
        except ConnectionRefusedError:
            self.after(0, lambda: self._set_status(
                f'Verbindung zu {TWS_HOST}:{TWS_PORT} fehlgeschlagen.', error=True))
        except Exception as exc:
            self.after(0, lambda e=exc: self._set_status(
                f'Fehler beim Datenabruf: {e}', error=True))
        finally:
            self.after(0, lambda: self._load_data_btn.configure(state=tk.NORMAL))

    def _on_data_received(self, data: dict, stock: dict):
        symbol = stock['symbol']
        msgs: list[str] = []

        if data['price'] is not None:
            self._spot_var.set(f"{data['price']:.2f}")
            msgs.append(f"Kurs {data['price']:.2f} €")
        else:
            self._set_status(f'{symbol}: Kein Kurs – bitte manuell eintragen.', error=True)

        if data['iv'] is not None:
            self._iv_var.set(f"{data['iv']:.2f}")
            msgs.append(f"ATM IV {data['iv']:.2f} %")
        else:
            msgs.append('IV nicht verfügbar')

        self._chain_data = data['chain']
        n_chain = len(self._chain_data)
        msgs.append(f'{n_chain} Optionen geladen')

        self._set_status(f'{symbol}: {" | ".join(msgs)}')
        self._update_table()

    # ------------------------------------------------------------------
    # Tabelle berechnen
    # ------------------------------------------------------------------

    def _update_table(self):
        """Berechnet BSM-Preise und vergleicht mit IB-Marktpreisen."""
        for row in self._tree.get_children():
            self._tree.delete(row)

        try:
            spot = float(self._spot_var.get())
        except ValueError:
            self._tree.insert('', tk.END, values=(
                'Kurs fehlt – bitte laden oder manuell eintragen', *[''] * 7))
            return

        try:
            iv = float(self._iv_var.get()) / 100.0
            r  = float(self._rate_var.get()) / 100.0
        except ValueError:
            self._set_status('Ungültige Eingabe – IV und Zinssatz müssen Zahlen sein.',
                             error=True)
            return

        if not self._chain_data:
            self._tree.insert('', tk.END, values=(
                'Keine Daten – Aktie und Optionskette laden', *[''] * 7))
            return

        prev_expiry = None
        grp = 0  # Gruppenindex für abwechselnde Farben

        for row_data in self._chain_data:
            exp_date  = row_data['expiry']
            strike    = row_data['strike']
            ib_price  = row_data['ib_price']
            iv_option = row_data.get('iv_option')   # per-Option IV in Prozent

            dte = max((exp_date - date.today()).days, 0)
            T   = dte / 365.0
            otm_pct = (strike / spot - 1.0) * 100.0

            # BSM mit per-Option IV (Skew-bereinigt); Fallback auf ATM IV
            iv_bsm = (iv_option / 100.0) if iv_option is not None else iv
            bsm    = bs_put_price(spot, strike, T, r, iv_bsm)

            iv_str   = f'{iv_option:.1f}'       if iv_option is not None else '–'
            ib_str   = f'{ib_price:.2f}'        if ib_price  is not None else '–'
            diff_str = f'{bsm - ib_price:+.2f}' if ib_price  is not None else '–'

            if exp_date != prev_expiry:
                grp = 1 - grp
                prev_expiry = exp_date

            tag = 'grp_a' if grp == 0 else 'grp_b'
            self._tree.insert('', tk.END, values=(
                exp_date.strftime('%d.%m.%Y'),
                str(dte),
                f'{strike:.0f}',
                f'{otm_pct:.1f} %',
                iv_str,
                f'{bsm:.2f}',
                ib_str,
                diff_str,
            ), tags=(tag,))

        symbol = self._stock_var.get() or '–'
        self._set_status(
            f'{symbol}  {spot:.2f} €  |  ATM IV {iv*100:.2f} %  |  '
            f'r {r*100:.2f} %  |  {len(self._chain_data)} Zeilen')

    # ------------------------------------------------------------------
    # Hilfsmethoden
    # ------------------------------------------------------------------

    def _set_status(self, message: str, error: bool = False):
        """Aktualisiert die Statusleiste (rot bei Fehler, grau sonst)."""
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
