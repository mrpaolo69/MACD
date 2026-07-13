import os
import io
import json
import pandas as pd
import yfinance as yf
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ── SETTINGS ──────────────────────────────────────────────────────────────────
LOOKBACK   = "5y"
RSI_PERIOD = 14
RSI_MA     = 14

SECTORS = {
    "DJI":  "Dow Jones Index",
    "SPY":  "S&P 500 Index",
    "QQQ":  "Nasdaq 100 Index",
    "IWM":  "Russell 2000 Index",
    "XLK":  "Technology",
    "XLF":  "Financials",
    "XLV":  "Health Care",
    "XLY":  "Consumer Discretionary",
    "XLC":  "Communication Services",
    "XLI":  "Industrials",
    "XLP":  "Consumer Staples",
    "XLE":  "Energy",
    "XLU":  "Utilities",
    "XLRE": "Real Estate",
    "XLB":  "Materials",
}

# Seed tickers — only used on very first run before spreadsheet exists
WATCHLIST_SEED = [
    "AMD","VRT","PLTR","FUTU","SHOP","DELL","CRDO","ANET","HOOD","WDC",
    "CCJ","KTOS","FTNT","INOD","CSCO","IBIT","META","APP","MSFT","TSLA",
    "AXP","AVGO","GE","JPM","CLS","TSM","AAPL","GOOGL","STX","AMZN",
    "MU","NVDA","ETHA","SOFI","CDE","IREN","AA","ADI","CCL","HL",
    "AMAT","LRCX","APH","EQT","NEM","CAT","FCX","RTX","GLW","COHR",
    "DRAM","INTC","SMH","CEG","NBIS","TER"
]

FINANCIALS_SEED = [
    "JPM","BAC","WFC","GS","MS","C","AXP","BLK","SCHW","COF"
]

# ── Analysis helpers ──────────────────────────────────────────────────────────
def calc_ema(series, span):
    return series.ewm(span=span, adjust=False).mean()

def calc_rsi(close, period=14):
    delta    = close.diff()
    gain     = delta.clip(lower=0)
    loss     = -delta.clip(upper=0)
    avg_gain = gain.ewm(com=period - 1, adjust=False).mean()
    avg_loss = loss.ewm(com=period - 1, adjust=False).mean()
    rs       = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))

def get_last_crossover(series_a, series_b, index):
    last_date = None
    last_type = None
    for i in range(1, len(series_a)):
        prev = float(series_a.iloc[i-1]) - float(series_b.iloc[i-1])
        curr = float(series_a.iloc[i])   - float(series_b.iloc[i])
        if prev <= 0 and curr > 0:
            last_date = index[i].strftime("%Y-%m-%d")
            last_type = "▲ Bullish"
        elif prev >= 0 and curr < 0:
            last_date = index[i].strftime("%Y-%m-%d")
            last_type = "▼ Bearish"
    return last_date or "N/A", last_type or "N/A"

def analyze_ticker(ticker):
    try:
        t = "^DJI" if ticker == "DJI" else ticker
        df = yf.download(t, period=LOOKBACK, interval="1wk",
                         auto_adjust=True, progress=False)
        if df.empty or len(df) < 40:
            return None
        close = df["Close"].squeeze()

        macd   = calc_ema(close, 12) - calc_ema(close, 26)
        signal = calc_ema(macd, 9)
        macd_date, macd_type = get_last_crossover(macd, signal, macd.index)

        rsi    = calc_rsi(close, RSI_PERIOD)
        rsi_ma = rsi.rolling(window=RSI_MA).mean()
        rsi_date, rsi_type = get_last_crossover(rsi, rsi_ma, rsi.index)

        ma200     = close.rolling(window=200).mean()
        current   = float(close.iloc[-1])
        ma200_val = float(ma200.iloc[-1]) if not pd.isna(ma200.iloc[-1]) else None
        if ma200_val is None:
            above_200 = "N/A"
        elif current > ma200_val:
            above_200 = "✅ Above"
        else:
            above_200 = "❌ Below"

        return {
            "Ticker":          ticker,
            "MACD Cross Date": macd_date,
            "MACD Signal":     macd_type,
            "RSI Cross Date":  rsi_date,
            "RSI Signal":      rsi_type,
            "200W MA":         above_200,
        }
    except Exception as e:
        print(f"  Error on {ticker}: {e}")
        return None

def run_tickers(tickers, label):
    results = []
    for i, ticker in enumerate(tickers):
        row = analyze_ticker(ticker)
        results.append(row if row else {
            "Ticker":          ticker,
            "MACD Cross Date": "N/A", "MACD Signal": "N/A",
            "RSI Cross Date":  "N/A", "RSI Signal":  "N/A",
            "200W MA":         "N/A",
        })
        if (i + 1) % 10 == 0 or (i + 1) == len(tickers):
            print(f"  [{label}] Processed {i+1}/{len(tickers)}...")
    return results

def get_spmo_holdings():
    print("Fetching SPMO holdings...")
    try:
        tables = pd.read_html("https://stockanalysis.com/etf/spmo/holdings/")
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            if any("ticker" in c or "symbol" in c for c in cols):
                ticker_col = [c for c in t.columns if "ticker" in str(c).lower() or "symbol" in str(c).lower()][0]
                tickers = [str(x).strip().upper() for x in t[ticker_col].dropna() if str(x).strip() not in ["", "nan"]]
                if tickers:
                    print(f"  Found {len(tickers)} SPMO holdings")
                    return tickers
    except:
        pass
    print("  Using hardcoded SPMO fallback")
    return [
        "MU","NVDA","AVGO","GOOGL","AMD","JNJ","LRCX","GOOG","CAT","INTC",
        "AMAT","GE","RTX","TXN","HON","DE","UNH","LLY","PLD","AMT",
        "CCI","EQIX","PSA","SPG","DLR","O","VICI","WM","RSG","FAST",
        "ODFL","CHRW","XPO","JBHT","SAIA","MNST","KDP","KO","PEP","PM",
        "MO","BTI","MDLZ","HSY","GIS","CPB","SJM","CAG","HRL","MKC",
        "CVX","XOM","COP","EOG","PXD","DVN","MPC","VLO","PSX","HES",
        "SLB","HAL","BKR","FTI","NOV","OXY","APA","MRO","AR","EQT",
        "NEE","DUK","SO","AEP","EXC","SRE","PEG","ED","FE","PPL",
        "AWK","WEC","CMS","NI","LNT","EVRG","OGE","PNW","NWE","AVA",
        "AME","ROP","ITW","EMR","ETN","PH","IR","OTIS","CARR","TT"
    ]

# ── Read editable tickers from existing spreadsheet ───────────────────────────
def read_tickers_from_sheet(wb_bytes, sheet_name, seed_tickers, data_start_row=6):
    """Read tickers from col A of a named sheet in an in-memory workbook bytes."""
    try:
        wb = load_workbook(io.BytesIO(wb_bytes), read_only=True)
        if sheet_name not in wb.sheetnames:
            print(f"  No {sheet_name} sheet found — using seed tickers")
            wb.close()
            return seed_tickers
        ws = wb[sheet_name]
        tickers = []
        for row in ws.iter_rows(min_row=data_start_row, max_col=1, values_only=True):
            val = row[0]
            if val and str(val).strip().upper() not in ["", "TICKER", "NAN"]:
                tickers.append(str(val).strip().upper())
        wb.close()
        if tickers:
            print(f"  Read {len(tickers)} tickers from {sheet_name}: {tickers}")
            return tickers
        print(f"  {sheet_name} sheet empty — using seed tickers")
        return seed_tickers
    except Exception as e:
        print(f"  Could not read {sheet_name} ({e}) — using seed tickers")
        return seed_tickers

# ── Style helpers ─────────────────────────────────────────────────────────────
def get_styles():
    return {
        "hdr_font":       Font(name="Calibri", bold=True, color="FFFFFF", size=12),
        "hdr_fill":       PatternFill("solid", start_color="0D1B2A"),
        "ticker_font":    Font(name="Calibri", bold=True, size=11, color="0D1B2A"),
        "ticker_fill_e":  PatternFill("solid", start_color="E8EDF2"),
        "ticker_fill_o":  PatternFill("solid", start_color="FFFFFF"),
        "sector_font":    Font(name="Calibri", size=11, color="2C3E50"),
        "bull_fill":      PatternFill("solid", start_color="1E8449"),
        "bull_font":      Font(name="Calibri", bold=True, color="FFFFFF", size=11),
        "bear_fill":      PatternFill("solid", start_color="C0392B"),
        "bear_font":      Font(name="Calibri", bold=True, color="FFFFFF", size=11),
        "bull_date_fill": PatternFill("solid", start_color="D5F5E3"),
        "bull_date_font": Font(name="Calibri", size=11, color="1E8449"),
        "bear_date_fill": PatternFill("solid", start_color="FADBD8"),
        "bear_date_font": Font(name="Calibri", size=11, color="C0392B"),
        "above_fill":     PatternFill("solid", start_color="D5F5E3"),
        "above_font":     Font(name="Calibri", bold=True, size=11, color="1E8449"),
        "below_fill":     PatternFill("solid", start_color="FADBD8"),
        "below_font":     Font(name="Calibri", bold=True, size=11, color="C0392B"),
        "na_font":        Font(name="Calibri", size=11, color="AAAAAA", italic=True),
        "alt_fill":       PatternFill("solid", start_color="F4F6F8"),
        "white_fill":     PatternFill("solid", start_color="FFFFFF"),
        "med":            Side(style="medium", color="0D1B2A"),
        "thin":           Side(style="thin",   color="BDC3C7"),
        "center":         Alignment(horizontal="center", vertical="center"),
    }

def _write_header(ws, title, generated_at, col_letter, has_instruction=False):
    s = get_styles()
    ws.merge_cells(f"A1:{col_letter}1")
    ws["A1"].value     = f"  {title}  —  Weekly MACD & RSI Signal Tracker"
    ws["A1"].font      = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill      = PatternFill("solid", start_color="0D1B2A")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[1].height = 32

    ws.merge_cells(f"A2:{col_letter}2")
    ws["A2"].value     = (f"  Generated: {generated_at}   |   MACD (12,26,9)   |   "
                          f"RSI({RSI_PERIOD}) vs MA({RSI_MA})   |   200W SMA   |   Weekly   |   Source: Yahoo Finance")
    ws["A2"].font      = Font(name="Calibri", italic=True, size=9, color="FFFFFF")
    ws["A2"].fill      = PatternFill("solid", start_color="1A2E45")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[2].height = 18

    if has_instruction:
        ws.merge_cells(f"A3:{col_letter}3")
        ws["A3"].value     = "  ✏️  Edit tickers in column A below to customise this sheet. Add or remove rows freely — results update on next run."
        ws["A3"].font      = Font(name="Calibri", italic=True, size=10, color="1A2E45")
        ws["A3"].fill      = PatternFill("solid", start_color="EAF2FB")
        ws["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
        ws.row_dimensions[3].height = 20
        ws.row_dimensions[4].height = 6
        return 5  # header row
    else:
        ws.row_dimensions[3].height = 6
        return 4  # header row

def _write_data(ws, results, header_row, headers, col_widths, has_sector=False,
                preserve_order=False):
    s        = get_styles()
    med      = s["med"]; thin = s["thin"]
    thick_b  = Border(left=med,  right=med,  top=med,  bottom=med)
    thin_b   = Border(left=thin, right=thin, top=thin, bottom=thin)
    col_letter = get_column_letter(len(headers))

    ws.row_dimensions[header_row].height = 26
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = s["hdr_font"]; cell.fill = s["hdr_fill"]
        cell.alignment = s["center"]; cell.border = thick_b
        ws.column_dimensions[get_column_letter(col)].width = w

    def sort_key(r):
        order = {"▲ Bullish": 0, "▼ Bearish": 1, "N/A": 2}
        return (order.get(r["MACD Signal"], 2),
                r["MACD Cross Date"] if r["MACD Cross Date"] != "N/A" else "")

    def signal_style(sig):
        if "Bullish" in str(sig):
            return s["bull_fill"], s["bull_font"], s["bull_date_fill"], s["bull_date_font"]
        if "Bearish" in str(sig):
            return s["bear_fill"], s["bear_font"], s["bear_date_fill"], s["bear_date_font"]
        return None, s["na_font"], None, s["na_font"]

    def ma_style(val):
        if "Above" in str(val): return s["above_fill"], s["above_font"]
        if "Below" in str(val): return s["below_fill"], s["below_font"]
        return None, s["na_font"]

    ordered = results if preserve_order else sorted(results, key=sort_key)

    for row_idx, r in enumerate(ordered, header_row + 1):
        is_even   = (row_idx % 2 == 0)
        base_fill = s["alt_fill"] if is_even else s["white_fill"]
        ws.row_dimensions[row_idx].height = 22

        ms_fill, ms_font, md_fill, md_font = signal_style(r["MACD Signal"])
        rs_fill, rs_font, rd_fill, rd_font = signal_style(r["RSI Signal"])
        m2_fill, m2_font                   = ma_style(r["200W MA"])

        tf = s["ticker_fill_e"] if is_even else s["ticker_fill_o"]
        if has_sector:
            row_data = [
                (r["Ticker"],         tf,                   s["ticker_font"]),
                (r.get("Sector",""),  base_fill,            s["sector_font"]),
                (r["MACD Cross Date"],md_fill or base_fill, md_font),
                (r["MACD Signal"],    ms_fill or base_fill, ms_font),
                (r["RSI Cross Date"], rd_fill or base_fill, rd_font),
                (r["RSI Signal"],     rs_fill or base_fill, rs_font),
                (r["200W MA"],        m2_fill or base_fill, m2_font),
            ]
        else:
            row_data = [
                (r["Ticker"],         tf,                   s["ticker_font"]),
                (r["MACD Cross Date"],md_fill or base_fill, md_font),
                (r["MACD Signal"],    ms_fill or base_fill, ms_font),
                (r["RSI Cross Date"], rd_fill or base_fill, rd_font),
                (r["RSI Signal"],     rs_fill or base_fill, rs_font),
                (r["200W MA"],        m2_fill or base_fill, m2_font),
            ]

        for col, (val, f, fn) in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = fn; cell.fill = f
            cell.alignment = s["center"]; cell.border = thin_b

    # Summary bar
    summary_row = header_row + len(results) + 2
    ws.merge_cells(f"A{summary_row}:{col_letter}{summary_row}")
    macd_bull = sum(1 for r in results if "Bullish" in str(r["MACD Signal"]))
    macd_bear = sum(1 for r in results if "Bearish" in str(r["MACD Signal"]))
    rsi_bull  = sum(1 for r in results if "Bullish" in str(r["RSI Signal"]))
    rsi_bear  = sum(1 for r in results if "Bearish" in str(r["RSI Signal"]))
    above_200 = sum(1 for r in results if "Above"   in str(r["200W MA"]))
    below_200 = sum(1 for r in results if "Below"   in str(r["200W MA"]))
    ws[f"A{summary_row}"].value = (
        f"  MACD: {macd_bull} Bullish  |  {macd_bear} Bearish     "
        f"RSI: {rsi_bull} Bullish  |  {rsi_bear} Bearish     "
        f"200W MA: {above_200} Above  |  {below_200} Below     "
        f"Total: {len(results)} tickers")
    ws[f"A{summary_row}"].font      = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    ws[f"A{summary_row}"].fill      = PatternFill("solid", start_color="1A2E45")
    ws[f"A{summary_row}"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[summary_row].height = 20

    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{col_letter}{header_row + len(results)}"

def build_sheet(ws, results, title, generated_at, has_sector=False):
    ws.title = title
    if has_sector:
        headers    = ["Ticker","Sector","MACD Cross Date","MACD Signal","RSI Cross Date","RSI Signal","200W MA"]
        col_widths = [12, 26, 20, 18, 20, 18, 14]
    else:
        headers    = ["Ticker","MACD Cross Date","MACD Signal","RSI Cross Date","RSI Signal","200W MA"]
        col_widths = [12, 20, 18, 20, 18, 14]
    col_letter = get_column_letter(len(headers))
    header_row = _write_header(ws, title, generated_at, col_letter)
    _write_data(ws, results, header_row, headers, col_widths, has_sector=has_sector)

def build_editable_sheet(ws, results, title, generated_at):
    """Sheet where tickers live in col A and are user-editable."""
    ws.title = title
    headers    = ["Ticker","MACD Cross Date","MACD Signal","RSI Cross Date","RSI Signal","200W MA"]
    col_widths = [14, 20, 18, 20, 18, 14]
    col_letter = get_column_letter(len(headers))
    header_row = _write_header(ws, title, generated_at, col_letter, has_instruction=True)
    _write_data(ws, results, header_row, headers, col_widths, preserve_order=True)

# ── Google Drive: download existing file ──────────────────────────────────────
def get_drive_service():
    creds_json = json.loads(os.environ["GDRIVE_CREDENTIALS"])
    creds = Credentials(
        token=creds_json["token"],
        refresh_token=creds_json["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=creds_json["client_id"],
        client_secret=creds_json["client_secret"],
    )
    return build("drive", "v3", credentials=creds)

def download_existing_file(service, file_id):
    """Download existing spreadsheet bytes from Drive, return None if not found."""
    try:
        request  = service.files().get_media(fileId=file_id)
        buffer   = io.BytesIO()
        from googleapiclient.http import MediaIoBaseDownload
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buffer.seek(0)
        print("  Downloaded existing spreadsheet from Drive")
        return buffer.read()
    except Exception as e:
        print(f"  No existing file found ({e}) — will use seed tickers")
        return None

def upload_to_drive(service, wb, file_id):
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )
    service.files().update(fileId=file_id, media_body=media).execute()
    print(f"✅ Uploaded to Google Drive (file ID: {file_id})")

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Connect to Drive and download existing file to read editable tickers
    print("Connecting to Google Drive...")
    service       = get_drive_service()
    file_id       = os.environ["GDRIVE_FILE_ID"]
    existing_bytes = download_existing_file(service, file_id)

    # Read editable tickers from Watchlist and Financials sheets
    print("\nReading editable tickers from spreadsheet...")
    watchlist_tickers  = read_tickers_from_sheet(existing_bytes, "Watchlist",  WATCHLIST_SEED,  data_start_row=6) if existing_bytes else WATCHLIST_SEED
    financials_tickers = read_tickers_from_sheet(existing_bytes, "Financials", FINANCIALS_SEED, data_start_row=6) if existing_bytes else FINANCIALS_SEED

    # Get QQQ holdings
    print("\nFetching QQQ holdings...")
    try:
        tables = pd.read_html("https://en.wikipedia.org/wiki/Invesco_QQQ_Trust")
        holdings_df = None
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            if any("ticker" in c or "symbol" in c for c in cols):
                holdings_df = t
                break
        if holdings_df is not None:
            ticker_col  = [c for c in holdings_df.columns if "ticker" in str(c).lower() or "symbol" in str(c).lower()][0]
            qqq_tickers = [str(t).strip().upper() for t in holdings_df[ticker_col].dropna() if str(t).strip() not in ["","nan"]]
        else:
            raise ValueError()
    except:
        print("Wikipedia scrape failed — using fallback")
        qqq_tickers = [
            "MSFT","AAPL","NVDA","AMZN","META","TSLA","GOOGL","GOOG","AVGO","COST",
            "NFLX","TMUS","AMD","PEP","LIN","CSCO","ADBE","TXN","QCOM","AMGN",
            "INTU","ISRG","CMCSA","AMAT","MU","HON","BKNG","LRCX","VRTX","REGN",
            "ADP","PANW","SBUX","GILD","ADI","MDLZ","KLAC","SNPS","CDNS","MELI",
            "PYPL","CSX","CTAS","ASML","MAR","ORLY","MNST","NXPI","WDAY","FTNT",
            "MRVL","CHTR","PCAR","CRWD","IDXX","ODFL","ROST","KDP","CEG","DXCM",
            "TTD","FANG","ON","VRSK","CTSH","GEHC","EXC","FAST","BIIB","XEL",
            "CCEP","DLTR","CSGP","CPRT","BKR","EA","ANSS","TEAM","ZS","ILMN",
            "GFS","SIRI","WBA","NTAP","SPLK","ENPH","LCID","RIVN","ZM","PARA",
            "OKTA","DDOG","ABNB","COIN","RBLX","HOOD","LYFT","UBER","PINS","SNAP"
        ]
    print(f"Found {len(qqq_tickers)} QQQ tickers")

    # Get SPMO holdings
    spmo_tickers = get_spmo_holdings()

    # Run analysis
    print("\nRunning analysis...")
    qqq_results        = run_tickers(qqq_tickers,       "QQQ")
    watchlist_results  = run_tickers(watchlist_tickers,  "Watchlist")
    financials_results = run_tickers(financials_tickers, "Financials")
    spmo_results       = run_tickers(spmo_tickers,       "SPMO")

    sector_tickers     = list(SECTORS.keys())
    sector_results_raw = run_tickers(sector_tickers,     "Sectors")
    sector_results     = [{**r, "Sector": SECTORS.get(r["Ticker"], "")} for r in sector_results_raw]

    # Build workbook
    generated_at = datetime.now().strftime("%B %d, %Y  %I:%M %p UTC")
    wb = Workbook()
    build_sheet(wb.active,         qqq_results,       "QQQ Holdings", generated_at)
    build_editable_sheet(wb.create_sheet(), watchlist_results,  "Watchlist",    generated_at)
    build_sheet(wb.create_sheet(), sector_results,     "Sectors",      generated_at, has_sector=True)
    build_sheet(wb.create_sheet(), spmo_results,       "SPMO",         generated_at)
    build_editable_sheet(wb.create_sheet(), financials_results, "Financials",   generated_at)

    # Upload to Drive
    upload_to_drive(service, wb, file_id)

    print(f"\nDone! Generated: {generated_at}")
    print(f"   QQQ        → {sum(1 for r in qqq_results        if 'Bullish' in str(r['MACD Signal']))} bull | {sum(1 for r in qqq_results        if 'Bearish' in str(r['MACD Signal']))} bear")
    print(f"   Watchlist  → {sum(1 for r in watchlist_results   if 'Bullish' in str(r['MACD Signal']))} bull | {sum(1 for r in watchlist_results   if 'Bearish' in str(r['MACD Signal']))} bear")
    print(f"   Sectors    → {sum(1 for r in sector_results      if 'Bullish' in str(r['MACD Signal']))} bull | {sum(1 for r in sector_results      if 'Bearish' in str(r['MACD Signal']))} bear")
    print(f"   SPMO       → {sum(1 for r in spmo_results        if 'Bullish' in str(r['MACD Signal']))} bull | {sum(1 for r in spmo_results        if 'Bearish' in str(r['MACD Signal']))} bear")
    print(f"   Financials → {sum(1 for r in financials_results  if 'Bullish' in str(r['MACD Signal']))} bull | {sum(1 for r in financials_results  if 'Bearish' in str(r['MACD Signal']))} bear")
