"""
UrbanCo Stock Dashboard - Data Fetcher
Fetches OHLCV data for UrbanCo (NSE+BSE) and peer stocks via yfinance.
Embeds data into dashboard.html for offline viewing.

Usage: python fetch_data.py
"""

import json
import os
import re
from datetime import datetime, timedelta

import pandas as pd
import yfinance as yf
import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(SCRIPT_DIR, "dashboard_data.json")
HTML_FILE = os.path.join(SCRIPT_DIR, "dashboard.html")

# UrbanCo tickers
STOCK_NSE = "URBANCO.NS"
STOCK_BSE = "URBANCO.BO"
NSE_SYMBOL = "URBANCO"

# Peer stocks for comparison table
PEERS = {
    "Urbanco":        "URBANCO.NS",
    "Eternal":        "ETERNAL.NS",
    "Swiggy":         "SWIGGY.NS",
    "Meesho":         "MEESHO.NS",
    "Groww":          "GROWW.NS",
    "Lenskart":       "LENSKART.NS",
    "Physics Wallah": "PWL.NS",
    "Blackbuck":      "BLACKBUCK.NS",
    "Pinelabs":       "PINELABS.NS",
}


def fetch_ohlcv(ticker_symbol: str, period: str = "3mo") -> pd.DataFrame:
    ticker = yf.Ticker(ticker_symbol)
    df = ticker.history(period=period)
    if df.empty:
        print(f"  Warning: No data for {ticker_symbol}")
        return pd.DataFrame()
    df.index = df.index.tz_localize(None)
    # Drop rows with NaN close (incomplete data from today/holidays)
    df = df.dropna(subset=["Close"])
    return df


def fetch_nse_delivery(symbol: str) -> dict:
    delivery = {}
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        })
        s.get("https://www.nseindia.com", timeout=10)
        url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}&section=trade_info"
        r = s.get(url, timeout=10)
        if r.status_code == 200:
            sec = r.json().get("securityWiseDP", {})
            if sec:
                pct = sec.get("delToTradQty", 0)
                delivery[datetime.now().strftime("%Y-%m-%d")] = float(pct) if pct else 0
                print(f"  NSE delivery % today: {pct}%")
    except Exception as e:
        print(f"  Could not fetch delivery data: {e}")
    return delivery


def fetch_nse_intimations(symbol: str) -> list:
    """Fetch recent corporate announcements/intimations from NSE."""
    intimations = []
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        })
        s.get("https://www.nseindia.com", timeout=10)
        url = f"https://www.nseindia.com/api/corporate-announcements?index=equities&symbol={symbol}"
        r = s.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            for item in data[:20]:  # last 20 announcements
                intimations.append({
                    "date": item.get("an_dt", ""),
                    "subject": item.get("desc", ""),
                    "category": item.get("attchmntText", ""),
                })
            print(f"  Got {len(intimations)} intimations")
    except Exception as e:
        print(f"  Could not fetch intimations: {e}")
    return intimations


def compute_combined_volumes(nse_df: pd.DataFrame, bse_df: pd.DataFrame,
                              delivery_pct: dict) -> dict:
    """Combine NSE+BSE volumes. Compute delivery vol, rolling avgs for all."""
    if nse_df.empty:
        return {
            "dates": [], "daily": [], "nse_daily": [], "bse_daily": [],
            "weekly_avg": [], "monthly_avg": [],
            "deli_vol_daily": [], "deli_vol_weekly_avg": [], "deli_vol_monthly_avg": [],
            "deli_pct_daily": [], "deli_pct_weekly_avg": [], "deli_pct_monthly_avg": [],
        }

    c = nse_df[["Volume"]].copy()
    c.columns = ["nse_vol"]
    c["bse_vol"] = 0

    if not bse_df.empty:
        bse_aligned = bse_df[["Volume"]].reindex(c.index, fill_value=0)
        c["bse_vol"] = bse_aligned["Volume"].fillna(0)

    c["total"] = c["nse_vol"] + c["bse_vol"]

    # Map delivery % to each date, compute absolute delivery volume
    c["deli_pct"] = c.index.to_series().apply(
        lambda d: delivery_pct.get(d.strftime("%Y-%m-%d"), 0)
    )
    c["deli_vol"] = (c["total"] * c["deli_pct"] / 100).round(0)

    # Rolling averages for total volume
    c["weekly_avg"] = c["total"].rolling(5, min_periods=1).mean()
    c["monthly_avg"] = c["total"].rolling(22, min_periods=1).mean()

    # Rolling averages for delivery volume
    c["deli_vol_weekly"] = c["deli_vol"].rolling(5, min_periods=1).mean()
    c["deli_vol_monthly"] = c["deli_vol"].rolling(22, min_periods=1).mean()

    # Rolling averages for delivery %
    c["deli_pct_weekly"] = c["deli_pct"].rolling(5, min_periods=1).mean()
    c["deli_pct_monthly"] = c["deli_pct"].rolling(22, min_periods=1).mean()

    dates = c.index.strftime("%Y-%m-%d").tolist()
    return {
        "dates": dates,
        "daily": [int(v) for v in c["total"]],
        "nse_daily": [int(v) for v in c["nse_vol"]],
        "bse_daily": [int(v) for v in c["bse_vol"]],
        "weekly_avg": [round(v) for v in c["weekly_avg"]],
        "monthly_avg": [round(v) for v in c["monthly_avg"]],
        "deli_vol_daily": [int(v) for v in c["deli_vol"]],
        "deli_vol_weekly_avg": [round(v) for v in c["deli_vol_weekly"]],
        "deli_vol_monthly_avg": [round(v) for v in c["deli_vol_monthly"]],
        "deli_pct_daily": [round(v, 1) for v in c["deli_pct"]],
        "deli_pct_weekly_avg": [round(v, 1) for v in c["deli_pct_weekly"]],
        "deli_pct_monthly_avg": [round(v, 1) for v in c["deli_pct_monthly"]],
    }


def compute_delivery_pct(nse_df: pd.DataFrame, nse_delivery_today: dict) -> dict:
    """Estimate delivery % series."""
    if nse_df.empty:
        return {}
    avg_vol = nse_df["Volume"].mean()
    result = {}
    for idx, row in nse_df.iterrows():
        ds = idx.strftime("%Y-%m-%d")
        if ds in nse_delivery_today and nse_delivery_today[ds] > 0:
            result[ds] = nse_delivery_today[ds]
        else:
            vol_ratio = row["Volume"] / avg_vol if avg_vol > 0 else 1
            result[ds] = round(max(20, min(75, 45 - (vol_ratio - 1) * 15)), 1)
    return result


def compute_prices(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"dates": [], "open": [], "high": [], "low": [], "close": [], "change_pct": []}
    df = df.copy()
    # Drop rows where Close is NaN (incomplete/holiday data)
    df = df.dropna(subset=["Close"])
    df["Change_Pct"] = df["Close"].pct_change() * 100
    return {
        "dates": df.index.strftime("%Y-%m-%d").tolist(),
        "open": [round(float(v), 2) if pd.notna(v) else 0 for v in df["Open"]],
        "high": [round(float(v), 2) if pd.notna(v) else 0 for v in df["High"]],
        "low": [round(float(v), 2) if pd.notna(v) else 0 for v in df["Low"]],
        "close": [round(float(v), 2) for v in df["Close"]],
        "change_pct": [round(float(v), 2) if pd.notna(v) else 0 for v in df["Change_Pct"]],
    }


def compute_peer_performance(peers: dict) -> list:
    """Compute % change over various periods for peer stocks."""
    periods = {
        "1d": 1, "1w": 5, "30d": 22, "90d": 63, "180d": 126, "365d": 252,
    }
    results = []

    for name, ticker in peers.items():
        print(f"  Fetching {name} ({ticker})...")
        row = {"name": name, "ticker": ticker}
        try:
            t = yf.Ticker(ticker)
            # Fetch 1yr+ of data
            df = t.history(period="2y")
            if df.empty:
                for k in periods:
                    row[k] = None
                row["price"] = None
                results.append(row)
                continue

            df.index = df.index.tz_localize(None)
            # Use last non-NaN close (today's may be NaN if market is open)
            closes = df["Close"].dropna()
            if closes.empty:
                for k in periods:
                    row[k] = None
                row["price"] = None
                row["mcap"] = None
                results.append(row)
                continue

            latest_close = closes.iloc[-1]
            row["price"] = round(float(latest_close), 1)

            # Market cap in thousands of crores
            info = t.info
            mcap = info.get("marketCap", 0)
            row["mcap"] = round(mcap / 1e11, 1) if mcap else None  # /1e11 = k Cr

            for label, trading_days in periods.items():
                if len(closes) > trading_days:
                    old_close = closes.iloc[-(trading_days + 1)]
                    pct = ((latest_close - old_close) / old_close) * 100
                    row[label] = round(float(pct), 1)
                else:
                    row[label] = None
        except Exception as e:
            print(f"    Error: {e}")
            for k in periods:
                row[k] = None
            row["price"] = None
            row["mcap"] = None

        results.append(row)

    return results


def fetch_results_dates(symbol: str) -> list:
    """Fetch actual results/board meeting dates from NSE announcements."""
    dates = []
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        })
        s.get("https://www.nseindia.com", timeout=10)
        url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&symbol={symbol}"
            f"&from_date=01-01-2025&to_date={datetime.now().strftime('%d-%m-%Y')}"
        )
        r = s.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            seen = set()
            for item in data:
                desc = item.get("desc", "").lower()
                if any(k in desc for k in ["outcome of board meeting", "financial result"]):
                    # Parse date from "23-Jan-2026 16:55:05" format
                    dt_str = item.get("an_dt", "")
                    try:
                        dt = datetime.strptime(dt_str.split()[0], "%d-%b-%Y")
                        ds = dt.strftime("%Y-%m-%d")
                        if ds not in seen:
                            seen.add(ds)
                            dates.append(ds)
                    except ValueError:
                        pass
            dates.sort()
    except Exception as e:
        print(f"  Could not fetch results dates: {e}")
    return dates


def main():
    print(f"UrbanCo Dashboard - Data Fetcher")
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    # --- UrbanCo Data ---
    print("[1/5] Fetching UrbanCo NSE data...")
    nse_df = fetch_ohlcv(STOCK_NSE, period="3mo")
    print(f"  {len(nse_df)} days")

    print("[2/5] Fetching UrbanCo BSE data...")
    bse_df = fetch_ohlcv(STOCK_BSE, period="3mo")
    print(f"  {len(bse_df)} days")

    print("[3/5] Fetching delivery data & intimations...")
    nse_delivery_today = fetch_nse_delivery(NSE_SYMBOL)
    intimations = fetch_nse_intimations(NSE_SYMBOL)

    # Delivery % first (needed for combined volumes)
    delivery_pct = compute_delivery_pct(nse_df, nse_delivery_today)

    # Combined volumes (includes delivery vol and delivery % rolling avgs)
    print("  Computing combined volumes...")
    combined_volumes = compute_combined_volumes(nse_df, bse_df, delivery_pct)

    # Prices
    print("  Computing prices...")
    nse_prices = compute_prices(nse_df)

    # Results dates (from NSE board meeting announcements)
    print("  Fetching results dates from NSE...")
    results_dates = fetch_results_dates(NSE_SYMBOL)
    print(f"  Found {len(results_dates)} results dates: {results_dates}")

    # Stock info
    stock_info = {"name": "Urban Company Limited"}
    try:
        info = yf.Ticker(STOCK_NSE).info
        stock_info = {
            "name": info.get("longName", "Urban Company Limited"),
            "sector": info.get("sector", "N/A"),
            "industry": info.get("industry", "N/A"),
            "market_cap": info.get("marketCap", 0),
            "pe_ratio": info.get("trailingPE", 0),
            "fifty_two_week_high": info.get("fiftyTwoWeekHigh", 0),
            "fifty_two_week_low": info.get("fiftyTwoWeekLow", 0),
        }
    except Exception as e:
        print(f"  Could not fetch info: {e}")

    # --- Peer Comparison ---
    print("[4/5] Fetching peer stock performance...")
    peer_data = compute_peer_performance(PEERS)

    # --- Build JSON ---
    print("[5/5] Building dashboard data...")
    dashboard_data = {
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "stock_info": stock_info,
        "combined_volumes": combined_volumes,
        "delivery_pct": delivery_pct,
        "prices": nse_prices,
        "results_dates": results_dates,
        "intimations": intimations,
        "peer_performance": peer_data,
    }

    # Write JSON
    with open(DATA_FILE, "w") as f:
        json.dump(dashboard_data, f, indent=2)
    print(f"  Saved {DATA_FILE}")

    # Embed into HTML
    if os.path.exists(HTML_FILE):
        with open(HTML_FILE, "r", encoding="utf-8") as f:
            html = f.read()

        data_js = f"var EMBEDDED_DATA = {json.dumps(dashboard_data)};"
        html = re.sub(
            r'var EMBEDDED_DATA = \{.*?\};',
            '// %%EMBEDDED_DATA_PLACEHOLDER%%',
            html, flags=re.DOTALL,
        )
        html = html.replace('// %%EMBEDDED_DATA_PLACEHOLDER%%', data_js)

        with open(HTML_FILE, "w", encoding="utf-8") as f:
            f.write(html)
        print("  Embedded data into dashboard.html")

    print(f"\nDone! Open dashboard.html in your browser.")


if __name__ == "__main__":
    main()
