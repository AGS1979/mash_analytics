import numpy as np
import pandas as pd
import cvxpy as cp
import statsmodels.api as sm
import requests
import os
from io import BytesIO
from zipfile import ZipFile

FMP_API_KEY = os.environ["FMP_API_KEY"]
FMP_BASE_URL = "https://financialmodelingprep.com/api/v3"

# Fetch historical price data from FMP
def get_price_history(ticker, period="3y"):
    url = f"{FMP_BASE_URL}/historical-price-full/{ticker}?serietype=line&apikey={FMP_API_KEY}"
    r = requests.get(url)
    data = r.json()
    try:
        prices = pd.DataFrame(data["historical"])
        prices['date'] = pd.to_datetime(prices['date'])
        prices = prices.set_index('date').sort_index()
        return prices['close'].pct_change().dropna()
    except Exception:
        return pd.Series()

# Pull historical prices for multiple stocks
def get_stock_returns(tickers, lookback_months=60):
    returns = {}
    for t in tickers:
        daily_prices = get_price_history(t)
        if not daily_prices.empty:
            # Ensure index is at month end to match Kenneth French format
            monthly_prices = daily_prices.resample("ME").last()
            monthly_prices.index = monthly_prices.index.to_period("M").to_timestamp("M")
            monthly_returns = monthly_prices.pct_change().dropna()
            returns[t] = monthly_returns
    return pd.DataFrame(returns).dropna()


# Load Kenneth French 5-factor monthly data from local file
def get_factor_returns():
    file_path = "data/F-F_Research_Data_5_Factors_2x3.csv"
    df = pd.read_csv(file_path, skiprows=3)

    # Defensive: Remove footer rows starting from "Annual" if it exists
    annual_rows = df[df.iloc[:, 0].astype(str).str.startswith("Annual", na=False)]
    if not annual_rows.empty:
        end_idx = annual_rows.index[0]
        df = df.iloc[:end_idx]

    df.rename(columns={df.columns[0]: "date"}, inplace=True)
    df['date'] = pd.to_datetime(df['date'], format='%Y%m', errors='coerce')
    df = df.dropna(subset=["date"])  # remove malformed dates
    df.set_index('date', inplace=True)

    # Convert percentage values to decimals
    df = df.apply(pd.to_numeric, errors='coerce')  # ensure all are numbers
    df = df.astype(float) / 100

    # Rename Mkt-RF to MKT for consistency, drop RF
    df.rename(columns={"Mkt-RF": "MKT"}, inplace=True)
    df.drop(columns=["RF"], inplace=True)

    return df


# Compute regression-based factor loadings
def compute_factor_loadings(stock_returns, factor_returns):
    loadings = {}
    for ticker in stock_returns.columns:
        Y = stock_returns[ticker]
        X = factor_returns

        # Align on index and drop any rows with NaNs
        Y_aligned, X_aligned = Y.align(X, join='inner', axis=0)
        combined = pd.concat([Y_aligned, X_aligned], axis=1).dropna()
        Y_clean = combined.iloc[:, 0]
        X_clean = combined.iloc[:, 1:]


        if len(Y_clean) < 30:
            print(f"Skipping {ticker} due to insufficient data after dropna.")
            continue

        X_clean = sm.add_constant(X_clean)
        model = sm.OLS(Y_clean, X_clean).fit()
        loadings[ticker] = model.params.drop('const')
    return pd.DataFrame(loadings).T

# Read and validate portfolio from a CSV file
def load_portfolio_from_csv(file_path):
    df = pd.read_csv(file_path)
    if 'ticker' not in df.columns or 'weight' not in df.columns:
        raise ValueError("CSV must contain 'ticker' and 'weight' columns.")
    return df.to_dict(orient='records')

# Main optimizer function
def run_factor_optimizer_csv(csv_file_path, target_exposures, turnover_limit=None):
    portfolio = load_portfolio_from_csv(csv_file_path)
    all_tickers = [p['ticker'] for p in portfolio]
    original_weights = np.array([p['weight'] for p in portfolio])

    # Get returns and factor data
    stock_returns = get_stock_returns(all_tickers)
    factor_returns = get_factor_returns()
    factor_matrix = compute_factor_loadings(stock_returns, factor_returns)

    # Keep only tickers that successfully got factor loadings
    valid_tickers = factor_matrix.index.tolist()
    if not valid_tickers:
        return {'status': 'error', 'message': 'No valid tickers with factor loadings.'}

    ticker_indices = [i for i, t in enumerate(all_tickers) if t in valid_tickers]
    tickers = valid_tickers
    current_weights = np.array([original_weights[i] for i in ticker_indices])

    # Build optimization matrices
    F = factor_matrix.loc[tickers].values
    try:
        target = np.array([target_exposures[f] for f in factor_matrix.columns])
    except KeyError as e:
        return {'status': 'error', 'message': f"Missing target exposure for factor: {e}"}

    # Set up optimization
    w = cp.Variable(len(tickers))
    objective = cp.Minimize(cp.sum_squares(w - current_weights))
    constraints = [cp.sum(w) == 1, w >= 0, F.T @ w == target]

    if turnover_limit:
        constraints.append(cp.norm1(w - current_weights) <= turnover_limit)

    problem = cp.Problem(objective, constraints)

    try:
        problem.solve()
        optimized_weights = w.value
    except Exception as e:
        return {'status': 'error', 'message': str(e)}

    return {
        'status': problem.status,
        'optimized_weights': [
            {'ticker': t, 'weight': float(round(wi, 6))}
            for t, wi in zip(tickers, optimized_weights)
        ],
        'target_exposures': target_exposures,
        'achieved_exposures': dict(zip(
            factor_matrix.columns,
            (F.T @ optimized_weights).round(6).tolist()
        ))
    }
