import numpy as np
import pandas as pd

import statsmodels.api as sm
import requests
import os
from io import BytesIO
from zipfile import ZipFile
from scipy.optimize import minimize


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
def get_stock_returns(tickers, max_date):
    returns = {}
    for t in tickers:
        daily = get_price_history(t)
        if not daily.empty:
            monthly = daily.resample("ME").last().pct_change()
            monthly = monthly.replace([np.inf, -np.inf], np.nan).dropna()
            monthly = monthly.clip(lower=-0.99, upper=1.0)  # cap extreme returns

            monthly = monthly[monthly.index <= max_date]  # ✅ Truncate to FF data range
            if not monthly.empty:
                print(f"{t}: monthly return range = {monthly.index.min()} to {monthly.index.max()}")
                returns[t] = monthly
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
    df['date'] = pd.to_datetime(df['date'], format='%Y%m', errors='coerce') + pd.offsets.MonthEnd(0)
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

        # Align and clean
        Y_aligned, X_aligned = Y.align(X, join='inner')
        print(f"{ticker}: overlapping months = {len(Y_aligned.dropna())}")

        combined = pd.concat([Y_aligned, X_aligned], axis=1).dropna()
        if len(combined) < 30:
            print(f"Skipping {ticker} due to insufficient overlap")
            continue

        Y_clean = combined.iloc[:, 0]
        X_clean = sm.add_constant(combined.iloc[:, 1:])
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
    factor_returns = get_factor_returns()
    stock_returns = get_stock_returns(all_tickers, factor_returns.index.max())
    print("\n📊 Preview of stock return data:")
    print(stock_returns.head(10))

    print("\n📉 Missing values per ticker:")
    print(stock_returns.isnull().sum())

    print("\n✅ Total valid rows after dropna:")
    print(len(stock_returns.dropna()))

    factor_matrix = compute_factor_loadings(stock_returns, factor_returns)


    valid_tickers = factor_matrix.index.tolist()
    if not valid_tickers:
        return {'status': 'error', 'message': 'No valid tickers with factor loadings.'}

    ticker_indices = [i for i, t in enumerate(all_tickers) if t in valid_tickers]
    tickers = valid_tickers
    current_weights = np.array([original_weights[i] for i in ticker_indices])


    # Normalize factor exposures and targets
    F = factor_matrix.loc[tickers].values
    F = (F - F.mean(axis=0)) / F.std(axis=0)
    target = np.array([target_exposures[f] for f in factor_matrix.columns])
    target = (target - target.mean()) / target.std()

    # Adjusted objective function with tunable penalty
    lambda_exposure = 10
    def objective(w):
        tracking_error = np.sum((w - current_weights) ** 2)
        exposure_penalty = np.sum((F.T @ w - target) ** 2)
        return tracking_error + lambda_exposure * exposure_penalty



    def optimize(turnover=None):
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
        if turnover_limit is not None:
            constraints.append({
                'type': 'ineq',
                'fun': lambda w: turnover_limit - np.sum(np.abs(w - current_weights))
            })

        bounds = [(0, 1) for _ in range(len(tickers))]

        print("🔎 Optimizing with:")
        print(" - Current weights:", current_weights)
        print(" - Target exposures:", target)
        print(" - Exposure matrix shape:", F.shape)
        print(" - Exposure matrix:\n", F)
        print(" - Bounds:", bounds)
        print(" - Constraints:", constraints)


        result = minimize(
            objective,
            x0=current_weights,
            bounds=bounds,
            constraints=constraints,
            method='SLSQP'
        )


        if not result.success:
            print("❌ Optimization failed:")
            print(" - Status:", result.status)
            print(" - Message:", result.message)
            print(" - Final weights:", result.x)
            print(" - Achieved exposures:", F.T @ result.x)
            print(" - Distance to target:", F.T @ result.x - target)
            print(" - Sum(weights):", np.sum(result.x))
            print(" - Turnover:", np.sum(np.abs(result.x - current_weights)))

        return result

    # First attempt: with turnover
    result = optimize(turnover_limit)

    if not result.success:
        print("⚠️ Turnover constraint caused failure. Retrying without it.")
        result = optimize(None)
        if not result.success:
            return {
                'status': 'error',
                'message': f"Optimization failed: {result.message}",
                'exposure_matrix': F.tolist(),
                'target_exposures': target.tolist(),
                'constraints': str(constraints),
                'weights': current_weights.tolist()
            }


    optimized_weights = result.x

    return {
        'status': 'success',
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

