import numpy as np
import pandas as pd
import statsmodels.api as sm
import requests
import os
from scipy.optimize import minimize

# ————————————————
# 1) LOAD FMP & K. French Data
# ————————————————
FMP_API_KEY = os.environ["FMP_API_KEY"]
FMP_BASE_URL = "https://financialmodelingprep.com/api/v3"

def get_price_history(ticker):
    """
    Fetch daily price history for a ticker from FMP,
    return a pandas Series of daily returns (pct_change).
    """
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

def get_stock_returns(tickers, max_date):
    """
    For each ticker in tickers, fetch daily returns, resample to month-end,
    and return a DataFrame of monthly returns truncated to max_date.
    """
    returns = {}
    for t in tickers:
        daily = get_price_history(t)
        if not daily.empty:
            monthly = daily.resample("ME").last().pct_change()
            monthly = monthly.replace([np.inf, -np.inf], np.nan).dropna()
            monthly = monthly.clip(lower=-0.99, upper=1.0)
            monthly = monthly[monthly.index <= max_date]
            if not monthly.empty:
                print(f"{t}: monthly return range = {monthly.index.min()} to {monthly.index.max()}")
                returns[t] = monthly
    return pd.DataFrame(returns).dropna()

def get_factor_returns():
    """
    Load Kenneth French 5-factor monthly data from a local CSV:
    - Skip header rows
    - Drop footer starting at "Annual"
    - Parse the YYYYMM date, convert to month-end timestamp
    - Convert percentages to decimals
    - Rename "Mkt-RF" → "MKT" and drop "RF"
    """
    file_path = "data/F-F_Research_Data_5_Factors_2x3.csv"
    df = pd.read_csv(file_path, skiprows=3)

    # Drop footer rows beginning with "Annual" (if present)
    annual_rows = df[df.iloc[:, 0].astype(str).str.startswith("Annual", na=False)]
    if not annual_rows.empty:
        end_idx = annual_rows.index[0]
        df = df.iloc[:end_idx]

    df.rename(columns={df.columns[0]: "date"}, inplace=True)
    df['date'] = pd.to_datetime(df['date'], format='%Y%m', errors='coerce') + pd.offsets.MonthEnd(0)
    df = df.dropna(subset=["date"])
    df.set_index('date', inplace=True)

    # Convert all columns to numeric, then to decimals
    df = df.apply(pd.to_numeric, errors='coerce') / 100.0

    # Rename and drop risk-free column
    df.rename(columns={"Mkt-RF": "MKT"}, inplace=True)
    df.drop(columns=["RF"], inplace=True)

    return df


# ————————————————————————
# 2) COMPUTE FACTOR LOADINGS
# ————————————————————————
def compute_factor_loadings(stock_returns, factor_returns):
    """
    For each ticker (column) in stock_returns, run an OLS regression
    of that ticker's monthly returns on the factor_returns (with intercept).
    Return a DataFrame of factor loadings (coefficients), indexed by ticker.
    """
    loadings = {}
    for ticker in stock_returns.columns:
        Y = stock_returns[ticker]
        X = factor_returns

        # Align on date index (inner join)
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

    return pd.DataFrame(loadings).T  # shape = (num_tickers, num_factors)


# ——————————————————————
# 3) LOAD PORTFOLIO CSV
# ——————————————————————
def load_portfolio_from_csv(file_path):
    """
    Read a CSV with columns 'ticker' and 'weight', return a list of dicts.
    """
    df = pd.read_csv(file_path)
    if 'ticker' not in df.columns or 'weight' not in df.columns:
        raise ValueError("CSV must contain 'ticker' and 'weight' columns.")
    return df.to_dict(orient='records')


# ————————————————————————————————————————————————————————————————————————
# 4) MAIN OPTIMIZER: run_factor_optimizer_csv(csv_file_path, target_exposures, turnover_limit)
# ————————————————————————————————————————————————————————————————————————
def run_factor_optimizer_csv(csv_file_path, target_exposures, turnover_limit=None):
    """
    1) Load the starting portfolio from CSV (ticker, weight).
    2) Fetch factor and stock returns, compute factor loadings.
    3) Standardize factor loadings (F) and target exposures in the same space.
    4) Solve a QP (via SLSQP) to find new weights w ∈ [0,1] summing to 1,
       minimizing (w - current_weights)^2 + λ * ||F^T w - target||^2,
       subject to an optional turnover constraint sum_i |w_i - current_i| ≤ turnover_limit.
    5) If the turnover-constrained solve fails, retry without the turnover constraint.
    Return a dict with status, optimized_weights, and standardized exposures.
    """
    # 4.1) Read portfolio CSV
    portfolio = load_portfolio_from_csv(csv_file_path)
    all_tickers = [p['ticker'] for p in portfolio]
    original_weights = np.array([p['weight'] for p in portfolio])

    # 4.2) Fetch factor & stock returns
    factor_returns = get_factor_returns()
    stock_returns = get_stock_returns(all_tickers, factor_returns.index.max())

    print("\n📊 Preview of stock return data:")
    print(stock_returns.head(10))

    print("\n📉 Missing values per ticker:")
    print(stock_returns.isnull().sum())

    print("\n✅ Total valid rows after dropna:")
    print(len(stock_returns.dropna()))

    # 4.3) Compute factor loadings
    factor_matrix = compute_factor_loadings(stock_returns, factor_returns)
    valid_tickers = factor_matrix.index.tolist()
    if not valid_tickers:
        return {'status': 'error', 'message': 'No valid tickers with factor loadings.'}

    # Only keep tickers with valid loadings
    ticker_indices = [i for i, t in enumerate(all_tickers) if t in valid_tickers]
    tickers = valid_tickers
    current_weights = np.array([original_weights[i] for i in ticker_indices])

    # 4.4) Build raw factor‐loading matrix (n × 5)
    F_raw = factor_matrix.loc[tickers].values
    # Compute means and stds of each factor‐column
    F_mean = F_raw.mean(axis=0)   # shape=(5,)
    F_std  = F_raw.std(axis=0)    # shape=(5,)
    # Standardize F so each column has zero mean, unit std
    F = (F_raw - F_mean) / F_std

    # 4.5) Build & standardize the target in the same space as F
    raw_target = np.array([target_exposures[f] for f in factor_matrix.columns])  # shape=(5,)
    target = (raw_target - F_mean) / F_std  # standardized target

    # 4.6) Penalty coefficient for exposure mismatch
    lambda_exposure = 10.0

    # 4.7) Objective function: (w – current)^2 + λ * ||F^T w – target||^2
    def objective(w):
        tracking_error = np.sum((w - current_weights) ** 2)
        exposure_penalty = np.sum((F.T @ w - target) ** 2)
        return tracking_error + lambda_exposure * exposure_penalty

    # 4.8) Optimize function that respects an optional turnover constraint
    def optimize(local_turnover=None):
        # (a) Always enforce sum(w) = 1
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

        # (b) If local_turnover is not None, add a smooth turnover constraint
        if local_turnover is not None:
            def turnover_constraint(w):
                eps = 1e-8
                return local_turnover - np.sum(np.sqrt((w - current_weights)**2 + eps))
            constraints.append({'type': 'ineq', 'fun': turnover_constraint})

        # (c) Bounds: 0 ≤ w_i ≤ 1
        bounds = [(0, 1)] * len(tickers)

        # Debug prints
        print("🔎 Optimizing with:")
        print(" - Current weights:", current_weights)
        print(" - Standardized target:", target)
        print(" - Exposure matrix shape:", F.shape)
        print(" - Bounds:", bounds)
        print(" - Constraints:", constraints)

        result = minimize(
            objective,
            x0=current_weights,
            bounds=bounds,
            constraints=constraints,
            method='SLSQP',
            options={'maxiter': 1000, 'disp': True}
        )

        if not result.success:
            print("❌ Optimization failed:")
            print(" - Status:", result.status)
            print(" - Message:", result.message)
            print(" - Final weights:", result.x)
            print(" - Achieved exposures (standardized):", (F.T @ result.x))
            print(" - Target (standardized):", target)
            print(" - Exposure difference:", (F.T @ result.x) - target)
            print(" - ∑(weights):", np.sum(result.x))
            if local_turnover is not None:
                print(" - Turnover:", np.sum(np.abs(result.x - current_weights)))
        return result

    # 4.9) First solve: include turnover constraint if provided
    result = optimize(turnover_limit)
    print(">>> First solve returned success =", result.success)

    # 4.10) If that fails, retry without the turnover constraint
    if not result.success:
        print("⚠️ Turnover constraint caused failure. Retrying *without* turnover constraint.")
        result = optimize(None)
        print(">>> Retry (no turnover) returned success =", result.success)
        if not result.success:
            return {
                'status': 'error',
                'message': f"Optimization failed: {result.message}",
                'exposure_matrix_standardized': F.tolist(),
                'target_exposures_standardized': target.tolist(),
                'constraints': 'Turnover omitted on retry',
                'current_weights': current_weights.tolist()
            }

    # 4.11) On success, build the output
    optimized_weights = result.x
    achieved_standardized = (F.T @ optimized_weights)

    return {
        'status': 'success',
        'optimized_weights': [
            {'ticker': t, 'weight': float(round(wi, 6))}
            for t, wi in zip(tickers, optimized_weights)
        ],
        'target_exposures_raw': raw_target.tolist(),
        'target_exposures_standardized': target.tolist(),
        'achieved_exposures_standardized': dict(
            zip(factor_matrix.columns, achieved_standardized.round(6).tolist())
        )
    }
