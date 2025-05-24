import numpy as np
import pandas as pd
import cvxpy as cp
import statsmodels.api as sm
import requests

FMP_API_KEY      = os.environ["FMP_API_KEY"]
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
        ret = get_price_history(t)
        if not ret.empty:
            returns[t] = ret
    return pd.DataFrame(returns).dropna()

# Simulate factor returns (replace with French library or custom data)
def get_factor_returns():
    np.random.seed(0)
    dates = pd.date_range(start='2019-01-01', periods=60, freq='M')
    return pd.DataFrame({
        'MKT': np.random.normal(0.01, 0.05, 60),
        'SMB': np.random.normal(0.005, 0.03, 60),
        'HML': np.random.normal(0.004, 0.025, 60),
    }, index=dates)

# Compute regression-based factor loadings
def compute_factor_loadings(stock_returns, factor_returns):
    loadings = {}
    for ticker in stock_returns.columns:
        Y = stock_returns[ticker].dropna()
        X = factor_returns.loc[Y.index]
        if len(Y) < 30:
            continue
        X = sm.add_constant(X)
        model = sm.OLS(Y, X).fit()
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
    tickers = [p['ticker'] for p in portfolio]
    current_weights = np.array([p['weight'] for p in portfolio])

    stock_returns = get_stock_returns(tickers)
    factor_returns = get_factor_returns()
    factor_matrix = compute_factor_loadings(stock_returns, factor_returns)

    factor_matrix = factor_matrix.loc[tickers]
    F = factor_matrix.values
    target = np.array([target_exposures[f] for f in factor_matrix.columns])

    # Optimization
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
