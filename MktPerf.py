import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import os
import uuid

# Function to get historical data from Yahoo Finance
def get_historical_data(symbol, start_date, end_date):
    df = yf.download(symbol, start=start_date, end=end_date)
    return df

# Function to find the closest date
def get_closest_date(df, target_date):
    dates = pd.to_datetime(df.index)
    closest_date = dates[dates <= target_date].max()
    return closest_date.strftime('%Y-%m-%d') if not pd.isna(closest_date) else None

# Function to calculate annualized returns
def calculate_annualized_return(start_value, end_value, periods):
    return ((end_value / start_value) ** (1 / periods) - 1) * 100

# Function to calculate performance
def calculate_performance(df):
    performance = {}
    values = {}
    today = datetime.today()
    current_value = df['Close'].iloc[-1]

    periods = {
        '1 Month': today - timedelta(days=30),
        '3 Months': today - timedelta(days=90),
        '6 Months': today - timedelta(days=180),
        'YTD': datetime(today.year, 1, 1),
        '1 Year': today - timedelta(days=365),
        '3 Years': today - timedelta(days=3*365),
        '5 Years': today - timedelta(days=5*365),
    }

    for period, target_date in periods.items():
        closest_date = get_closest_date(df, target_date)
        if closest_date:
            historical_value = df.loc[closest_date, 'Close']
            values[f'{period} Value'] = historical_value
            if 'Years' in period:
                years = int(period.split()[0])
                performance[f'{years}Y Ann.'] = calculate_annualized_return(historical_value, current_value, years)
            else:
                performance[f'{period} Return'] = ((current_value - historical_value) / historical_value) * 100
        else:
            values[f'{period} Value'] = None
            performance[f'{period} Return'] = None

    values['Current Value'] = current_value
    return performance, values

# Main equity indices for MSCI Emerging Market Index
emerging_market_indices = {
    'China': ('000001.SS', 'Shanghai Composite'),
    'India': ('^BSESN', 'BSE SENSEX'),
    'Brazil': ('^BVSP', 'Bovespa'),
    'South Africa': ('^JN0U.JO', 'JSE All Share Index'),
    'Mexico': ('^MXX', 'IPC Mexico'),
    'Qatar*': ('QAT', 'Qatar ETF'),
    'Turkey': ('XU100.IS', 'BIST 100'),
    'Indonesia': ('^JKSE', 'Jakarta Composite Index'),
    'Malaysia': ('^KLSE', 'FTSE Bursa Malaysia KLCI'),
    'Philippines': ('^PSEI', 'PSEi Composite'),
    'Taiwan': ('^TWII', 'Taiwan Weighted'),
    'Thailand': ('^SET.BK', 'SET Index'),
    'Saudi Arabia': ('^TASI.SR', 'Tadawul All Share Index'),
    'Emerging Market': ('EEM', 'MSCI EM Index')
}

# Main equity indices for MSCI Developed Market Index
developed_market_indices = {
    'USA': ('^GSPC', 'S&P 500'),
    'UK': ('^FTSE', 'FTSE 100'),
    'Germany': ('^GDAXI', 'DAX'),
    'France': ('^FCHI', 'CAC 40'),
    'Japan': ('^N225', 'Nikkei 225'),
    'Canada': ('^GSPTSE', 'S&P/TSX Composite'),
    'Austria': ('^ATX', 'Austrian Traded Index'),
    'Belgium': ('^BFX', 'BEL 20'),
    'Finland': ('^OMXH25', 'OMX Helsinki 25'),
    'Ireland': ('^ISEQ', 'ISEQ Overall'),
    'Israel': ('^TA125', 'TA-125 Index'),
    'Italy': ('^FTMIB', 'FTSE MIB'),
    'Netherlands': ('^AEX', 'AEX Index'),
    'Norway': ('^OSEAX', 'Oslo All Share'),
    'Spain': ('^IBEX', 'IBEX 35'),
    'Sweden': ('^OMXS30', 'OMX Stockholm 30'),
    'Switzerland': ('^SSMI', 'SMI'),
    'Australia': ('^AXJO', 'S&P/ASX 200'),
    'New Zealand': ('^NZ50', 'NZX 50'),
    'Singapore': ('^STI', 'Straits Times Index'),
    'Developed Market': ('EFA', 'MSCI EAFE Index'),
    'All Country WI': ('ACWI', 'MSCI ACWI Index')
}

# Function to generate market performance
def generate_market_performance(indices, start_date, end_date):
    market_data = []
    for country, (symbol, index_name) in indices.items():
        print(f"Fetching data for {symbol} ({country}) from {start_date} to {end_date}")
        df = get_historical_data(symbol, start_date, end_date)
        if not df.empty:
            print(f"Data fetched for {symbol}: {df.head()}")
            performance, values = calculate_performance(df)
            print(f"Performance for {symbol}: {performance}")
            print(f"Values for {symbol}: {values}")
            row_values = {
                'Country': country,
                'Index Ticker': symbol,
                'Index Name': index_name,
                **values,
            }
            market_data.append(row_values)

            row_returns = {
                'Country': country,
                'Index Ticker': symbol,
                'Index Name': index_name,
                **performance,
            }
            market_data.append(row_returns)
        else:
            print(f"No data available for {symbol}")
    return market_data

def process_query_3(query):
    """
    Determines the type of market performance report based on the query.
    Returns 'full', 'emerging', or 'developed'.
    """
    query = query.lower()
    if "global" in query or "all country" in query or "full market" in query:
        return "full"
    elif "emerging" in query or "developing market" in query:
        return "emerging"
    elif "developed" in query or "advanced market" in query:
        return "developed"
    else:
        return None  # Unrecognized query type

def render_market_data_to_html(data, market_type):
    """Generate an HTML table from market data."""
    if not data:
        return f"<h1>No data available for {market_type}</h1>"

    html = f"<h1>{market_type} Market Data</h1>"
    html += "<table border='1' style='border-collapse:collapse;'>"
    # Add headers
    headers = data[0].keys()
    html += "<tr>" + "".join(f"<th>{header}</th>" for header in headers) + "</tr>"
    # Add rows
    for row in data:
        html += "<tr>" + "".join(f"<td>{value}</td>" for value in row.values()) + "</tr>"
    html += "</table>"
    return html


# Generate full market performance report
def generate_full_market_performance(start_date, end_date):
    print(f"Generating full market performance report for range: {start_date} to {end_date}")
    emerging_market_data = generate_market_performance(emerging_market_indices, start_date, end_date)
    developed_market_data = generate_market_performance(developed_market_indices, start_date, end_date)

    full_html = render_market_data_to_html(emerging_market_data, 'Emerging Market') + \
                render_market_data_to_html(developed_market_data, 'Developed Market')
    
    filename = f"Full_Market_Performance_{uuid.uuid4().hex}.html"
    print(f"Full market performance report generated: {filename}")
    return save_html_to_file(full_html, filename)


def generate_emerging_market_performance(start_date, end_date):
    print(f"Generating emerging market performance report for range: {start_date} to {end_date}")
    emerging_market_data = generate_market_performance(emerging_market_indices, start_date, end_date)

    full_html = render_market_data_to_html(emerging_market_data, 'Emerging Market')
    
    filename = f"Emerging_Market_Performance_{uuid.uuid4().hex}.html"
    print(f"Emerging market performance report generated: {filename}")
    return save_html_to_file(full_html, filename)


def generate_developed_market_performance(start_date, end_date):
    print(f"Generating developed market performance report for range: {start_date} to {end_date}")
    developed_market_data = generate_market_performance(developed_market_indices, start_date, end_date)

    full_html = render_market_data_to_html(developed_market_data, 'Developed Market')
    
    filename = f"Developed_Market_Performance_{uuid.uuid4().hex}.html"
    print(f"Developed market performance report generated: {filename}")
    return save_html_to_file(full_html, filename)

# Function to save HTML content to a file in the Stock Reports folder
def save_html_to_file(html_content, filename):
    reports_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Stock Reports')
    os.makedirs(reports_dir, exist_ok=True)  # Ensure the directory exists
    file_path = os.path.join(reports_dir, filename)
    
    with open(file_path, 'w') as f:
        f.write(html_content)
    
    return file_path

# Main execution guard
if __name__ == "__main__":
    start_date = (datetime.today() - timedelta(days=5*365)).strftime('%Y-%m-%d')
    end_date = datetime.today().strftime('%Y-%m-%d')
    print("Script executed directly. Use the functions to generate market performance reports.")
