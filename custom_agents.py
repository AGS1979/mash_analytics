from flask import Blueprint, render_template

custom_agents_bp = Blueprint('custom_agents', __name__, template_folder='templates')

AGENTS = [
    {
        "id": "qsg",
        "name": "Quantitative Signal Generator",
        "category": "Alpha Generation",
        "description": "Generates trading signals using machine learning based on historical prices and fundamentals.",
        "output": "Buy Signals: AAPL, MSFT | Sell Signals: GME, AMC"
    },
    {
        "id": "factor_optimizer",
        "name": "Factor Exposure Optimizer",
        "category": "Portfolio Optimization",
        "description": "Rebalances portfolio to maintain desired exposure to style factors like value and momentum.",
        "output": "Action: Reduce exposure to high beta, increase quality stocks (e.g., JNJ, PG)"
    },
    {
        "id": "macro_classifier",
        "name": "Macro Regime Classifier",
        "category": "Risk Management",
        "description": "Classifies the current macro regime (e.g., inflationary slowdown) and recommends tilts.",
        "output": "Current Regime: Inflationary slowdown. Tilt: Defensive sectors, short-duration bonds."
    }
]

@custom_agents_bp.route('/custom-agents')
def show_agents():
    return render_template('custom_agents.html', agents=AGENTS)
