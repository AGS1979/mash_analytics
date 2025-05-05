import os
import openai
import requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

# Read your OpenAI API key from the environment
openai.api_key = os.environ["OPENAI_API_KEY"]

# Function to fetch news from Google News (past 7 days)
def fetch_google_news(query):
    """
    Fetches the latest news articles for the given query from Google News for the past 7 days.
    """
    search_url = f"https://news.google.com/rss/search?q={query.replace(' ', '+')}&hl=en-US&gl=US&ceid=US:en"
    response = requests.get(search_url)

    if response.status_code != 200:
        print(f"❌ Error fetching Google News RSS: {response.status_code}")
        return []

    soup = BeautifulSoup(response.content, "xml")
    items = soup.find_all("item")

    articles = []
    cutoff_date = datetime.today() - timedelta(days=7)

    for item in items:
        title = item.title.text
        link = item.link.text
        pub_date = datetime.strptime(item.pubDate.text, "%a, %d %b %Y %H:%M:%S %Z")

        # Only include articles from the last 7 days
        if pub_date >= cutoff_date:
            articles.append({"title": title, "url": link, "published_at": pub_date})

    print(f"✅ Found {len(articles)} articles from Google News for: {query}")
    return articles[:5]  # Return top 5 articles

# Function to summarize news using GPT
def summarize_news_articles(articles):
    """
    Uses GPT to summarize the fetched news articles.
    """
    summaries = []
    for article in articles:
        prompt = f"Summarize the following news article: {article['title']}"
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=150
            )
            summary = response["choices"][0]["message"]["content"]
            summaries.append({"title": article["title"], "summary": summary, "url": article["url"]})
        except Exception as e:
            print(f"Error in GPT summarization: {e}")
            summaries.append({"title": article["title"], "summary": "Summary unavailable", "url": article["url"]})

    return summaries

# Main function to process query
def get_latest_news_summary(query):
    """
    Given a query, fetch the latest Google News articles and summarize them.
    """
    news_articles = fetch_google_news(query)

    if not news_articles:
        return {"error": "No recent news found for this query."}

    summarized_news = summarize_news_articles(news_articles)

    print(f"📄 Summarized {len(summarized_news)} articles.")  # Debugging

    return summarized_news
