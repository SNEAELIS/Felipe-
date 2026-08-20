import urllib.parse

import feedparser

from typing import List, Dict, Any



def fetch_24h_news(query: str, max_results: int = 15) -> List[Dict[str, str]]:
    """
    Fetches news headlines and URLs published in the last 24 hours for a given query
    targeting Brazilian news outlets via Google News RSS.

    :param query: Name of politician or target topic (e.g., "Lula", "Tarcísio de Freitas").
    :param max_results: Maximum number of articles to return.
    :return: List of dicts containing 'title', 'link', 'published', and 'source'.
    """
    # URL encode query parameters safely
    encoded_query = urllib.parse.quote(f'"{query}" when:1d')
    
    # Brazilian Portuguese setup parameters for Google News RSS
    rss_url = (
        f"https://news.google.com/rss/search?"
        f"q={encoded_query}&hl=pt-BR&gl=BR&ceid=BR:pt-419"
    )

    feed = feedparser.parse(rss_url)
    articles = []

    if not hasattr(feed, "entries") or not feed.entries:
        return articles

    for entry in feed.entries[:max_results]:
        # Clean title splitting: Google News titles usually end with " - Source Name"
        raw_title = entry.get("title", "")
        source_name = entry.get("source", {}).get("title", "Unknown Source")
        
        articles.append({
            "title": raw_title,
            "link": entry.get("link", ""),
            "published": entry.get("published", ""),
            "source": source_name
        })

    return articles


if __name__ == "__main__":
    # Internal module testing
    target_politician = "Ciro Nogueira"
    print(f"--- Searching news from the last 24h for: {target_politician} ---")
    
    results = fetch_24h_news(query=target_politician, max_results=5)
    
    for idx, item in enumerate(results, 1):
        print(f"\n[{idx}] {item['title']}")
        print(f"    Source: {item['source']}")
        print(f"    Link:   {item['link']}")
        print(f"    Date:   {item['published']}")