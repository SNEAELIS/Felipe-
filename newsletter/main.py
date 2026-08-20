from newslatter import fetch_24h_news
from filter import process_and_group_articles
from email_sender import send_daily_briefing

def run_pipeline():
    target_politician = "Ciro Nogueira"
    recipient = "felipe.rsouza@esporte.gov.br"

    print(f"1. Ingesting news for {target_politician}...")
    raw_articles = fetch_24h_news(query=target_politician, max_results=20)

    print("2. Classifying and grouping articles...")
    grouped_news = process_and_group_articles(raw_articles)
    print("3. Rendering email...")
    send_daily_briefing(
        politician_name=target_politician,
        grouped_news=grouped_news,
        recipient_email=recipient,
    )
    print("Done!")

if __name__ == "__main__":
    run_pipeline()
