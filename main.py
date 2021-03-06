# apikey 0e1298a6f4534461b0484eb799f39fb6
import requests


class NewsFeed:

    base_url = "https://newsapi.org/v2/everything?"
    api_key = "0e1298a6f4534461b0484eb799f39fb6"

    def __init__(self, interest, from_date, to_date, language='en'):
        self.interest = interest
        self.from_date = from_date
        self.to_date = to_date
        self.language = language

    def get(self):
        url = f"{self.base_url}" \
              f"qInTitle={self.interest}&" \
              f"from={self.from_date}&" \
              f"to={self.to_date}&" \
              f"language={self.language}&" \
              f"apiKey={self.api_key}"

        response = requests.get(url)
        content = response.json()
        articles = content['articles']

        email_body = ''

        for article in articles:
            email_body = email_body + article['title'] + "\n" + article['url'] + "\n\n"

        return email_body


if __name__ == "__main__":
    news_feed = NewsFeed(interest='AI', from_date='2020-05-20', to_date='2020-05-21')
    news_feed.get()
