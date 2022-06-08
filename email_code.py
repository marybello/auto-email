from main import NewsFeed
import yagmail
import pandas as pd
import datetime
import openpyxl
import time


data = pd.read_excel('emails.xlsx')
today = datetime.datetime.now().strftime('%Y-%m-%d')
yesterday_date = datetime.datetime.now() - datetime.timedelta(days=1)  # use timedelta to subtract days from dates
yesterday = yesterday_date.strftime('%Y-%m-%d')
email = yagmail.SMTP(user='learndatascience2@gmail.com',password='Maryamnafiu351')
firstnames = data['Firstname'].tolist()
emails = data['email'].tolist()

while True:
    if datetime.datetime.now().hour == 7 and datetime.datetime.now().minute == 15:
        for name, email1 in zip(firstnames, emails):

            newsfeed = NewsFeed(interest='AI', from_date=yesterday, to_date=today)
            email.send(to=email1,
                       subject=f"Your AI news for today!",
                       contents=f"Hi {name} \n See what's on about today. \n  {newsfeed.get()} \n"
                                f" This is your friend Maryam Bello. I am testing a new python app that can send "
                                f"automatic emails so if you get multiple emails from me tonight, I'm sorry..\n"
                                f"Kindly respond to this mail if you receive this email. \n Kind regards, \n"
                                f" Maryam"
                       )
    time.sleep(60)
