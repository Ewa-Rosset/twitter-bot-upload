import tweepy
import openpyxl
from openpyxl import load_workbook
import time
import datetime

#Insert when quotes should be posted:
postingTimes = {"14:58":1, "14:59":1, "14:57":1}

# Authenticate to Twitter - Insert your own credentials
auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
auth.set_access_token(access_token, access_token_secret)

# Create API object
api = tweepy.API(auth)

# Load workbook
wb = load_workbook(filename='Quotes.xlsx')
quotes = wb['Quotes Database']


def selectQuote(rowNumber):
    currentRow = 'A' + str(rowNumber)
    quote = str(quotes[currentRow].value)
    return quote


def post(rowNumber):  
    api.update_status(selectQuote(rowNumber))
    print("Posted! Posted row number: ", rowNumber)
    
def main():

    #row 1 is the column title
    rowNumber = 2

    while True:
        currentDateTime = datetime.datetime.now()
        currentTime = currentDateTime.strftime("%H:%M")

        try:
            if postingTimes.get(currentTime) != None:
                post(rowNumber)
                rowNumber += 1
                time.sleep(60)  
        except tweepy.error.TweepError:
            rowNumber += 1

main()


    