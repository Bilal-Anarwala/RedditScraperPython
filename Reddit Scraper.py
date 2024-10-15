from operator import truediv
from pickle import FALSE
import praw
from openpyxl import Workbook

#You may retrieve this information using this guide https://www.geeksforgeeks.org/how-to-get-client_id-and-client_secret-for-python-reddit-api-registration/
reddit = praw.Reddit(
    client_id = "",
    client_secret = "",
    user_agent = "My first scraper",
    username = '',
    password = '',
)
print('''Welcome to my Reddit Scraper! This program will take a subreddit of your choosing and save the title
and the numebr of upvotes into an excel file''')

def dataFetchHot():
    done = True
    while done:
        try:
            subr = input("Enter a subreddit to scrape data on: ")
            numLimit = int(input("How many posts would you like to gather data on?: "))
            fileName = input("Please enter the name you would like to call the file: ")

            dataDict = {}
            for submission in reddit.subreddit(subr).hot(limit=numLimit):
                val = {submission.title: submission.score}
                dataDict.update(val)
            excelSave(dataDict, fileName)
            done = False 
        except:
            print("You did not enter a valid subreddit, please try again")

def dataFetchNew():
    done = True
    while done:
        try:
            subr = input("Enter a subreddit to scrape data on: ")
            numLimit = int(input("How many posts would you like to gather data on?: "))
            fileName = input("Please enter the name you would like to call the file: ")

            dataDict = {}
            for submission in reddit.subreddit(subr).new(limit=numLimit):
                val = {submission.title: submission.score}
                dataDict.update(val)
            excelSave(dataDict, fileName)
            done = False 
        except:
            print("You did not enter a valid subreddit, please try again")

def dataFetchRising():
    done = True
    while done:
        try:
            subr = input("Enter a subreddit to scrape data on: ")
            numLimit = int(input("How many posts would you like to gather data on?: "))
            fileName = input("Please enter the name you would like to call the file: ")

            dataDict = {}
            for submission in reddit.subreddit(subr).rising(limit=numLimit):
                val = {submission.title: submission.score}
                dataDict.update(val)
                
            excelSave(dataDict, fileName)
            done = False 
        except:
            print("You did not enter a valid subreddit, please try again")

def dataFetchTop():
    done = True
    while done:
        try:
            subr = input("Enter a subreddit to scrape data on: ")
            numLimit = int(input("How many posts would you like to gather data on?: "))
            fileName = input("Please enter the name you would like to call the file: ")

            dataDict = {}
            for submission in reddit.subreddit(subr).top(limit=numLimit):
                val = {submission.title: submission.score}
                dataDict.update(val)

            excelSave(dataDict, fileName)
            done = False 
        except:
            print("You did not enter a valid subreddit, please try again")

def excelSave(dataDict, file):
    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = "Submission Title"
    sheet["B1"] = "Number of Upvotes"

    for row, (title, upvote) in enumerate(dataDict.items(), start=2):
        sheet [f"A{row}"] = title
        sheet [f"B{row}"] = upvote

    workbook.save(f'{file}.xlsx')
    print(f"An excel sheet named {file}.xlsx containing data on your subreddit has been saved")



done = True
while done:
    type = input("What kind of posts would you like to see, hot, new, rising, or top?: ")
    if type.lower() == 'hot':
        done = False
        dataFetchHot()
    elif type.lower() == 'new':
        done = False
        dataFetchNew()
    elif type.lower() == 'rising':
        done = False
        dataFetchRising()
    elif type.lower() == 'top':
        done = False
        dataFetchTop()
    else:
        print("That was not a valid input, please try again")
