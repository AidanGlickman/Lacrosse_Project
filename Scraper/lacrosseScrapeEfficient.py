#!/usr/bin/env python3

import requests # This library handles web requests
from bs4 import BeautifulSoup # This library handles web parsing
from openpyxl import Workbook # This library handles writing to the excel format
import datetime # This library converts strings to dates, so that they will be the correct format in excel
import re # This library is for regular expressions

wb = Workbook() # Create an excel workbook
ws = wb.active # Get the active worksheet
ws.append(["Date", "Team A", "Score A", "Score B", "Team B"]) # Create the headers

url = "http://highschoolsports.nj.com/sprockets/game_search_results/?limit=25&season=3842&sport=300&page=" # This url, when combined with a page number, gets a barebones version of the scores table
newYear = False # Keep track of whether the year has been incremented (for seasons that go over a new year)
year = 2018 # This year is used to help create the date objects, has to be entered as the website does not mention the year for some reason

page = requests.get(url + "1") # GETs the first page of scores
soup = BeautifulSoup(page.content, "html.parser") # Parses the page

pageNum, maxPage = list(map(int, soup.find("span", attrs={"class": "page-of"}).text.strip("Page ").split(" of "))) # Finds the current page and max page, to be looped through to get every score

for i in range(pageNum, maxPage + 1): # Loop through each each numbered page
    print(f"PAGENUM {i}")
    page = requests.get(url + str(i)) # GET the page
    soup = BeautifulSoup(page.content, "html.parser") # Parse the page

    rows = soup.select("tr")[1:] # Find all of the rows of the table

    for x in rows: # Loop through the rows
        try: # this is a bit of a hacky way to check if the scores are valid, by trying to add them and if there is an error then just continuing
            children = x.findChildren("td", recursive=False) # Get all the elements of each row, which are the children of the row element
            date = re.sub(" |(\*)|(\#)", "", children[0].text).split("/") # Get the date, sanitize it and split it for use.
            if date[0] == "1" and not newYear: # If the month is january and the year hasn't yet been incremented, increment the date
                year += 1 # Increment the year
            date = datetime.date(year, int(date[0]), int(date[1])) # Create the date in the format year, month, day
            away, home = children[1].text.strip().split("@\n") # Get the away and home teams, which are in the 2nd column and split them on the "@" used on the site
            awayScore, homeScore = children[3].text.strip().split() # Get the scores of the two teams
            print(f"{date}\n{home}: {homeScore}\n{away}: {awayScore}\n\n")
            ws.append([date, home, int(homeScore), int(awayScore), away]) # Add all the fields to the excel sheet
        except ValueError: # Check if the is an error
            pass # If there is an error, then just do the next step of the loop

wb.save("girlslax.xlsx") # Save the file
