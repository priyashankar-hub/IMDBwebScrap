from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)   #gives ['Sheet']
sheet = excel.active      #it has one sheet,to make sure we are working on active sheet we use this.
sheet.title = "Top Rated Movies"  #to change sheet name
print(sheet.title)   #gives name as-    Top Rated Movies
sheet.append(["Movie Rank","Movie Name","Year of Release","IMDB Rating"])  # to get headings i.e,,column name

try:
    source = requests.get("https://www.imdb.com/chart/top/?ref_=nv_mv_250")
    source.raise_for_status()
    soup = BeautifulSoup(source.text,"html.parser")

    movies = soup.find("tbody",class_="lister-list")
    movies = soup.find("tbody", class_="lister-list").find_all("tr")

    for movie in movies:
        name = movie.find("td",class_="titleColumn").a.text
        rank = movie.find("td", class_="titleColumn").get_text(strip=True).split(".")[0]
        year = movie.find("td", class_="titleColumn").span.text.strip("()")
        rating = movie.find("td", class_="ratingColumn imdbRating").strong.text
        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])   # to load into excel

except Exception as e:
    print(e)

excel.save("IMDB Movie Ratings.xlsx")       #it will save and create xlsx file

