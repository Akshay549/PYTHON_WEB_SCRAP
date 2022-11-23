import openpyxl
import requests
from bs4 import BeautifulSoup
wb=openpyxl.Workbook()
sheet=wb.active
sheet.title='top 250 countries'
sheet.append(['rank','movie name','year of release','imdb rating'])
print(wb.sheetnames)
website=requests.get('https://www.imdb.com/chart/top/')
website.raise_for_status()

movie=BeautifulSoup(website.text,'html.parser')
cinema=movie.find('tbody',class_="lister-list").find_all('tr')
try:
   for x in cinema:
    rank=x.find('td',class_="titleColumn").text.split()[0].strip('.')
    name=x.find('td',class_="titleColumn").a.text
    year=x.find('td',class_="titleColumn").span.text.strip('(()')
    rating=x.find('td',class_="ratingColumn imdbRating").strong.text
    sheet.append([rank,name,year,rating])
    print(rank,name,year,rating)
except Exception as e:
    print(e)
    
wb.save('C:\\Users\\AKSHAY AKHADE\\Desktop\\PROJECT FILES\\TOPMOVESONIMDB.xlsx')
