import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


url = "https://www.scrapethissite.com/pages/simple/"

response = requests.get(url)

soup = BeautifulSoup(response.text, "html.parser")

countries = soup.find_all("div", class_="country")

workbook = Workbook()

worksheet = workbook.active

worksheet.title = "Countries"

worksheet.append([
    "Country",
    "Capital",
    "Population",
    "Area"
])

for country in countries:
    country_name = country.find(
    "h3",
    class_="country-name"
    ).text.strip()

    capital = country.find(
    "span",
    class_="country-capital"
    ).text.strip()

    population = country.find(
    "span",
    class_="country-population"
    ).text.strip()

    area = country.find(
    "span",
    class_="country-area"
).text.strip()

    worksheet.append([
    country_name,
    capital,
    population,
    area
    ])
workbook.save("output/countries.xlsx")