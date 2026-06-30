import requests
from openpyxl import Workbook

url = "https://dummyjson.com/products"

response = requests.get(url)

data = response.json()

products = data["products"]

workbook = Workbook()

worksheet = workbook.active

worksheet.title = "Products"

worksheet.append([
    "Title",
    "Description",
    "Category",
    "Price",
    "Rating"
])

for product in products:

    title = product["title"]
    description = product["description"]
    category = product["category"]
    price = product["price"]
    rating = product["rating"]

    worksheet.append([
        title,
        description,
        category,
        price,
        rating
    ])

workbook.save("output/products.xlsx")

print("Done ✅")