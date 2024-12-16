import openpyxl

name_list = [
    "ALice",
    "Bob",
    "Carol",
    "Eve",
    "alex",
]

phone_num_list = [
    "555",
    "666",
    "777",
    "888",
    "999",
]

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Names and Phone nums"
sheet ["A1"] = "Names"
sheet ["B1"] = "Nums"

for name, number in zip(name_list, phone_num_list):
    sheet.append([name, number])
wb.save("ninja_phone_book.xlsx")