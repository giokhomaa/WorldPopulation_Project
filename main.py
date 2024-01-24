import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook  # Workbook ფუნქციით შეგვიძლია შევქმნათ, შევცვალოთ, წავშალოთ ექსელში ფანჯრები
from openpyxl.styles import *
from openpyxl.worksheet.table import Table, TableStyleInfo

# openpyxl-ით შეგვიძლია შვცვალოთ სტილები თავისი ფუნქციებით

records = []

r = requests.get('https://www.worldometers.info/world-population/population-by-country/')
#print(r)

c = r.text
#print(c) # მოაქვს მთლიანი კოდი

soup = BeautifulSoup(c, "html.parser")

data = soup.find('tbody') # მოაქვს მთლიანი ცხრილი

rows = data.find_all('tr')

#print(rows[0])  # მოაქვს პირველი სტრიქონი

#columns = rows[0].find_all('td')  # სტრიქონში ანაწევრებს სვეტებს td ცვლადით
# print(columns[1].text)  # მოაქვს პირველი სტრიქონის მეორე სვეტი ანუ ქვეყნის დასახელება
# print(columns[2].text)
# print(columns[3].text)  # .text გარდაქმნის კოდს ტექსტად

# for index, row in enumerate(rows, 1):  # index ცვლადია რომელიც ინომრება enumerate ფუნქციით, ანუ ყველა ქვეყანა დაინომრება და დაიწყება 1-ით
#     columns = row.find_all('td')
#     print(f'{index}) {columns[1].text} - {columns[2].text} - {columns[3].text}')  # პირველ რიგში იპრინტება ნომერაცია, შემდეგ მოდის მეორე, მესამე და მეოთხე
#     # სვეტები ყველა სტრიქონიდან, ანუ ქვეყნის დასახელება, მოსახლეობის რაოდენობა და წლიური ზრდა/კლება

for index, row in enumerate(rows, 1):
    columns = row.find_all('td')

    country = columns[1].text
    pop = int(columns[2].text.replace(',', ''))
    percentage = columns[3].text  # გავაკეთეთ ცვლადები სათითაოდ

    item = [index, country, pop, percentage]  # გავაკეთეთ სიები
    records.append(item)  # სიები დავამატეთ მთავარ სიაში


records.insert(0, ['N', 'Countries', 'Population', 'Percentage'])  # მთავარ სიას დავამატეთ პირველი სტრიქონი 4 სვეტით, რომლებიც იქნება დანარჩენების სათაურები
#print(records)

# ვიწყებთ ექსელის ფაილთან მუშაობას

workbook = Workbook()

filename = 'population.xlsx'  # შევქმენით ექსელის ფაილი
workbook.save(filename)

sheet = workbook['Sheet']  # ვასწავლეთ ფანჯრის სახელი
sheet.title = 'Population'  # შევუცვალეთ სახელი
sheet = workbook['Population']

for info in records:
    sheet.append(info)  # sheet ფანჯარას ექსელში დავამატეთ records მონაცემები სათითაოდ info ცვლადით

table = Table(displayName='Population_Data', ref='A1:D236')
style = TableStyleInfo(name='TableStyleMedium8', showRowStripes=True, showColumnStripes=True)

table.tableStyleInfo = style
sheet.add_table(table)

# font = Font(color = '00FF0000', bold = True, italic = True)
#
# for cell_no in range(2, 237):
#     if sheet[f'C{cell_no}'].value < 5000000:
#         sheet[f'C{cell_no}'].font = font

workbook.save(filename)
workbook.close()



























