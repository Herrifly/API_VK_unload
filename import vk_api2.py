import vk_api
import json
import pickle
import pandas as pd
import openpyxl 
from config import token




session=vk_api.VkApi(token=token)
vk=session.get_api()




friends=vk.friends.get( order="name" ,fields=['sex','country','city','bdate','status'])
j = json.dumps(friends)

f = open ("file.json", "w", encoding="utf-8")
f.write(j)
  

file = open ("file.json", "r", encoding="utf-8")
data=json.load(file)


book = openpyxl.Workbook()
sheet = book.active

sheet['A1']="ID"
sheet['B1']="NAME"
sheet['C1']="SURNAME"
sheet['D1']="GENDER"
sheet['E1']="COUNTRY"
sheet['F1']="CITY"
sheet['G1']="Status"

row=2
for item in data["items"]:
    sheet[row][0].value=item["id"]
    sheet[row][1].value=item["first_name"]
    sheet[row][2].value=item["last_name"]
    sheet[row][3].value=item["sex"]
    if 'country' in item:
     sheet[row][4].value=item['country']['title']
     
    if 'city' in item:
     sheet[row][5].value=item['city']['title']
     
    sheet[row][6].value=item["status"]
    row+=1
   

book.save("my_book.xlsx")    
book.close()