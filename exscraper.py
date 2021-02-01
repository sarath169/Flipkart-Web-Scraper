from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import datetime
laptops=[]
for i in range(1,5):
    print(i)
    source = requests.get("https://www.flipkart.com/search?q=laptops&as=on&as-show=on&otracker=AS_Query_HistoryAutoSuggest_1_3_na_na_na&otracker1=AS_Query_HistoryAutoSuggest_1_3_na_na_na&as-pos=1&as-type=HISTORY&suggestionId=laptops&requestId=9fc7ba87-bccb-4aca-a0c3-2984664a6f7a&as-backfill=on&page="+str(i)).text
    soup = BeautifulSoup(source,'lxml')
    #E2-pcE _1q8tSL

    for div in soup.find_all(class_='_1fQZEK'):
        # print(div1.prettify()
        name = div.find('div',class_="_4rR01T").text
        if div.find('div',class_='_3LWZlK'):
            ratings=div.find('div',class_='_3LWZlK').text
        else:
            ratings="NA"
        list_price = div.find('div',class_="_30jeq3 _1_WHN1").text
        if div.find('div',class_='_3I9_wc _27UcVY'):
            actual_price = div.find('div',class_='_3I9_wc _27UcVY').text
        else:
            actual_price=list_price
        date_time=datetime.datetime.now()
        print(name,ratings,list_price,actual_price)
        laptops.append([name,ratings,list_price,actual_price,date_time])
    print("------------------------")
for i in laptops:
    print(i)
wb=Workbook()
# grab the active worksheet
ws = wb.active

# Data can be assigned directly to cells
ws['A1'] = 'Name'
ws['B1'] = 'Ratings'
ws['C1'] = 'Listing Price'
ws['D1'] = 'Actual Price'
ws['E1'] = 'Date and Time'
# Rows can also be appended
for i in laptops:
    ws.append(i)
# Python types will automatically be converted
# Save the file
wb.save("flipkart_top96_laptops.xlsx")
