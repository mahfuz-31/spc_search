import requests as rq
from bs4 import BeautifulSoup as bs
import pandas as pd

def convert_to_number(s):
    # Remove commas from the string
    s = s.replace(",", "")
    # Convert to float if there's a decimal point, otherwise to convert_to_number
    return float(s) if '.' in s else int(s)

print('-------------- Mahfuz Special Search ----------------\n')
ordersdf = pd.read_excel('orders.xlsx')

df = pd.DataFrame()
columns = ["Buyer", "Order", 'Style',	'UoF',	'F. Color',	'G. Color', 'Y. Type', 'F. Type', 'GSM', 'Dia',	'G/F Order With S.Note Qty', 'G/F S.Note Qty',	'Net Grey Receive Qty',	'G/F Rcv Balance Qty',	'F/F Order with S.Note Qty',	'F/F S.Note Qty',	'F/F Delv Qty',	'F/F Delv. Balance Qty',	'Replacement Delivery',	'F/F Excess Delv.Qty',	'Transfer To',	'Transfer From',	'Return Receive',	'Return Delivery',	'Dyeing Unit',	'Order Sheet Receive Date', 'Cut Plan Start Date',	'Cut Plan End Date']
df = df.reindex(columns=df.columns.tolist() + columns)

row_idx = 0
no_of_orders = len(ordersdf['FRS No.'])

for index, row in ordersdf.iterrows():
    order = row['FRS No.']
    print("Process completed", int((index + 1) * 100 / no_of_orders), "%  <---->", 'FRS:', order)
    # print("\nCalculating for order: ", order)
    url = 'http://192.168.13.253/mymun/Work%20Order/combineSearchResult.php?Welcome=7&GetOrderNO=' + str(order)
    response = rq.get(url)

    html_content = bs(response.content, 'html.parser')
    rows = []
    for row in html_content.find_all('tr'):
        row_data = [cell.get_text(strip=True) for cell in row.find_all('td')]
        rows.append(row_data)
        if len(row_data) == 1:
            print(row_data)