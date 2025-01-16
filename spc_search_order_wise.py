import pandas as pd
import requests as rq
from bs4 import BeautifulSoup as bs
from typing import Dict

url = 'http://192.168.13.253/mymun/finishstore/FFS_Orderwise_RunningStatus_show.php?orderno='
result = pd.DataFrame(columns=['Order', 'F/F Order Qty', 'F/F Delv Qty', 'F/F Delv Bal Qty'])

orders = pd.read_excel('orders.xlsx')
for index, row in orders.iterrows():
    order = str(row['FRS No.'])
    print("Process Completed", int(((index + 1) / len(orders)) * 100), "% -----", order, "-----")
    response = rq.get(url + order)
    html_content = bs(response.content, 'html.parser')
    rows = []
    for row in html_content.find_all('tr'):
        row_data = [cell.get_text(strip=True) for cell in row.find_all('td')]
        rows.append(row_data)
    my_dict : Dict[str, Dict[str, list[int, int, int]]] = {}
    for row in rows:
        # if len(row) == 18:
        #     my_dict[row[0][row[2]]] += [float(row[6]), float(row[7]), float(row[8])] 
        if len(row) == 13:
            result.loc[index] = [order, float(row[2]), float(row[3]), float(row[4])]
            break
# print(my_dict)
result.to_excel('Output.xlsx', index=False)