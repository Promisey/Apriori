# -*- coding: utf-8 -*-
from efficient_apriori import apriori
import pandas as pd
import xlwt
import datetime
# import csv

df = pd.read_excel(r".\Apriori\订单明细.xlsx")

data = []
orders = list(set(df['订单号']))
order_num = len(orders)
for order in orders:
    data.append(list(df[df['订单号']==order]['商品名称']))

itemsets, rules = apriori(data, min_support=0.1,  min_confidence=0.8)
print(itemsets, '\n')
print(rules)

"""
# 输出CSV文件
file_name = r'.\Apriori\输出结果.csv'
out = open(file_name,'w', newline='', encoding='utf-8-sig')
csv_write = csv.writer(out, dialect='excel')
for item in itemsets.items():
    csv_write.writerow(item)
out.close()
"""

wk = xlwt.Workbook(encoding="utf-8")
sheet = wk.add_sheet("符合要求的商品组合")
sheet1 = wk.add_sheet("符合要求的关联关系")
row = 0
col = 0
for rule in rules:
    sheet1.write(row, col, str(rule))
    col += 1

col = 0
for keys, values in itemsets.items():
    sheet.write(row, col, keys)
    row += 1
    for key, value in values.items():
        support = '%.2f%%' % (value/order_num)
        sheet.write(row, col, str(key)+":"+str(support))
        row += 1
    col += 1
    row = 0
T = datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d%H%M")
wk.save("{}_{}.xls".format("商品关联结果", T))