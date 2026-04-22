import csv
from collections import defaultdict

rows = []
with open(r'e:\BEE\Section1_Dataset\retail_sales_final.csv', newline='', encoding='utf-8') as f:
    for r in csv.DictReader(f):
        rows.append(r)

for r in rows:
    r['Sales']    = float(r['Sales'])
    r['Profit']   = float(r['Profit'])
    r['Quantity'] = int(r['Quantity'])
    r['Discount'] = float(r['Discount'])
    parts = r['Order Date'].split('-')
    r['Day'], r['Month'], r['Year'] = int(parts[0]), int(parts[1]), int(parts[2])

yr_sales = defaultdict(float)
yr_profit = defaultdict(float)
for r in rows:
    yr_sales[r['Year']]  += r['Sales']
    yr_profit[r['Year']] += r['Profit']

cats  = defaultdict(float)
regs  = defaultdict(float)
segs  = defaultdict(float)
subcats = defaultdict(float)
prods = defaultdict(float)
for r in rows:
    cats[r['Category']]       += r['Sales']
    regs[r['Region']]         += r['Sales']
    segs[r['Segment']]        += r['Sales']
    subcats[r['Sub-Category']] += r['Profit']
    prods[r['Product Name']]  += r['Profit']

total_s = sum(r['Sales']  for r in rows)
total_p = sum(r['Profit'] for r in rows)
total_q = sum(r['Quantity'] for r in rows)
total_o = len(rows)
margin  = total_p / total_s * 100

years   = sorted(yr_sales.keys())
yoy     = (yr_sales[2021] - yr_sales[2020]) / yr_sales[2020] * 100 if 2020 in yr_sales else 0

print(f"Total Rows   : {len(rows)}")
print(f"Years        : {years}")
print(f"Total Sales  : ${total_s:,.2f}")
print(f"Total Profit : ${total_p:,.2f}")
print(f"Margin       : {margin:.2f}%")
print(f"Total Orders : {total_o}")
print(f"Total Qty    : {total_q}")
print(f"Loss rows    : {sum(1 for r in rows if r['Profit']<0)}")
print()
for y in years:
    print(f"  {y} Sales : ${yr_sales[y]:,.2f}  Profit: ${yr_profit[y]:,.2f}")
print(f"  YoY Growth : +{yoy:.1f}%")
print()
print("By Category:")
for k,v in sorted(cats.items(), key=lambda x: -x[1]):
    print(f"  {k}: ${v:,.2f}")
print("By Region:")
for k,v in sorted(regs.items(), key=lambda x: -x[1]):
    print(f"  {k}: ${v:,.2f}")
print("By Segment:")
for k,v in sorted(segs.items(), key=lambda x: -x[1]):
    print(f"  {k}: ${v:,.2f}")
print("Top 5 Sub-Categories by Profit:")
for k,v in sorted(subcats.items(), key=lambda x: -x[1])[:5]:
    print(f"  {k}: ${v:,.2f}")
print("Bottom 3 Sub-Categories (Losses):")
for k,v in sorted(subcats.items(), key=lambda x: x[1])[:3]:
    print(f"  {k}: ${v:,.2f}")
