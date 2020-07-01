import openpyxl



wb = openpyxl.load_workbook("E2004076A01見積書_税抜.xlsx")
ws = wb["template1"]


# 見積data取得
## excel data >> 二次元リスト化

## Header Left
quote_header = []

for i in ws.iter_rows(min_col=3, min_row=3, max_col=3, max_row=4):
    values_original = [cell.value for cell in i]
    quote_header.append(values_original)

for i in ws.iter_rows(min_col=65, min_row=1, max_col=65, max_row=2):
    values_original = [cell.value for cell in i]
    quote_header.append(values_original)

for i in ws.iter_rows(min_col=59, min_row=6, max_col=59, max_row=7):
    values_original = [cell.value for cell in i]
    quote_header.append(values_original)

for i in ws.iter_rows(min_col=65, min_row=8, max_col=65, max_row=10):
    values_original = [cell.value for cell in i]
    quote_header.append(values_original)


print(quote_header)

## 見積詳細

quote_details = []

for i in ws.iter_rows(min_col=2, min_row=32, max_col=78, max_row=35):
    values_original = [cell.value for cell in i]
    values_ogriginal_filtered = [j for j in values_original if j is not None]
    quote_details.append(values_ogriginal_filtered)
    #print(values_ogriginal_filtered)
    # print('----------------------')
    # print(quote_details)
    # print('----------------------')

## 二次元リスト >> 一次元リスト化

dt_col1 = [k[0] for k in quote_details] # 作業名
dt_col2 = [k[2] for k in quote_details] # 数量
dt_col3 = [k[3] for k in quote_details] # 単位
dt_col4 = [k[4] for k in quote_details] # 通貨
dt_col5 = [k[5] for k in quote_details] # 単価
dt_col6 = [k[6] for k in quote_details] # 金額
dt_col7 = [k[7] for k in quote_details] # Ex.Rate
dt_col8 = [k[8] for k in quote_details] # 金額(JPY)
dt_col9 = [k[9] for k in quote_details] # 課税

# print(dt_col1)

## 見積総合計金額計算

sum_gtotal = sum(dt_col8)
# print('----------------------')
print("JPY " + str(sum_gtotal))

## Excel 閉じる
wb.close()

