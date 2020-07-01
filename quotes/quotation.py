import openpyxl
import pprint


wb = openpyxl.load_workbook("quotes/E2004076A01見積書_税抜.xlsx")
ws = wb["template1"]


# 見積data取得
## excel data >> 二次元リスト化

## Header Left
quote_header = []

for row in ws.iter_rows(min_col=3, max_col=78, max_row=29):
    for cell in row:
        cell_value = cell.value
        if cell_value is not None:
            quote_header.append(cell_value)

## Quote Header Index番号確認

#for i, j in enumerate(quote_header):
#    print('{0}:{1}'.format(i,j))



## 見積詳細

quote_details = []


#for row in ws.iter_rows(min_col=2, min_row=32, max_col=78, max_row=35):
#    for cell in row:
#        cell_value = cell.value
#        if cell_value is not None:
#            quote_details.append(cell_value)

#print(quote_details)

# Quote details Index番号確認

#for i, j in enumerate(quote_details):
#    print('{0}:{1}'.format(i,j))



for i in ws.iter_rows(min_col=2, min_row=32, max_col=78, max_row=35):
    values_original = [cell.value for cell in i]
    values_original_filtered = [j for j in values_original if j is not None]
    quote_details.append(values_original_filtered)
    #print(values_ogriginal_filtered)
    # print('----------------------')
    # print(quote_details)
    # print('----------------------')

dt_col_list =[]

for k in quote_details:
    l = k[0]
    dt_col_list.append({"works" : l})

print(dt_col_list)
    



## 二次元リスト >> 一次元リスト化

#dt_col1 = [k for k in dt_col_list] # 作業名
#dt_col2 = [k[2] for k in quote_details] # 数量
#dt_col3 = [k[3] for k in quote_details] # 単位
#dt_col4 = [k[4] for k in quote_details] # 通貨
#dt_col5 = [k[5] for k in quote_details] # 単価
#dt_col6 = [k[6] for k in quote_details] # 金額
#dt_col7 = [k[7] for k in quote_details] # Ex.Rate
#dt_col8 = [k[8] for k in quote_details] # 金額(JPY)
#dt_col9 = [k[9] for k in quote_details] # 課税

#print(dt_col1[0])

## 見積総合計金額計算

#sum_gtotal = sum(dt_col8)
# print('----------------------')
#print("JPY " + str(sum_gtotal))

## Excel 閉じる
wb.close()

