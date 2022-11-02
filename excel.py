import openpyxl
from openpyxl import load_workbook
wb = load_workbook('C:\\Users\\zzhuetan\\OneDrive - PERNOD RICARD\\Desktop\\user list.xlsx')
ws = wb.active

rows_data = list(ws.rows)
# 获取表单的表头信息(第一行)，也就是列表的第一个元素
titles = [title.value for title in rows_data.pop(0)]
# print(titles)

all_row_dict = []
# 遍历出除了第一行的其他行
for a_row in rows_data:
    the_row_data = [cell.value for cell in a_row]
    # 将表头和该条数据内容，打包成一个字典
    row_dict = dict(zip(titles, the_row_data))
    # print(row_dict)
    all_row_dict.append(row_dict)

print(all_row_dict)
# for i in range(10):
# 	if all_row_dict[i]['Shanghai'] == 'NO':
# 		print(all_row_dict[i]['User First Name'])
# 		print(all_row_dict[i]['Email'])
