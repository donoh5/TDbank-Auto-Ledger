import pandas as pd
from openpyxl import load_workbook

filename = 'accountactivity.csv'
main_df = pd.read_csv(filename, names=['date', 'name', 'in', 'out', 'total'])
main_df[['month','day', 'year']] = main_df.date.str.split("/",expand=True,)
data_date = pd.DataFrame(list(main_df.date.str.split("/")), columns=['month', 'day', 'year'])
main_df['total'] = ["=E%d-C%d+D%d" % (i+1, i+2, i+2) for i in range(len(main_df['total']))]
main_df.fillna(0, inplace=True)
main_num = len(main_df['month'].drop_duplicates())
first_num = int(main_df['month'][0])
whole_list = []

for i in range(main_num):
    new = main_df['month'] == format(i+1, '02')
    new2 = main_df[new]
    list_new2 = new2.values.tolist()
    whole_list.append(list_new2)

wb = load_workbook('MONEY.xlsx')
template = wb['template']
start_point = 0

for row in range(len(whole_list)):
    for cell in range(len(whole_list[row])):
        print(row)
        print(cell)
        if (whole_list[row][cell][7] + '-' + whole_list[row][cell][5]) in wb.sheetnames:
            origin_sh = wb[whole_list[row][cell][7] + '-' + whole_list[row][cell][5]]
            origin_sh.append(whole_list[row][cell])
        else:
            ws = wb.create_sheet(whole_list[row][cell][7] + '-' + whole_list[row][cell][5])
            ws.append(whole_list[row][cell])
wb.save('MONEY.xlsx')