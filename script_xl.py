import os.path
from pandas.io.excel import ExcelWriter
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from goog_exl import data
from string import ascii_uppercase

alt_rout = './output.xlsx'


def cell_let(b):
    a = list(ascii_uppercase)
    if b < 27:
        return (a[b % 26])
    elif b < 26 ** 2 + 26:
        return (a[b // 26 - 1] + a[b % 26])
    else:
        return (a[int(b / 26) // 26 - 1]) + (a[(b // 26) % 26 - 1]) + (a[b % 26])


if os.path.exists(alt_rout):
    today = datetime.today().strftime(('%d.%m.%Y'))
    main_data = data()
    df = pd.ExcelFile(alt_rout)
    df_1 = df.parse('test')
    wb = load_workbook(alt_rout)
    ws_1 = wb['1']
    ws_2 = wb['2']
    ws_3 = wb['test']
    len_test = len(ws_1['B']) - 1
    len_main = len(main_data[2])
    cell_num = cell_let(len(ws_1[1]))
    ws_1[cell_num + '1'] = today
    ws_2[cell_num + '1'] = today
    a = df_1['Месяц'].tolist()
    b = df_1['Дата'].tolist()
    if len_main > len_test:
        for i in range(len_test, len_main):
            ws_1[f'A{i + 2}'] = main_data[0][i]
            ws_1[f'B{i + 2}'] = main_data[3][i]
            ws_1[f'{cell_num}{i + 2}'] = main_data[2][i]
            ws_2[f'A{i + 2}'] = main_data[0][i]
            ws_2[f'B{i + 2}'] = main_data[3][i]
            ws_2[f'{cell_num}{i + 2}'] = main_data[1][i]
            ws_3[f'A{i + 2}'] = main_data[3][i]
            ws_3[f'B{i + 2}'] = main_data[1][i]
            ws_3[f'C{i + 2}'] = main_data[2][i]

    for i in range(len_test):
        if a[i] != main_data[2][i]:
            ws_3[f'C{i + 2}'] = main_data[2][i]
            ws_1[f'{cell_num}{i + 2}'] = main_data[2][i]
        wb.save(alt_rout)
    for i in range(len_test):
        if b[i] != main_data[1][i]:
            ws_3[f'B{i + 2}'] = main_data[1][i]
            ws_2[f'{cell_num}{i + 2}'] = main_data[1][i]
        else:
            pass
        wb.save(alt_rout)
    wb.close()

else:
    main_data = data()
    today = datetime.today().strftime(('%d.%m.%Y'))
    count = len(main_data[3])
    df = pd.DataFrame(
        {'ФИО/Название\nподрядчика': main_data[0],
         'Уникальный номер размещения': main_data[3],
         f'{today}': main_data[2],

         })

    df2 = pd.DataFrame(
        {'ФИО/Название\nподрядчика': main_data[0],
         'Уникальный номер размещения': main_data[3],
         f'{today}': main_data[1],
         })
    df_test = pd.DataFrame(
        {
            'Уникальный номер размещения': main_data[3],
            'Дата': main_data[1],
            'Месяц': main_data[2],
        })

    with ExcelWriter(alt_rout) as writer:
        df.to_excel(writer, sheet_name='1', index=False)
        df2.to_excel(writer, sheet_name='2', index=False)
        df_test.to_excel(writer, sheet_name='test', index=False)
    print('ok')
