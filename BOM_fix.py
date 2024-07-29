import pandas as pd
from tkinter import filedialog

# 讀取 Excel 文件
file_path = filedialog.askopenfilename(title= '選擇要讀取的Excel文件: ', filetypes=[("Excel files", "*.xlsx; *.xls")])
if file_path:
    excel_data = pd.read_excel(file_path)

    # 顯示 Excel 表格的前幾行數據，以便查看數據結構
    print("原始數據：")
    print(excel_data.head())

    # 篩選出料件編號開頭為英文的項目，並刪除
    prefixs = ("Substitute")
    excel_data1 = excel_data[~excel_data['BOM.Primary/Substitute'].str.startswith(prefixs, na=False)]
    excel_data1 = excel_data1.dropna(subset=['BOM.Primary/Substitute'])
    print(excel_data1.head())

    excel_data2 = excel_data1.dropna(subset=['BOM.Ref Des'])
    print(excel_data2.head())

    # 將篩選後的數據保存到新的 Excel 文件中
    filtered_file_path = 'filtered_excel_file.xlsx'
    excel_data2.to_excel(filtered_file_path, index=False)
    print("已保存篩選後的數據到:", filtered_file_path)

    df_stacked = excel_data2['BOM.Ref Des'].str.split(',', expand=True).stack()

    df_stacked.index = df_stacked.index.droplevel(-1)

    df_stacked.to_csv('BOM_Fixed_file.txt', header=False, index=False)



