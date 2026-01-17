import pandas as pd

# 读取Excel文件
try:
    excel_file = pd.ExcelFile('sample_data.xlsx')
    print('子表名称:', excel_file.sheet_names)
    
    for sheet_name in excel_file.sheet_names:
        print('\n' + '-'*50)
        print(f'子表: {sheet_name}')
        df = excel_file.parse(sheet_name)
        print(f'数据行数: {df.shape[0]}')
        print(f'数据列数: {df.shape[1]}')
        print('\n数据预览:')
        print(df.head())
        if df.shape[0] > 5:
            print('\n... 更多数据 ...')
    print('\n' + '-'*50)
except Exception as e:
    print(f'读取文件时出错: {e}')
