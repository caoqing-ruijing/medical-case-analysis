import pandas as pd
import sys

try:
    # 读取Excel文件
    df = pd.read_excel('END.xlsx')
    
    # 打印文件结构
    print('Excel文件结构:')
    print(df.head(3).to_string())
    
    # 打印列名
    print('\n列名:', list(df.columns))
    
    # 打印行数
    print('\n总行数:', len(df))
    
    # 打印每列的数据类型
    print('\n数据类型:')
    print(df.dtypes)
    
except Exception as e:
    print(f'错误: {str(e)}')