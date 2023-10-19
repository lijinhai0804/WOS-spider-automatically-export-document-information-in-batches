import pandas as pd
import xlrd
import os
import glob
# 要合并的文件所处路径
path=r'D:\test'
# 创建一个空的列表，用来存储每个Excel文件的DataFrame
df_list = []
# 表示对glob.glob函数返回的列表中的每个元素执行一次循环体（即冒号后面缩进的部分）。
# glob.glob函数接受一个字符串作为参数，表示要匹配的文件路径模式。
# 这里的参数是path + ‘/*.xlsx’，表示在path指定的文件夹中查找所有以.xlsx结尾的文件
for file in glob.glob(path + '/*.xls'):
    # 用pandas.read_excel函数读取每个Excel文件的内容，并返回一个DataFrame对象
    df = pd.read_excel(file)
    # 将这个DataFrame对象添加到df_list列表中
    df_list.append(df)
    # 用pandas.concat函数将df_list列表中的所有DataFrame对象拼接成一个大的DataFrame对象
merged_df = pd.concat(df_list, ignore_index=True)
# 用pandas.to_excel函数将这个大的DataFrame对象写入到新建的Excel文件中,文件保存在代码运行的文件夹中
merged_df.to_excel('merge_result.xlsx', sheet_name='OriginalData', index=False)
print('done')