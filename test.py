import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import pandas as pd 
from pandasql import sqldf
xl_input = pd.ExcelFile('input.xlsx')
xl_compare = pd.ExcelFile('compare.xlsx')
df_input = xl_input.parse('Sheet1')
df_compare = xl_compare.parse('Sheet1')

df_matam= df_input[~df_input['col1'].isin(df_compare['col1'])]
df_input = df_input.append(df_matam).drop_duplicates(keep=False)
print('Chua Co To Khai')
print(df_matam)
df_unique_maMST = df_compare.drop_duplicates('col1',keep=False)

df_compare = df_compare.append(df_unique_maMST).drop_duplicates(keep=False)
pysqldf = lambda q: sqldf(q, globals())
print('Thua Dat Duy Nhat')
print(pysqldf("select  df_input.col1,df_input.col2,df_unique_maMST.col2,df_input.col2 / df_unique_maMST.col3 as div from df_input INNER JOIN df_unique_maMST on df_input.col1 = df_unique_maMST.col1"))
print('Nhieu Thua Dat')
print(pysqldf("select  df_input.col1,df_input.col2,df_compare.col2,df_input.col2 / df_compare.col3 as div from df_input INNER JOIN df_compare on df_input.col1 = df_compare.col1"))