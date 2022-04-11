import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import pandas as pd 
from pandasql import sqldf
xl_input = pd.ExcelFile('input.xlsx')
xl_compare = pd.ExcelFile('compare.xlsx')
df_input = xl_input.parse('Sheet1')
df_compare = xl_compare.parse('Sheet1')

print('Chua Co To Khai')
df_delta=df_input[df_input['col1'].apply(lambda x: x not in df_compare['col1'].values)]
print(df_delta)
df_input=df_input.append(df_delta)
df_input =df_input.drop_duplicates(subset=['col1','SoBL'],keep=False)
m = df_input.merge(df_compare, on='col1', how='outer', suffixes=['', '_'], indicator=True).loc[lambda x : x['_merge']=='both'].drop_duplicates('col1',keep=False)
print(m)
df_input=df_input.append(m)
df_input =df_input.drop_duplicates(subset=['col1','SoBL'],keep=False)
pysqldf = lambda q: sqldf(q, globals())
print('Nhieu Thua Dat')
df_result =pysqldf("select distinct df_input.col1,df_input.SoBL,df_input.col2,df_compare.col2,CAST(df_input.col2 as REAL) / CAST(df_compare.col3 as REAL) as div  from df_input LEFT  JOIN df_compare on df_input.col1 = df_compare.col1")
df_result['sub'] = round(df_result['div']) - df_result['div']
mask = abs(df_result['sub']).lt(0.01)
df_result.loc[ mask,'div']= round(df_result['div'])
df_result.loc[mask,'sub']=0

result_df =df_result[(df_result['sub']==0)].drop_duplicates(subset=['col1','SoBL'],keep='first')


print(result_df)
df_result=df_result.append(result_df)
df_result =df_result.drop_duplicates(subset=['col1','SoBL'],keep=False)
conlai_df = df_result[(df_result['div']!=1 )]
print('Con Lai')
print(conlai_df.drop_duplicates('col1',keep=False))

