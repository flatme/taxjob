import os
import sys
import pandas as pd
from BienlaiUNT import Bienlai
import sqlite3
conn = sqlite3.connect('database.db')
df = pd.read_excel('phuoctrung.xlsx',sheet_name='Sheet1',dtype='str')
df["SoBL"] = "CCT09BLK15" + df["SoBL"]
df['NgayBL']=pd.to_datetime(df['NgayBL']).dt.strftime('%d.%m.%y')
df.to_sql('Table_UNT',conn,if_exists='append')
df1 = pd.read_sql_query("SELECT * FROM Table_UNT", conn)
print(df1)