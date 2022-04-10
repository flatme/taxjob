import pandas as pd
import numpy as np
this_dir = "kimdinh2.xlsx"
xl = pd.ExcelFile(this_dir)
df_bienlai_unt = xl.parse("Sheet1",dypte=str)
mask = np.column_stack([df_bienlai_unt[col].astype(str).str.contains("Nguyễn|Bùi|Lê|Trần", na=False) for col in df_bienlai_unt])
mask2 = np.column_stack([df_bienlai_unt[col].astype(str).str.contains("8229", na=False) for col in df_bienlai_unt])
print(df_bienlai_unt.loc[mask.any(axis=1)])
