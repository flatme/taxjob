import pandas as pd
from Sap_login import SapGui
this_dir = "phuocnguyen.xlsx"
xl = pd.ExcelFile(this_dir)
df_bienlai_unt = xl.parse("Sheet1",dypte=str)
df_bienlai_unt['TrangThai'] = 'FALSE'
df_bienlai_unt['KyThue'] = '032022'
df_bienlai_unt['PhuongXa'] = 'Phường Phước Nguyên'
df_bienlai_unt = df_bienlai_unt.fillna(0)
df_bienlai_unt['MaDTNT'] = df_bienlai_unt['MaDTNT'].astype(str)
df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype('int64')
df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype(str).str.zfill(7)
df_bienlai_unt['SoBL']= "CCT09B"+ df_bienlai_unt['KyHieu'] + df_bienlai_unt['SoBL']
#df_bienlai_unt['NgayBL'] = pd.to_datetime(df_bienlai_unt['NgayBL']).dt.strftime('%d.%m.%Y')
df_input = df_bienlai_unt[['MaDTNT','TenDTNT','ToThon','PhuongXa','SoBL','SoTien','NopCham','NgayBL','TrangThai','QuyenBL','KyThue','SoThua','SoTo']]
df_input.to_excel("phuocnguyen_final.xlsx")
df_sothunop = SapGui.getsothunop(df_input["MaDTNT"].tolist(),'71703015')
for row in df_input:
    



