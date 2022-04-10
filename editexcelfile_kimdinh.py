import pandas as pd

this_dir = "kimdinh.xlsx"
xl = pd.ExcelFile(this_dir)
df_bienlai_unt = xl.parse("Sheet1",dypte=str)
df_bienlai_unt['TrangThai'] = 'FALSE'
df_bienlai_unt['KyThue'] = '032022'
df_bienlai_unt['PhuongXa'] = 'Phường Kim Dinh'
df_bienlai_unt = df_bienlai_unt.fillna(0)
df_bienlai_unt['MaDTNT'] = df_bienlai_unt['MaDTNT'].astype(str)
#df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype('int64')
#df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype(str).str.zfill(6)
#df_bienl#ai_unt['SoBL']= "CCT09B"+ df_bienlai_unt['KyHieu'] + df_bienlai_unt['SoBL']
#['NgayBL'] = pd.to_datetime(df_bienlai_unt['NgayBL']).dt.strftime('%d.%m.%Y')
df_bienlai_unt[['MaDTNT','TenDTNT','ToThon','PhuongXa','SoBL','SoTien','NopCham','NgayBL','TrangThai','QuyenBL','KyThue','SoThua','SoTo']].to_excel("kimdinh_final.xlsx")
