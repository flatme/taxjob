import pandas as pd

this_dir = "hoalong.xlsx"
xl = pd.ExcelFile(this_dir)
df_bienlai_unt = xl.parse("Sheet1",dypte=str)
df_bienlai_unt['STT'] = df_bienlai_unt['STT'].astype(str)
df_bienlai_unt.columns = df_bienlai_unt.columns.str.strip()
df_bienlai_unt = df_bienlai_unt[df_bienlai_unt["STT"].str.contains("Tổng quyển|Quyển số|TỔNG CỘNG") == False]
df_bienlai_unt.rename(columns = {'Mã số thuế':'MaDTNT', 'Họ và tên':'TenDTNT','Địa chỉ':'ToThon','Số biên lai':'SoBL','Ngày thu':'NgayBL','Số thửa':'SoThua','Số tờ':'SoTo','Thuế chậm nộp':'NopCham','Tổng thu':'SoTien'}, inplace = True)
df_bienlai_unt['TrangThai'] = 'FALSE'
df_bienlai_unt['KyThue'] = '032022'
df_bienlai_unt['PhuongXa'] = 'Xã Hòa Long'
df_bienlai_unt = df_bienlai_unt.fillna(0)
df_bienlai_unt['MaDTNT'] = df_bienlai_unt['MaDTNT'].astype(str)
df_bienlai_unt = df_bienlai_unt[df_bienlai_unt["MaDTNT"].str.contains("Xoá|Hủy|Xóa") == False]
df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype('int64')
df_bienlai_unt['SoBL'] = df_bienlai_unt['SoBL'].astype(str).str.zfill(7)
df_bienlai_unt['SoBL']= "CCT09B"+ df_bienlai_unt['KyHieu'] + df_bienlai_unt['SoBL']
df_bienlai_unt['NgayBL'] = pd.to_datetime(df_bienlai_unt['NgayBL']).dt.strftime('%d.%m.%Y')

df_bienlai_unt[['MaDTNT','TenDTNT','ToThon','PhuongXa','SoBL','SoTien','NopCham','NgayBL','TrangThai','QuyenBL','KyThue','SoThua','SoTo']].to_excel("Hoalong_final.xlsx")
