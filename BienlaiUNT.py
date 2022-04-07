import os 
import sys
import sqlite3
class Bienlai:
   def __init__(self,MaDTNT=None,TenDTNT=None,ToThon=None,PhuongXa=None,SoBL=None,SoTien=None,NopCham=None,NgayBL=None,TrangThai=None,Ky=None,QuyenBL=None):
   # Instance Variable    
        self.MaDTNT = MaDTNT
        self.TenDTNT = TenDTNT
        self.ToThon = ToThon
        self.PhuongXa = PhuongXa
        self.SoBL = SoBL
        self.SoTien = SoTien
        self.NopCham = NopCham
        self.NgayBL = NgayBL
        self.TrangThai = TrangThai
        self.Ky = Ky
        self.QuyenBL = QuyenBL
   def chamnhanh(_maDTNT,_soTien):
          with sqlite3.connect('bienlaiunt.db') as conn:
           conn.execute('SELECT * from table_unt')
          print('t')
   def cham_tothon():
          print('t')
   def cham_matam():
          print('t')
