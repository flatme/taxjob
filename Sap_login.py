import win32com.client
import sys
import subprocess
import pandas as pd
import time
from pathlib import Path

class SapGui():
    def __init__(self):
        # Hàm chạy TMS, gán tham số kết nối
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
       
        subprocess.Popen(self.path)
        time.sleep(5)        
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto)== win32com.client.CDispatch:
            return            
        application = self.SapGuiAuto.GetScriptingEngine       
        self.connection = application.OpenConnection("TMS", True)
        time.sleep(3)

        self.session = self.connection.Children(0)
        self.sapLogin()
    def sapLogin(self):
       # Hàm Login TMS
        try:
            #THE CLIENT
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "500"
            #USERNAME
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "ntcuong.brv"
            #PASSWORD
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Brv@2021"
            #LANGUAGE
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = ""
            #ENTER
            self.session.findById("wnd[0]").sendVKey(0)            
        except:
            print(sys.exc_info()[0])
    def tongquanNNT(self):
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00003"
        self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00003")
        self.session.findById("wnd[0]").sendVKey(4)
        self.session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[3,24]").text = "3500563736"
        self.session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[3,24]").setFocus()
        self.session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[3,24]").caretPosition = 10
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/lbl[8,3]").setFocus()
        self.session.findById("wnd[1]/usr/lbl[8,3]").caretPosition = 60
        self.session.findById("wnd[1]").sendVKey(2) 
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/tabsDATA_DISP/tabpDATA_DISP_FC1/ssubDATA_DISP_SCA:RFMCA_COV:0202/cntlRFMCA_COV_0100_CONT5/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        self.session.findById("wnd[0]/usr/tabsDATA_DISP/tabpDATA_DISP_FC1/ssubDATA_DISP_SCA:RFMCA_COV:0202/cntlRFMCA_COV_0100_CONT5/shellcont/shell").selectContextMenuItem("&LOAD") 
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        self.session.findById("wnd[0]/usr/tabsDATA_DISP/tabpDATA_DISP_FC1/ssubDATA_DISP_SCA:RFMCA_COV:0202/cntlRFMCA_COV_0100_CONT5/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        self.session.findById("wnd[0]/usr/tabsDATA_DISP/tabpDATA_DISP_FC1/ssubDATA_DISP_SCA:RFMCA_COV:0202/cntlRFMCA_COV_0100_CONT5/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/tabsDATA_DISP/tabpDATA_DISP_FC1/ssubDATA_DISP_SCA:RFMCA_COV:0202/cntlRFMCA_COV_0100_CONT5/shellcont/shell").selectContextMenuItem("&PC")
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
    def KetxuatSoThue(self):      
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZTC_KETXUAT_PNN"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtS_PERSL-LOW").text = "17CN"
        self.session.findById("wnd[0]/usr/ctxtS_PERSL-HIGH").text = "21CN"
        self.session.findById("wnd[0]/usr/ctxtS_TIN-LOW").text = "8410669166"
        self.session.findById("wnd[0]/usr/ctxtS_TIN-LOW").setFocus()
        self.session.findById("wnd[0]/usr/ctxtS_TIN-LOW").caretPosition = 10
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCUS_0101/shellcont/shell").pressToolbarContextButton("&MB_VARIANT") 
        self.session.findById("wnd[0]/usr/cntlCUS_0101/shellcont/shell").selectContextMenuItem("&LOAD")
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()        
        self.session.findById("wnd[0]/usr/cntlCUS_0101/shellcont/shell").pressToolbarContextButton("&MB_EXPORT") 
        self.session.findById("wnd[0]/usr/cntlCUS_0101/shellcont/shell").selectContextMenuItem("&PC") 
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
    def danhsachchungtu(self,CQT):
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZTC_RPT_PAYMENT"
        self.session.findById("wnd[0]").sendVKey(0)   
        self.session.findById("wnd[0]/usr/ctxtSP$00006-LOW").text = "7703"
        self.session.findById("wnd[0]/usr/ctxtSP$00006-LOW").setFocus()        
        self.session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/ctxtSP$00011-LOW").text = "01012022"
        self.session.findById("wnd[0]/usr/ctxtSP$00011-HIGH").text = "26012022"
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&LOAD")
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")     
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = 'D:\\Resources\\'
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = CQT+"_chungtu_tms.xlsx"
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()                          
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()     
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press() 
    def danhbaNNT(self,CQT):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZTC_RPT_BP"
        self.session.findById("wnd[0]").sendVKey(0)    
        self.session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(2,"TEXT")  
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleclickCurrentCell()
        self.session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = CQT
        self.session.findById("wnd[0]/tbar[1]/btn[26]").press()
        _returnitem = self.session.findById("wnd[1]/usr/txtMESSTXT1").text
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        # dem so luong ntt   
        _returnitem = _returnitem.replace('Số lượng NNT là','')
        _returnitem = _returnitem.replace('.','')
        return _returnitem
    def getsothunop(self,list_nnt,_MaPhuongXa):
        df_sothunop = pd.DataFrame()
        #copy list_nnt value to clipboard

        #run SAP script copy result to clipboard paste clipboard to DataFrame
        return df_sothunop


    
        
