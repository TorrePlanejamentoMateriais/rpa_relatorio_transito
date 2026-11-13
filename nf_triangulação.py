session.findById("wnd[0]/tbar[0]/okcd").text = "ztpm028"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_DOCN_1-LOW").text = "1" 
session.findById("wnd[0]/usr/ctxtS_DOCN_1-HIGH").text = "9999999999"
session.findById("wnd[0]/usr/ctxtS_PSTDAT-LOW").text = "10.11.2026" #abrir um input para eu inserir a data de inicio 
session.findById("wnd[0]/usr/ctxtS_PSTDAT-HIGH").text = "11.11.2026"#abrir um input para eu inserir a data de fim
session.findById("wnd[0]/usr/ctxtS_PSTDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_PSTDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCC_ALV/shellcont/shell").setCurrentCell 10,"DOCNUM_ESC_EST"
session.findById("wnd[0]/usr/cntlCC_ALV/shellcont/shell").selectedRows = "10"
session.findById("wnd[0]/usr/cntlCC_ALV/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCC_ALV/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 8
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[6]/usr/ctxtDY_PATH").text = "C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Repositorio\NF Triangulação"#depois de salvar, abrir essa planilha e pegar o Nº documento 2
session.findById("wnd[6]/usr/ctxtDY_FILENAME").text = "ztpm028.xlsx" 
session.findById("wnd[3]/usr/ctxtDY_PATH").caretPosition = 90
session.findById("wnd[3]/tbar[0]/btn[0]").press
session.findById("wnd[2]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12


session.findById("wnd[0]/tbar[0]/okcd").text = "ztmm204"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "001"
session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = "10.11.2026"
session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").text = "13.11.2026"
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").caretPosition = 3
session.findById("wnd[0]/usr/btn%_S_BUKRS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "705"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "881"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 3
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]/usr/btn%_S_DOCNUM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 8,"N_REMESSA_REM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "8"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[3]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[3]").sendVKey 4
session.findById("wnd[4]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[4]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[4]").sendVKey 4
session.findById("wnd[5]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[5]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[5]").sendVKey 4
session.findById("wnd[6]/usr/ctxtDY_PATH").text = "C:\Users\F252727\OneDrive - Claro SA\Planilha de Atendimento\Repositorio\NF Triangulação"
session.findById("wnd[6]/usr/ctxtDY_FILENAME").text = "ztmm204.xlsx"
session.findById("wnd[6]/usr/ctxtDY_FILENAME").caretPosition = 7
session.findById("wnd[6]").sendVKey 0
session.findById("wnd[5]/tbar[0]/btn[11]").press
session.findById("wnd[4]/tbar[0]/btn[0]").press
session.findById("wnd[3]/tbar[0]/btn[0]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
