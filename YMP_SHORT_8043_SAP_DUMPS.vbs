If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

StartTime = timer

pocip_location = "C:\_Private Data - NO BACKUP\POCIP\Dumps\8043"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radRB_PCNF").select
session.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").text = "31.12.2021"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").caretPosition = 10
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_PCNF_2021.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.01.2019"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "30.06.2019"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2019_H1.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.07.2019"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "31.12.2019"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2019_H2.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.01.2020"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "30.06.2020"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2020_H1.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.07.2020"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "31.12.2020"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2020_H2.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.01.2021"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "31.03.2021"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2021_Q1.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.04.2021"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "30.06.2021"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2021_Q2.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_SHORT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_GETRI-LOW").text = "01.06.2021"
session.findById("wnd[0]/usr/ctxtS_GETRI-HIGH").text = "30.09.2021"
session.findById("wnd[0]/usr/ctxtS_WERKS").text = "8043"
session.findById("wnd[0]/usr/ctxtS_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = pocip_location
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_SHORT_8043_CNF_2021_Q3.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press

EndTime = timer

DeltaTime = FormatNumber( EndTime - StartTime , 2 ) & " seconds"

MsgBox ( "POCIP 8043 script completed in " & DeltaTime )
