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

CurrentDateTime = iso8601DateTime(Now)
CurrentTimeRounded = timefx (Now)

Function iso8601DateTime(dt)
  s = datepart("yyyy",dt) & "-"
  s = s & RIGHT("0" & datepart("m",dt),2) & "-"
  s = s & RIGHT("0" & datepart("d",dt),2)
  iso8601DateTime = s
End Function

Function timefx(dt)
  s = RIGHT("0" & datepart("h",dt),2)
  s = s & RIGHT("0" & datepart("n",dt),2)
  roundby = 5
  x = round ( s / roundby ) * roundby
  timefx = x
End Function

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nYMP_S05"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 2
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\pattankarak\WILO\Wilo_PPC_India - PGI Reporting - PGI Reporting\YMP_S05\" & DatePart( "yyyy" , Now() ) & "-" & RIGHT( "0" & DatePart( "m", Now() ) , 2 )
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YMP_S05_" & CurrentDateTime & "-" & CurrentTimeRounded & ".XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
