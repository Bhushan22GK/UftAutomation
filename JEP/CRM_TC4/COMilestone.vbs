Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Datasheet\Data.xlsx")
set xlwks = xlwk.Worksheets("Jep")
Rowcount = xlwks.UsedRange.Rows.Count

CFNo = xlwks.cells(Rowcount,9)

xlwk.Save
xlwk.Close
xlobj.Quit
Set xlws = Nothing
Set xlwk = Nothing
Set xlobj = Nothing



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

on error resume next
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
on error resume next
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus
on error resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "se37"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38L-NAME").text = "ZCRM_MILESTONE_UPDATE"
session.findById("wnd[0]/usr/ctxtRS38L-NAME").caretPosition = 21
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/lbl[34,9]").setFocus
session.findById("wnd[0]/usr/lbl[34,9]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/txt[39,3]").text = CFNo
session.findById("wnd[1]/usr/txt[39,5]").text = "21"
session.findById("wnd[1]/usr/txt[39,7]").text = "75"
session.findById("wnd[1]/usr/txt[39,12]").text = "99"
session.findById("wnd[1]/usr/txt[39,13]").text = "S"
session.findById("wnd[1]/usr/txt[39,20]").text = "20180404121212"
session.findById("wnd[1]/usr/txt[39,20]").setFocus
session.findById("wnd[1]/usr/txt[39,20]").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[34,17]").setFocus
session.findById("wnd[0]/usr/lbl[34,17]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

