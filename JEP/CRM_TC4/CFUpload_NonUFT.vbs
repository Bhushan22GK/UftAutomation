set oshell = Createobject("WScript.shell")
oshell.run "cmd /k C:\\RIL_TTAF_REPLICA\\JEP\\CRM_TC4\\CRM_Login.vbs"
WScript.sleep 25000

'*************Variable declarations***************************************
order_type = 21
channel = 75
t_code = "se37"
Milestone_ID = "99"
status = "S"
dim CFNo(100)
'*************************************************************************


'************Fetching Index for CF Number*********************************
Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
xlobj.DisplayAlerts = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Datasheet\Data.xlsx")
set xlwks = xlwk.Worksheets("QUICK_INPUT")
start_index = XLWKS.cells(20,6)
OrderID_size = 0
while XLWKS.cells(start_index+OrderID_size,1) <> ""
	CFNo(OrderID_size+1) = XLWKS.cells(start_index+OrderID_size,1)
	OrderID_size = OrderID_size+1
wend																																																																																																					
'WIND UP
xlwk.sAVE
xlwk.cLOSE
SET xlwk = Nothing
SET xlobj = NOTHING
'*************************************************************************
'*************************************************************************


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

for iter = 1 to OrderID_size+1

	session.findById("wnd[0]/tbar[0]/okcd").text = t_code
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/ctxtRS38L-NAME").text = "ZCRM_ENT_ORD_STAT_UPDATE"
	session.findById("wnd[0]").sendVKey 8
	session.findById("wnd[0]/usr/lbl[34,9]").setFocus
	session.findById("wnd[0]/usr/lbl[34,9]").caretPosition = 0
	session.findById("wnd[0]").sendVKey 2
	session.findById("wnd[0]/usr/txt[12,3]").text = CFNo(iter)
	session.findById("wnd[0]/usr/txt[53,3]").text = order_type
	session.findById("wnd[0]/usr/txt[116,3]").text = channel
	session.findById("wnd[0]/usr/lbl[56,3]").setFocus
	session.findById("wnd[0]/usr/lbl[56,3]").caretPosition = 0
	session.findById("wnd[0]").sendVKey 2
	session.findById("wnd[0]/tbar[1]/btn[19]").press
	session.findById("wnd[1]/usr/txt[43,5]").text = Milestone_ID
	session.findById("wnd[1]/usr/txt[43,6]").text = status
	session.findById("wnd[1]/usr/txt[43,9]").text = "20180404121212"
	session.findById("wnd[1]/usr/txt[43,9]").setFocus
	session.findById("wnd[1]/usr/txt[43,9]").caretPosition = 14
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press



next

	session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
	session.findById("wnd[0]").sendVKey 0

