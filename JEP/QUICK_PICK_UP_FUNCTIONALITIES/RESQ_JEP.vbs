Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
xlobj.DisplayAlerts = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")

for i = 8 to 10
	for j = 1 to 5
		XLWKS.cells(i,j) = ""
	next
next

'STUFF WORK ORDER
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_MPLS"
	XLWKS.CELLS(8,2)	=	"RESQ_CycloBackEnd_JEP"
	XLWKS.CELLS(8,3)	=	"Fn_EFL_RESQ_CycloBackEnd_JEP(" &Chr(34) & Chr(34) &")"
	XLWKS.CELLS(8,4)	=	"SAPGuiOKCode"
	XLWKS.CELLS(8,5)	=	"wui_sso"
	XLWKS.CELLS(8,6)	=	"Linkname"
	XLWKS.CELLS(8,7)	=	"ZRESQ_CASECO - ZResQ Case Coordinator"
	XLWKS.CELLS(8,8)	=	"LinkIndex"
	XLWKS.CELLS(8,9)	=	"0"
	XLWKS.CELLS(8,10)	=	"Serial_no"
	XLWKS.CELLS(8,11)	=	"1297040"
	
																																																																																																									
'WIND UP
xlwk.sAVE
xlwk.cLOSE
SET xlwk = Nothing
SET xlobj = NOTHING

msgbox "ResQ WorkOrder function successfully loaded into Input sheet!"
'Quick Run the Script
'Set App = CreateObject("QuickTest.Application")
'App.Open "C:\RIL_TTAF_REPLICA\Scripts\AutomationDriverScript_Fulfilment"
'Set qtResult = CreateObject("QuickTest.RunResultsOptions")
'qtResult.ResultsLocation = "C:\RIL_TTAF_REPLICA\Output"
'App.Test.Run qtResult
'SET qtResult = Nothing
'Set App = NOTHING



