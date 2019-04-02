Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
xlobj.DisplayAlerts = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")
Columncount = xlwks.UsedRange.Columns.Count

for j = 1 to Columncount
	XLWKS.cells(8,j) = ""
	XLWKS.cells(9,j) = ""
next

'STUFF WORK ORDER
	XLWKS.CELLS(8,1)	=	"TestCase_Jep"
	XLWKS.CELLS(8,2)	=	"Fn_Fast_COApproval"
	XLWKS.CELLS(8,3)	=	"Fn_EFL_Fast_COApproval()"
																																																																																															
'WIND UP
xlwk.sAVE
xlwk.cLOSE
SET xlwk = Nothing
SET xlobj = NOTHING

msgbox "CO Approval function successfully loaded into Input sheet!"
'Quick Run the Script
'Set App = CreateObject("QuickTest.Application")
'App.Open "C:\RIL_TTAF_REPLICA\Scripts\AutomationDriverScript_Fulfilment"
'Set qtResult = CreateObject("QuickTest.RunResultsOptions")
'qtResult.ResultsLocation = "C:\RIL_TTAF_REPLICA\Output"
'App.Test.Run qtResult
'SET qtResult = Nothing
'Set App = NOTHING



