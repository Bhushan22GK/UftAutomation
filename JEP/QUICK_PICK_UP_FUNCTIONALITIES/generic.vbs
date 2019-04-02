'E2E Input sheet
Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
xlobj.DisplayAlerts = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")
set xlwks1 = xlwk.Worksheets("jep_lookup")
Columncount = xlwks1.UsedRange.Columns.Count

'DataSheet QUICK INPUT Tab
Set xlobj1 = CreateObject("Excel.Application")
xlobj1.Application.Visible = False
xlobj1.DisplayAlerts = False
Set xlwm = xlobj1.Workbooks.Open("C:\RIL_TTAF_REPLICA\Datasheet\Data.xlsx")
set xlwks2 = xlwm.Worksheets("QUICK_INPUT")
start_row = xlwks2.cells(19,6)
end_row = xlwks2.cells(19,7)
no_of_rows = end_row - start_row


'Clear rows first
for j = 1 to Columncount
	xlwks.cells(8+no_of_rows+1,j) = ""
next
xlwk.Save


'STUFF Input
for i = 1 to no_of_rows
	for j = 1 to columncount
		xlwks.cells(8+i-1,j) = xlwks1.cells(start_row+i-1,j)
	next
next
xlwk.save

																																																																																																									
'WIND UP
xlwk.sAVE
xlwk.cLOSE
SET xlwk = Nothing
SET xlobj = NOTHING

xlwm.sAVE
xlwm.cLOSE
SET xlwm = Nothing
SET xlobj1 = NOTHING

'Indicating that required operation has been done
msgbox "HQ/BL/AS/CAF Functions has been loaded successfully to Input Sheet!!"


'Quick Run the Script
'Set App = CreateObject("QuickTest.Application")
'App.Open "C:\RIL_TTAF_REPLICA\Scripts\AutomationDriverScript_Fulfilment"
'Set qtResult = CreateObject("QuickTest.RunResultsOptions")
'qtResult.ResultsLocation = "C:\RIL_TTAF_REPLICA\Output"
'App.Test.Run qtResult
'SET qtResult = Nothing
'Set App = NOTHING



