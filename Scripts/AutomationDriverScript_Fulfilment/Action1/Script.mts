'''*************************************************
	'''*****Action Begins******'''
'''*************************************************
Call Fn_Delete_Cookies()
Call Fn_Set_BrowsePage
Call Fn_CreatSapObject
Call Fn_GetMachineIP

ApplicationType = "Mobile"
'ApplicationType = ""
sRelease = " Build0.5 || E2E || Fulfillment || E2E Orders "
'sRelease = " Build0.4 || ST || Fulfillment || Order Creation "


iTemp = Environment.Value("ActionIteration")
If DataTable.LocalSheet.GetRowCount < 2 Then
    On Error Resume Next
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
    If Err.NUmber <> "0" Then
        DataTable.LocalSheet.AddParameter "Testing", ""
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
        Err.Clear
    	DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 1) = iTemp + 1
	End If
End If

'Adding new rows to localsheet to increment QTP build in ActionIteration counter
DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 2) = iTemp + 2

'Initial activities
If DataTable.LocalSheet.GetCurrentRow = 1 Then
    'Load Env variables
    arrProcessNamesToKill = Array("notepad", "excel","iexplore","chrome","firefox") ' "outlook"
    Call Fn_Close_Process (arrProcessNamesToKill)
'    sRelPath = Fn_RegExpReplace (Environment.Value("TestDir"), "\\RIL_TTAF.*", "", True, True)
	'''Getting Drive name of testcase
	Set fsoObj = CreateObject("Scripting.filesystemobject")
	sRelPath = fsoObj.GetDriveName(Environment("TestDir"))
	Set fsoObj = Nothing
    'Reset processes
    Call Fn_TestCasesBaseState(arrProcessNamesToKill)
    'Variable - value assignment
    sConfigPath = sRelPath & "\RIL_TTAF_REPLICA\Configuration\Config.XML"
    sEnvPath = sRelPath & "\RIL_TTAF_REPLICA\Environment\EnvVariables.XML"
    sFileName = sRelPath & "\RIL_TTAF_REPLICA\Log_Files\DashboardReport\R4GDashbaordReport.html"
    'Environment.LoadFromFile sEnvPath
    g_tStart_Time = Now
    
    'Sheet to be Imported - EXECUTION SHEET
    executionSheet = sRelPath & "\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls"
    If StrComp(ApplicationType , "Mobile") = 0 Then
        DataTable.AddSheet "MobilityProperties"
        DataTable.ImportSheet executionSheet, "MobilityProperties", "MobilityProperties"
    End If
    
    Call Fn_HTMLRep_Create()
    DataTable.ImportSheet executionSheet, "DetailedTestPlan", "Global"
    iDriver_RowCount = DataTable.GlobalSheet.GetRowCount
'    Call PassSteps("Total Execution row (iDriver_RowCount) : " & iDriver_RowCount)
    Report_Row = Fn_CreateExcelDashBoardReport()
Else
    'Case I : Issue with test case execution. Proceed to next test case iteration
    If bTestCompleted = False Then
         bTestCompleted = True
         dtRow = dtRow + iDriver_Count
         Call Fn_SaveExcelDashBoardReport()
         Call PassSteps("Issue with execution of function " & functionname & ". Move2Next Testcase action has been initiated")
         Call PassSteps("Current Execution row (dtRow) : " & dtRow - iDriver_Count)
         Call PassSteps("Current TetCase Steps Count (iDriver_Count) : " & iDriver_Count)
         Call PassSteps("Total Execution row (iDriver_RowCount) : " & iDriver_RowCount)
        'Perform the logging activity
        J= DataTable.LocalSheet.GetParameter("Testing").ValueByRow(3)
        Call Fn_ReportFailure(J)
        Call PassSteps("Move2Next Testcase action has been completed. Next test case execution started (if any).")
        'CASE II : last test case having some issues
        If cint(dtRow) >= cint(iDriver_RowCount) Then
            Call Fn_ExecutionWrapUp()
            ExitTest
        Else
            Call Fn_ExecutionWrapUp()
        End If
    Else
        'CASE II : Empty TestCaseID
        If bEmptyTest Then
            Call PassSteps("Empty testCaseID found. Possibly End of Execution.")
            Call Fn_ExecutionWrapUp()
        End If
        'CASE III : Normal flow
        If cint(dtRow) >= cint(iDriver_RowCount) Then
            Call PassSteps("End of Execution.")
            Call Fn_ExecutionWrapUp()
        End If
        'CASE IV : Critical error. Error with driver functions.
        If criticalError = true Then
            Call FailSteps("Error::Error:Issue with driver functions. Check the logs.")
            Call Fn_ExecutionWrapUp()
        End If
        'Execution wrapup
        Call Fn_ExecutionWrapUp()
    End If
End If

StartCount = dtRow - 1
Failed_Counter = 0
Report_Col =  7
iDriver_Count = 0

For J = dtRow To iDriver_RowCount step 1
    DataTable.LocalSheet.GetParameter("Testing").ValueByRow(3) = J
    DataTable.GlobalSheet.SetCurrentRow J
    TestCaseID = DataTable.Value(1, 1)
    OTestCaseName = TestCaseID
    DataTable.GlobalSheet.SetNextRow
    TestCaseID1 = DataTable.Value(1, 1)
    DataTable.GlobalSheet.SetPrevRow
    iDriver_Count = iDriver_Count + 1
    If TestCaseID <> TestCaseID1 Then
    	oname = TestCaseID
    	oStartTIme = Now
    	if instr(date, "/") > 0 then
    		otime = Replace(date,"/","_") &"_"& Replace(time,":","_")
    	ElseIf instr(date, "-") > 0 Then
    		otime = Replace(date,"-","_") &"_"& Replace(time,":","_")
    	End If
        EndCount = J
        bTestCompleted = false
        criticalError = false
        Call Create_dictionaryobjects (Report_Row, Report_Col, iDriver_Count, StartCount + 1, EndCount)
        LogCount = LogCount + 1
        criticalError = true
        bTestCompleted = true
        oStopTime = Now
		oTotalTime =  Fn_ExecutionTime(oStartTIme,oStopTime)
        oDict1(oname) = blnOutput &"|"& oTotalTime
        If blnOutput = false Then
        	oDict2(oname) = datatable.value(3,1)
        End If
        Call Fn_TestCaseWrapUp (J)
    Else
		oname = TestCaseID
    	otime = Replace(date,"-","_") &"_"& Replace(time,":","_")
    	oDict1(oname) = blnOutput
        If blnOutput = false Then
        	oDict2(oname) = datatable.value(3,1)
        End If
    End If

    If TestCaseID1 = "" Then
        bEmptyTest = True
        Exit For
    End If
Next

'''HTMl Report
Call Create_Validation_HTML(oDict1,oDict2,oDict3)

'''Send Email
'Call Fn_Sendmailreport_fromQTP(htmlFile)


'''*************************************************
	'''*****Action Ends******'''
'''*************************************************



'Set objExcel = CreateObject("Excel.Application")
'
'objExcel.Visible = False
'
'Set objWorkbook = objExcel.Workbooks.Open("C:\Users\priyanka1.das\Desktop\Data_y.xlsx")
'Set objWorksheet = objExcel.ActiveWorkbook.Worksheets("FLN_List")
'
''i = 2
'
'
'Do Until objWorksheet.Cells(2, 1).Value = "delete"
'
''    If objExcel.Cells(i, 1).Value = "delete" Then
'
'        Set objRange = objWorksheet.Cells(2, 1).EntireRow
'
'        objRange.Delete
'
''        i = i - 1
'
''    End If
'
''    i = i + 1
'
'Loop
'
'objWorkbook.Save

''Set obj=SAPGuiSession("type:=GuiSession").SAPGuiWindow("type:=GuiMainWindow")
'obj.SAPGuiGrid("columncount:=8").ActivateRow(2)
'obj.SAPGuiGrid("columncount:=8").SelectRow(2)
'obj.SAPGuiGrid("columncount:=8").ClickCell 1,"Selection"


'strProcess = "Inventory Admin Client.exe *32"
'
'Dim oShell
'Set oShell = CreateObject ("WScript.Shell")
''SET oExec=WshShell.Exec("taskkill /F /IM Inventory Admin Client.exe")
'
'oShell.Run "TaskKill /f /im " & strProcess
'
'
'
'wtitle = "Inventory Administration v8.1.0.00 Logon"
'
'wtitle = "Inventory"
'Systemutil.CloseProcessByWndTitle wtitle, True



