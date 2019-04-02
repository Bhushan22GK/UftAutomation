Call Fn_Delete_Cookies()
Call Fn_Set_BrowsePage
Call Fn_CreatSapObject


ApplicationType = ""
sRelease = " Build0.5 || SIT || Billing || Regression Execution "

iTemp = Environment.Value("ActionIteration")
If DataTable.LocalSheet.GetRowCount < 2 Then
    On Error resume next
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
    If err.NUmber <> "0" Then
        DataTable.LocalSheet.AddParameter "Testing", ""
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
        Err.clear
    End If
    On Error goto 0
    DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 1) = iTemp + 1
End If

'Adding new rows to localsheet to increment QTP build in ActionIteration counter
DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 2) = iTemp + 2

'Initial activities
If DataTable.LocalSheet.GetCurrentRow = 1 Then
    'Load Env variables
    arrProcessNamesToKill = Array("notepad", "excel","iexplore") ' "outlook"
    Call Fn_Close_Process (arrProcessNamesToKill)
    sRelPath = Fn_RegExpReplace (Environment.Value("TestDir"), "\\RIL_TTAF.*", "", true, true)
    'Reset processes
    Call Fn_TestCasesBaseState(arrProcessNamesToKill)
    'Variable - value assignment
    sConfigPath = sRelPath & "\RIL_TTAF\Configuration\Config.XML"
    sEnvPath = sRelPath & "\RIL_TTAF\Environment\EnvVariables.XML"
    sFileName = sRelPath & "\RIL_TTAF\Log_Files\DashboardReport\R4GDashbaordReport.html"
    sFileName2 = sRelPath & "\RIL_TTAF\Log_Files\DashboardReport\R4GDashbaordReportBackup.html"
    'Environment.LoadFromFile sEnvPath
    g_tStart_Time = Now
    
    'Sheet to be Imported
     executionSheet = sRelPath & "\RIL_TTAF\Input\AutomationInput_Billing_SIT_Regression_NonMobilityTestcasesBuild0.5.xls"
    If StrComp(ApplicationType , "Mobile") = 0 Then
        DataTable.AddSheet "MobilityProperties"
        DataTable.ImportSheet executionSheet, "MobilityProperties", "MobilityProperties"
    End If
    
    Call Fn_HTMLRep_Create()
    DataTable.ImportSheet executionSheet, "DetailedTestPlan", "Global"
    iDriver_RowCount = DataTable.GlobalSheet.GetRowCount
    Call PassSteps("Total Execution row (iDriver_RowCount) : " & iDriver_RowCount)
    Report_Row = Fn_CreateExcelDashBoardReport()
Else
    'Case I : Issue with test case execution. Proceed to next test case iteration
    If bTestCompleted = false Then
         bTestCompleted = true
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
        End if
        'CASE III : Normal flow
        If cint(dtRow) >= cint(iDriver_RowCount) Then
            Call PassSteps("End of Execution.")
            Call Fn_ExecutionWrapUp()
        End if
        'CASE IV : Critical error. Error with driver fucntions.
        If criticalError = true Then
            Call FailSteps("Issue with driver functions. Check the logs.")
            Call Fn_ExecutionWrapUp()
        End If
        'Execution wrapup
        Call Fn_ExecutionWrapUp()
    End if
End If

StartCount = dtRow - 1
Failed_Counter = 0
Report_Col =  7
iDriver_Count = 0

For J = dtRow To iDriver_RowCount step 1
    DataTable.LocalSheet.GetParameter("Testing").ValueByRow(3) = J
    DataTable.GlobalSheet.SetCurrentRow J
    TestCaseID = DataTable.Value(1, 1)
    DataTable.GlobalSheet.SetNextRow
    TestCaseID1 = DataTable.Value(1, 1)
    DataTable.GlobalSheet.SetPrevRow
    iDriver_Count = iDriver_Count + 1
    If TestCaseID <> TestCaseID1 Then
        EndCount = J
        bTestCompleted = false
        criticalError = false
        Call Create_dictionaryobjects (Report_Row, Report_Col, iDriver_Count, StartCount + 1, EndCount)
        criticalError = true
        bTestCompleted = true
        Call Fn_TestCaseWrapUp (J)
    End If
    'sUNIL RATHORE ON 27jUN2014... HANDLING EMPTY ROWS
    If TestCaseID1 = "" Then
        bEmptyTest = True
        Exit For
    End If
Next