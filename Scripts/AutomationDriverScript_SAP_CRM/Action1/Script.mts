RegAppName = "SAPCRM O2A"
Call Fn_Set_BrowsePage ()
Call Fn_Delete_Cookies()
Call Fn_CreatSapObject

sRelease = "SIT || SAPCRM O2A Automation || R0.5 || Execution Report"

iTemp = Environment.Value("ActionIteration")
If DataTable.LocalSheet.GetRowCount < 2 Then
    On Error resume next
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
    If err.NUmber <> "0" Then
        DataTable.LocalSheet.AddParameter "Testing", ""
        DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp) = iTemp
        err.clear
    End If
    On Error goto 0
    DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 1) = iTemp + 1
End If

'Adding new rows to localsheet to increment QTP build in ActionIteration counter
DataTable.LocalSheet.GetParameter("Testing").ValueByRow(iTemp + 2) = iTemp + 2

'Initial activities
If DataTable.LocalSheet.GetCurrentRow = 1 Then
    arrProcessNamesToKill = Array("notepad", "excel","iexplore") ' "outlook"
    Call Fn_TestCasesBaseState(arrProcessNamesToKill)
    'Load Env variables
    sRelPath = Fn_RegExpReplace (Environment.Value("TestDir"), "\\RIL_TTAF.*", "", true, true)
    sConfigPath = sRelPath & "\RIL_TTAF\Configuration\Config.XML"
    sEnvPath = sRelPath & "\RIL_TTAF\Environment\EnvVariables.XML"
    sFileName = sRelPath & "\RIL_TTAF\Log_Files\DashboardReport\R4GDashbaordReport.html"
    
    sysDate = now()
    newDateTime = day(sysDate) & month(sysDate) & year(sysDate) & "_" & Hour(Time) & Minute(Time) & Second(Time)
    
    '   Backup Previous Run Dashboard (.html) report
    Set fso=createobject("Scripting.FileSystemObject")
    sFileNameBackup_PreviousRun = sRelPath & "\RIL_TTAF\Log_Files\DashboardReport\R4GDashbaordReport" & "_" & newDateTime &".html"
    If fso.FileExists(sFileName) Then
       fso.CopyFile sFileName, sFileNameBackup_PreviousRun
       fso.DeleteFile sFileName
    End If
    
    '   Backup Previous Run Logfile (.txt) report
    sLogFileDir = "C:\RIL_TTAF\Log_Files\Detail_Logs\"
    sLogFileName = sLogFileDir & "LogFile.txt"
    sLogFileNameBackup_PreviousRun = sLogFileDir & "\LogFile" & "_" & newDateTime &".txt"
    If fso.FileExists(sLogFileName) Then
       fso.CopyFile sLogFileName, sLogFileNameBackup_PreviousRun
       fso.DeleteFile sLogFileName
    End If
    strScriptFile= sLogFileName
    
    '	Check SAP CRM HTML Logs Folder (SAPCRM_HTML_REPORTS\ExecutionDate) Exists
    If RegAppName = "SAPCRM O2A" Then
    	sSAPCRM_HTML_Path = sRelPath & "\RIL_TTAF\Log_Files\SAPCRM_HTML_REPORTS\"
    	sSAPCRM_HTML_FolderName = sSAPCRM_HTML_Path & year(sysDate) & month(sysDate) & day(sysDate) & "\"
	    If fso.FolderExists(sSAPCRM_HTML_FolderName) = True Then    
	    	Call PassSteps("SAPCRMHTML Results Folder Exists " &sSAPCRM_HTML_FolderName)
	    Else
	    	fso.CreateFolder(sSAPCRM_HTML_FolderName)
	    End If
    End If
      
    Environment.LoadFromFile sEnvPath
    g_tStart_Time = Now
    Fn_HTMLRep_Create()
    
 	DataTable.ImportSheet sRelPath & "\RIL_TTAF\Input\ART_SAPCRM_O2A.xls", "DetailedTestPlan", "Global"
    
	Row_Count = DataTable.GlobalSheet.GetRowCount
    Report_Row = Fn_CreateExcelDashBoardReport()
Else

    'Case I : Issue with test case execution. Proceed to next test case iteration
    If bTestCompleted = false Then
         criticalError = true
         Call Fn_SaveExcelDashBoardReport()
         Call PassSteps("Issue with execution of function " &functionname & ". Move2Next Testcase action has been initiated")
        'Perform the logging activity
        J= DataTable.LocalSheet.GetParameter("Testing").ValueByRow(3)
        Call Fn_ReportFailure(J)
        Call Fn_Sendmail_fromQTP()
        arrProcessNamesToKill = Array("notepad", "iexplore") '  "outlook"
        Call Fn_TestCasesBaseState(arrProcessNamesToKill)
        DataTable.LocalSheet.SetCurrentRow 1
        dtRow = dtRow + Count
        Count = 0
        Call PassSteps("Move2Next Testcase action has been completed. Next test case execution started (if any).")
        'CASE II : last test case having some issues
        If cint(dtRow) >= cint(Row_Count) Then
            Call Fn_ExecutionWrapUp()
            ExitTest
        End if
    Else
        
        'CASE II : Empty TestCaseID
        If bEmptyTest Then
            Call PassSteps("Empty testCaseID found. Possibly End of Execution.")
            Call Fn_ExecutionWrapUp()
        End if
        'CASE III : Normal flow
        If cint(dtRow) >= cint(Row_Count) Then
            Call PassSteps("End of Execution.")
            Call Fn_ExecutionWrapUp()
        End if
        'CASE IV : Critical error. Error with driver fucntions.
        If criticalError = true Then
            Call FailSteps("Issue with driver functions. Check the logs.")
            Call Fn_ExecutionWrapUp()
        End If
'        'Execution wrapup
'        Call Fn_ExecutionWrapUp()
    End if
End If

StartCount = dtRow - 1
Failed_Counter = 0
Report_Col =  7

For J = dtRow To Row_Count
    DataTable.LocalSheet.GetParameter("Testing").ValueByRow(3) = J
    DataTable.GlobalSheet.SetCurrentRow J
    TestCaseID = DataTable.Value(1, 1)
    DataTable.GlobalSheet.SetNextRow
    TestCaseID1 = DataTable.Value(1, 1)
    DataTable.GlobalSheet.SetPrevRow
    Count = Count + 1
    If TestCaseID <> TestCaseID1 Then
        EndCount = J
        bTestCompleted = false
        criticalError = false
        Call Create_dictionaryobjects (Report_Row, Report_Col, Count, StartCount + 1, EndCount)
        criticalError = true
        bTestCompleted = true
        Call Fn_TestCaseWrapUp (J)
    End If
    'sUNIL RATHOREON 27jUN2014... HANDLING EMPTY ROWS
    If TestCaseID1 = "" Then
        bEmptyTest = true
        Exit For
    End If
Next

Call Fn_ExecutionWrapUp()

Function Fn_ExecutionWrapUp()
        'Save all the reports
        g_tEnd_Time = now()
        Call Fn_SaveExcelDashBoardReport()
        call Fn_HTMLRep_Close()
        Call Fn_CloseExcelDashBoardReport()
        Call Fn_Sendmailreport_fromQTP(sRelease)
        wait 5        
        ExitTest
End Function