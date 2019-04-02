'Fetch User& Passwd
Set xlobj1 = CreateObject("Excel.Application")
xlobj1.Application.Visible = False
xlobj1.DisplayAlerts = False
Set xlwk1 = xlobj1.Workbooks.Open("C:\RIL_TTAF_REPLICA\Datasheet\Data.xlsx")
Set xlws1 = xlwk1.Worksheets("CREDENTIALS N DATA")
User=xlws1.Cells(13, 3)
Password=xlws1.Cells(14, 3)
xlwk1.Save
xlwk1.Close
xlobj1.Quit
Set xlws1 = Nothing
Set xlwk1 = Nothing
Set xlobj1 = Nothing

'SAPGuiUtil.AutoLogonByIP TestPlan_Obj.Item("ConnectionString"),TestPlan_Obj.Item("Client"),TestPlan_Obj.Item("User"),TestPlan_Obj.Item("Password"),TestPlan_Obj.Item("Language"),TestPlan_Obj.Item("SystemNumber")

Set txtobjct = CreateObject("Scripting.FileSystemObject")
'Delete the previous COApproval file, if it exists.
If txtobjct.FileExists("C:\RIL_TTAF_REPLICA\JEP\CRM_TC4\COApproval.bat") Then
  txtobjct.DeleteFile("C:\RIL_TTAF_REPLICA\JEP\CRM_TC4\COApproval.bat")
End If
'Create the COApproval File.
txtobjct.CreateTextFile("C:\RIL_TTAF_REPLICA\JEP\CRM_TC4\COApproval.bat")
set txtob = txtobjct.OpenTextFile("C:\RIL_TTAF_REPLICA\JEP\CRM_TC4\COApproval.bat",8)

'Prepare SAP Login String and load in COApproval Bat file.
login_str = "start sapshcut -system=TC4 -client=900 -user=" & User & " -pw=" & Password
txtob.WriteLine(login_str)
'Prepare and run COApproval file.
set oshell = Wscript.CreateObject("Wscript.shell")
oShell.Run "taskkill /im saplogon.exe", , True
oshell.run "cmd /k C:\\RIL_TTAF_REPLICA\\JEP\\CRM_TC4\\COApproval.bat"



