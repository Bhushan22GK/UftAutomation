
'**********************************************************************************************************************
'***********************CODE TO FETCH ASID LIST FROM MAIL**************************************************************
Dim olFolderInbox, Subject, SavePath
olFolderInbox = 6 : Subject = ""

SavePath = "C:\RIL_TTAF_REPLICA\JEP\"
Subject = "Jio ID"
Unread = True

'Delete previous ASID file, if any.
Set txtobjct = CreateObject("Scripting.FileSystemObject")    
'''Check if ORN_srvc.txt file exists or not
If txtobjct.FileExists("C:\RIL_TTAF_REPLICA\JEP\ASID.csv") Then
	on error resume next
	txtobjct.DeleteFile("C:\RIL_TTAF_REPLICA\JEP\ASID.csv")
End IF


Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")

'Create reference to Inbox Folder
Set oInbox = objNamespace.GetDefaultFolder(olFolderInbox)

'Find all items in the Inbox Folder
Set oAllMails = oInbox.Items
Set colFilteredItems  = oAllMails.Restrict("[Unread]=" &Unread)
Set colFilteredItems = colFilteredItems.Restrict("[Subject] = " & Subject)

dim k
For k = colFilteredItems.Count to 1 step -1
    set objMessage  = colFilteredItems.Item(k)
    intCount = objMessage.Attachments.Count
    If intCount > 0 Then
        For i = 1 To intCount 
            path = SavePath &objMessage.Attachments.Item(i).FileName
            objMessage.Attachments.Item(i).SaveAsFile path
        Next
        objMessage.Unread = False
    End If
Next
'**************************************************************************************************************************
'**************************************************************************************************************************


'**************************************************************************************************************************
'***********************CODE TO FETCH ASID FROM FILE EXTRACTED FROM MAIL***************************************************
Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\JEP\ASID.csv")
set xlws = xlwk.Worksheets("ASID")
Rowcount = xlws.UsedRange.Rows.Count

Dim ASID(100),JioID(100),Mob(100),Email(100)

if Rowcount = 1 then
   msgbox "NO ASID for today :)"
else 
    for i = 2 to Rowcount
    
	    ASID(i-1) = xlws.cells(Rowcount,1)
	    JioID(i-1) = xlws.cells(Rowcount,5)
		Mob(i-1) = xlws.cells(Rowcount,3)
		msgbox Mob(i-1)
		Email(i-1) = xlws.cells(Rowcount,4)
    
    Next
end if

xlwk.save
xlwk.close
'*************************************************************************************************************************
'**************************************************************************************************************************




'**************************************************************************************************************************
'*******************CODE TO ACTIVATE JIO ID FROM SELFCARE********************************************************************
''le=  1
'While JioID(le) <> ""
'	le=  le+1
'Wend
'if brwsrpg.WebElement("innertext:=Sign up","class:=userLoginHeader").exist then
'		Brwsrpg.WebEdit("html id:=pt1:r2:0:sb11:it1::content","name:=pt1:r2:0:sb11:it1").set JioID(1)
'		Brwsrpg.WebButton("html id:=pt1:r2:0:sb11:cb1","name:=GENERATE OTP").Click
'		Brwsrpg.Sync
'		if brwsrpg.WebElement("class:=successMsgContainerNew af_panelGroupLayout","outertext:=OTP has been sent.*").exist then
'			
'		else if Brwsrpg.WebElement("class:=errorMsgNew","html tag:=SPAN").exist then
'			msgbox Brwsrpg.WebElement("class:=errorMsgNew","html tag:=SPAN").GetROProperty("innertext")
'		end if
'end if
'


'**************************************************************************************************************************
'**************************************************************************************************************************
