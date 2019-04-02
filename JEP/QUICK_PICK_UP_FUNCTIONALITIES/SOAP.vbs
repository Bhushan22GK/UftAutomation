

Serv_ID = "482221971115"
Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
sURL = "http://10.64.76.12:7285/enterprise/InventoryProvisioningServices/InventoryProvisioningServices"
SoapRequest = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:dat='http://rjil.com/ericsson/inventoryProvisioning/datastructure'><soapenv:Header/><soapenv:Body><dat:enterpriseServiceNIPRequest><!--Optional:--><dat:ServiceID>"&Serv_ID&"</dat:ServiceID></dat:enterpriseServiceNIPRequest></soapenv:Body></soapenv:Envelope>"
WinHttpRequest.Open  "POST", sURL, False 
WinHttpRequest.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
WinHttpRequest.setRequestHeader "Content-Transfer-Encoding", "binary"
WinHttpRequest.setRequestHeader "Connection", "keep-alive"
WinHttpRequest.setRequestHeader "SOAPAction", "enterpriseServiceNIP"
WinHttpRequest.Send  SoapRequest
sResponseText   = CStr(WinHttpRequest.ResponseText)
sFilenameOutput = "C:\RIL_TTAF_REPLICA\Output\dd.txt"
Set objFSO5 = CreateObject("Scripting.FileSystemObject") 

'''Check if MasterORN.txt file exists or not
If objFSO5.FileExists(sFilenameOutput) Then
    objFSO5.DeleteFile(sFilenameOutput)
End IF

Set objFile5 = objFSO5.CreateTextFile(sFilenameOutput,TRUE)  'saving the XML file
objFile5.Write sResponseText

set txtob = objFSO5.OpenTextFile(sFilenameOutput)
while txtob.AtEndofStream <> True
	data = txtob.ReadLine()
	
	'VPN ID Checkpoint
	VPN_ID = ""
	if instr(data,"<VPN_ID>") <> 0 then
		'call Passsteps("VPNID is fetched for"&OrderID)
		s1 = split(data,"<VPN_ID>")
		s2 = split(s1(1),"</VPN_ID>")
		VPN_ID = s2(0)
	end if
	'BVI 
	BVI = ""
	if instr(data,"<Param name=""BVI"">") <> 0 then
		'call Passsteps("BVI is fetched for"&OrderID)
		s1 = split(data,"<Param name=""BVI"">")
		s2 = split(s1(1),"</Param>")
		BVI = s2(0)
	end if
	'BVI 2
	BVI2 = ""
	if instr(data,"<Param name=""BVI2"">") <> 0 then
		'call Passsteps("BVI is fetched for"&OrderID)
		s1 = split(data,"<Param name=""BVI2"">")
		s2 = split(s1(1),"</Param>")
		BVI2 = s2(0)
	end if
	'VLAN 
	VLAN_ID = ""
	if instr(data,"<VLAN_ID>") <> 0 then
		'call Passsteps("VLAN_ID is fetched for"&OrderID)
		s1 = split(data,"<VLAN_ID>")
		s2 = split(s1(1),"</VLAN_ID>")
		VLAN_ID = s2(0)
	end if
	'PWID 
	PWID = ""
	if instr(data,"<Pseud_Wire_ID>") <> 0 then
		'call Passsteps("PWID is fetched for"&OrderID)
		s1 = split(data,"<Pseud_Wire_ID>")
		s2 = split(s1(1),"</Pseud_Wire_ID>")
		PWID = s2(0)
	end if
	
wend

NETWORK_ATTRIBUTES = " VPN_ID : " &VPN_ID &" || BVI : "&BVI &" || BVI_2 : "&BVI2 &" || VLAN_ID : "&VLAN_ID&" || PWID : "&PWID 
msgbox NETWORK_ATTRIBUTES

objFile5.Close 
Set objFSO5 = Nothing




