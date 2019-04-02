RequestID = InputBox("Enter RequestID")

envelope_URL = "http://schemas.xmlsoap.org/soap/envelope/"
SURL = "http://10.63.52.206:9002/cwf/services/updateO2A?wsdl"
SoapRequest = "<soapenv:Envelope xmlns:soapenv=" &Chr(34)  &envelope_URL & Chr(34) &"><soapenv:Header/><soapenv:Body><FTTXAsyncUpdate><!--Optional:--><!--Optional:--><correlationID>" &RequestID &"</correlationID><status>1</status></FTTXAsyncUpdate></soapenv:Body></soapenv:Envelope>"
				
				

Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
'sURL = "http://10.64.76.12:7285/enterprise/InventoryProvisioningServices/InventoryProvisioningServices"
'SoapRequest = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:dat='http://rjil.com/ericsson/inventoryProvisioning/datastructure'><soapenv:Header/><soapenv:Body><dat:enterpriseServiceNIPRequest><!--Optional:--><dat:ServiceID>"&Serv_ID&"</dat:ServiceID></dat:enterpriseServiceNIPRequest></soapenv:Body></soapenv:Envelope>"
WinHttpRequest.Open  "POST", sURL, False 
WinHttpRequest.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
WinHttpRequest.setRequestHeader "Content-Transfer-Encoding", "binary"
WinHttpRequest.setRequestHeader "Connection", "keep-alive"
WinHttpRequest.setRequestHeader "SOAPAction", "update"
WinHttpRequest.Send  SoapRequest