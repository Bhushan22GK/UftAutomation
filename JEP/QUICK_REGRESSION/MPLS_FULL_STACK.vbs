Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
Set xlwk = xlobj.Workbooks.Open("D:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")

for i = 8 to 150
	for j = 1 to 60
		xlwks.cells(i,j) = ""
	next
next

RANDOMIZE
RAND_NO = 100*RND

'JEP LOGIN
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(8,2)	=	"Jep_Login"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(8,4)	=	"URL"
	XLWKS.CELLS(8,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(8,6)	=	"User"
	XLWKS.CELLS(8,7)	=	"karan1.gupta"
	XLWKS.CELLS(8,8)	=	"Password"
	XLWKS.CELLS(8,9)	=	"Denmark100"
	
'CAF Create
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(9,2)	=	"Jep_CAF_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_CAF_Creation()"
	XLWKS.CELLS(9,4)	=	"Circle"
	XLWKS.CELLS(9,5)	=	"MUMBAI"
	XLWKS.CELLS(9,6)	=	"ProductType"
	XLWKS.CELLS(9,7)	=	"L3 MPLS VPN"
	XLWKS.CELLS(9,8)	=	"Email"
	XLWKS.CELLS(9,9)	=	""
	XLWKS.CELLS(9,10)	=	"AsName"
	
'CF Create
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(10,2)	=	"Jep_CF_Creation"
	XLWKS.CELLS(10,3)	=	"Fn_Jep_CF_Creation()"
	XLWKS.CELLS(10,4)	=	"NoOfSites"
	XLWKS.CELLS(10,5)	=	"10"
	XLWKS.CELLS(10,6)	=	"LastMile"
	XLWKS.CELLS(10,7)	=	"Ethernet"
	XLWKS.CELLS(10,8)	=	"Cos"
	XLWKS.CELLS(10,9)	=	"CoS0"
	XLWKS.CELLS(10,10)	=	"Bandwidth"
	XLWKS.CELLS(10,11)	=	"2 Mbps"
	XLWKS.CELLS(10,12)	=	"SLAType"
	XLWKS.CELLS(10,13)	=	"SLA - Standard"
	XLWKS.CELLS(10,14)	=	"vpnTopology"
	XLWKS.CELLS(10,15)	=	"HUB_SPOKE"
	XLWKS.CELLS(10,16)	=	"CPEProvidedBy1"
	XLWKS.CELLS(10,17)	=	"Customer"
	XLWKS.CELLS(10,18)	=	"PaymentTerm"
	XLWKS.CELLS(10,19)	=	"Arrears-30 Days Terms"
	XLWKS.CELLS(10,20)	=	"BillPeriod"
	XLWKS.CELLS(10,21)	=	"Monthly"
	XLWKS.CELLS(10,22)	=	"ContractPeriod"
	XLWKS.CELLS(10,23)	=	"24"
	XLWKS.CELLS(10,24)	=	"CPEVendor1"
	XLWKS.CELLS(10,25)	=	"REL"
	XLWKS.CELLS(10,26)	=	"CPEModel1"
	XLWKS.CELLS(10,27)	=	"REL"
	XLWKS.CELLS(10,28)	=	"CPEType1"
	XLWKS.CELLS(10,29)	=	"Router"
	XLWKS.CELLS(10,30)	=	"Product"
	XLWKS.CELLS(10,31)	=	"MPLS"
	XLWKS.CELLS(10,32)	=	"ProductOfr"
	XLWKS.CELLS(10,33)	=	"MPLS VPN Offering"
	XLWKS.CELLS(10,34)	=	"ProductNm"
	XLWKS.CELLS(10,35)	=	"MPLS"
	XLWKS.CELLS(10,36)	=	"BillMode"
	XLWKS.CELLS(10,37)	=	"Postpaid"
	XLWKS.CELLS(10,38)	=	"ConnectionString"
	XLWKS.CELLS(10,39)	=	"10.64.61.86"
	XLWKS.CELLS(10,40)	=	"Client"
	XLWKS.CELLS(10,41)	=	"900"
	XLWKS.CELLS(10,42)	=	"User"
	XLWKS.CELLS(10,43)	=	"T37300175"
	XLWKS.CELLS(10,44)	=	"Password"
	XLWKS.CELLS(10,45)	=	"Karan@2026"
	XLWKS.CELLS(10,46)	=	"Language"
	XLWKS.CELLS(10,47)	=	"EN"
	XLWKS.CELLS(10,48)	=	"SystemNumber"
	XLWKS.CELLS(10,49)	=	"00"
	
'CO APPROVE
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(11,2)	=	"COApproval"
	XLWKS.CELLS(11,3)	=	"Fn_Fast_COApproval()"
	XLWKS.CELLS(11,4)	=	"User"
	XLWKS.CELLS(11,5)	=	"T37300175"
	XLWKS.CELLS(11,6)	=	"Password"
	XLWKS.CELLS(11,7)	=	"Karan@2026"
	XLWKS.CELLS(11,8)	=	"ConnectionString"
	XLWKS.CELLS(11,9)	=	"10.64.61.86"
	XLWKS.CELLS(11,10)	=	"SystemNumber"
	XLWKS.CELLS(11,11)	=	"00"
	XLWKS.CELLS(11,12)	=	"Language"
	XLWKS.CELLS(11,13)	=	"EN"
	XLWKS.CELLS(11,14)	=	"Client"
	XLWKS.CELLS(11,15)	=	"900"
	XLWKS.CELLS(11,16)	=	"RFCCall"
	XLWKS.CELLS(11,17)	=	"ZCRM_ENT_ORD_STAT_UPDATE"

	'SAF Create
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(12,2)	=	"Jep_SAF_Creation"
	XLWKS.CELLS(12,3)	=	"Fn_Jep_SAF_Creation()"
	XLWKS.CELLS(12,4)	=	"SiteType"
	XLWKS.CELLS(12,5)	=	"Hub"
	XLWKS.CELLS(12,6)	=	"CPERoutingProtocol"
	XLWKS.CELLS(12,7)	=	"STATIC"
	XLWKS.CELLS(12,8)	=	"LanIp"
	XLWKS.CELLS(12,9)	=	"10.2.3.21"
	XLWKS.CELLS(12,10)	=	"WanIp"
	XLWKS.CELLS(12,11)	=	"10.3.3.22"
	XLWKS.CELLS(12,12)	=	"CpeIp1"
	XLWKS.CELLS(12,13)	=	"10.4.3.21"
	XLWKS.CELLS(12,14)	=	"WanIp1"
	XLWKS.CELLS(12,15)	=	"10.5.3.22"
	XLWKS.CELLS(12,16)	=	"OtherWan"
	XLWKS.CELLS(12,17)	=	"10.9.2.1"
	XLWKS.CELLS(12,18)	=	"IpAddrSelect"
	XLWKS.CELLS(12,19)	=	"IPv6"
	XLWKS.CELLS(12,20)	=	"Product"
	XLWKS.CELLS(12,21)	=	"MPLS"
	XLWKS.CELLS(12,22)	=	"InstLocNm"
	XLWKS.CELLS(12,23)	=	"InstallationLocation"
	XLWKS.CELLS(12,24)	=	"Pincode"
	XLWKS.CELLS(12,25)	=	"401107"
	XLWKS.CELLS(12,26)	=	"TechFname"
	XLWKS.CELLS(12,27)	=	"Tfirst"
	XLWKS.CELLS(12,28)	=	"TechMname"
	XLWKS.CELLS(12,29)	=	"Tmiddle"
	XLWKS.CELLS(12,30)	=	"TechLname"
	XLWKS.CELLS(12,31)	=	"Tlast"
	XLWKS.CELLS(12,32)	=	"TechMob"
	XLWKS.CELLS(12,33)	=	"9478299383"
	XLWKS.CELLS(12,34)	=	"TechEmail"
	XLWKS.CELLS(12,35)	=	"Tech@mail.jio.com"
	XLWKS.CELLS(12,36)	=	"custHandOff"
	XLWKS.CELLS(12,37)	=	"GE_OPTICAL"

	
'LOGOUT
	XLWKS.CELLS(13,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(13,2)	=	"Jep_Logout"
	XLWKS.CELLS(13,3)	=	"Fn_Jep_Logout()"
	
'NPE LOGIN
	XLWKS.CELLS(14,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(14,2)	=	"Jep_Login"
	XLWKS.CELLS(14,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(14,4)	=	"URL"
	XLWKS.CELLS(14,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/"
	XLWKS.CELLS(14,6)	=	"User"
	XLWKS.CELLS(14,7)	=	"Sunil1.Yadav"
	XLWKS.CELLS(14,8)	=	"Password"
	XLWKS.CELLS(14,9)	=	"Linus@1236"
	
'NPE APPROVAL
	XLWKS.CELLS(15,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(15,2)	=	"Jep_NEP_Creation"
	XLWKS.CELLS(15,3)	=	"Fn_Jep_NEP_Approval()"
	XLWKS.CELLS(15,4)	=	"Product"
	XLWKS.CELLS(15,5)	=	"MPLS"
	XLWKS.CELLS(15,6)	=	"NEIDType"
	XLWKS.CELLS(15,7)	=	"L2 Switch"
	XLWKS.CELLS(15,8)	=	"NESiteName"
	XLWKS.CELLS(15,9)	=	"INMUMUMBXXXXNB0006"
	XLWKS.CELLS(15,10)	=	"NEDeviceName"
	XLWKS.CELLS(15,11)	=	"INMUMUMBXXXXNB0006ENBESS002"
	XLWKS.CELLS(15,12)	=	"NEPortAccess"
	XLWKS.CELLS(15,13)	=	""
	XLWKS.CELLS(15,14)	=	"PinCode"
	XLWKS.CELLS(15,15)	=	"421204"


'LOGOUT
	XLWKS.CELLS(16,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(16,2)	=	"Jep_Logout"
	XLWKS.CELLS(16,3)	=	"Fn_Jep_Logout()"
	
	
'TIBCO QUERY	
	XLWKS.CELLS(17,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(17,2)	=	"TIBCOCLE_JEP"
	XLWKS.CELLS(17,3)	=	"Fn_TIBCOCLE_JEP()"
	
'INSTALLATION WO
	XLWKS.CELLS(18,1)	=	"TestCase_Jep_MPLS" &RAND_NO
	XLWKS.CELLS(18,2)	=	"RESQ_CycloBackEnd_JEP"
	XLWKS.CELLS(18,3)	=	"Fn_RESQ_CycloBackEnd_JEP()"
	XLWKS.CELLS(18,4)	=	"ConnectionString"
	XLWKS.CELLS(18,5)	=	"10.64.74.87"
	XLWKS.CELLS(18,6)	=	"Client"
	XLWKS.CELLS(18,7)	=	"444"
	XLWKS.CELLS(18,8)	=	"User"
	XLWKS.CELLS(18,9)	=	"P55013142"
	XLWKS.CELLS(18,10)	=	"Password"
	XLWKS.CELLS(18,11)	=	"Mumbai@2020"
	XLWKS.CELLS(18,12)	=	"Language"
	XLWKS.CELLS(18,13)	=	"EN"
	XLWKS.CELLS(18,14)	=	"SystemNumber"
	XLWKS.CELLS(18,15)	=	"00"
	XLWKS.CELLS(18,16)	=	"SAPGuiOKCode"
	XLWKS.CELLS(18,17)	=	"wui_sso"
	XLWKS.CELLS(18,18)	=	"Linkname"
	XLWKS.CELLS(18,19)	=	"ZRESQ_CASECO - ZResQ Case Coordinator"
	XLWKS.CELLS(18,20)	=	"LinkIndex"
	XLWKS.CELLS(18,21)	=	"0"
	XLWKS.CELLS(18,22)	=	"Serial_no"
	XLWKS.CELLS(18,23)	=	"1297040"
	XLWKS.CELLS(18,24)	=	"Case"
	
	
																																																																																																																																																																																																																																																											



'WIND UP
xlwk.sAVE
xlwk.cLOSE
SET xlwk = Nothing
SET xlobj = NOTHING


Set App = CreateObject("QuickTest.Application")
App.Open "D:\RIL_TTAF_REPLICA\Scripts\AutomationDriverScript_Fulfilment"
Set qtResult = CreateObject("QuickTest.RunResultsOptions")
qtResult.ResultsLocation = "D:\RIL_TTAF_REPLICA\Output"
App.Test.Run qtResult
SET qtResult = Nothing
Set App = NOTHING