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
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(8,2)	=	"Jep_Login"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(8,4)	=	"URL"
	XLWKS.CELLS(8,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(8,6)	=	"User"
	XLWKS.CELLS(8,7)	=	"karan1.gupta"
	XLWKS.CELLS(8,8)	=	"Password"
	XLWKS.CELLS(8,9)	=	"Denmark100"


	
'FILL CF DATA
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(9,2)	=	"Jep_CF_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_CF_Creation()"
	XLWKS.CELLS(9,4)	=	"NoOfSites"
	XLWKS.CELLS(9,5)	=	"9"									'GENERALISE
	XLWKS.CELLS(9,6)	=	"LastMile"
	XLWKS.CELLS(9,7)	=	"Ethernet"							'GENERALISE
	XLWKS.CELLS(9,8)	=	"CPEProvidedBy1"
	XLWKS.CELLS(9,9)	=	"Customer"							'GENERALISE
	XLWKS.CELLS(9,10)	=	"Bandwidth"							
	XLWKS.CELLS(9,11)	=	"10 Mbps"
	XLWKS.CELLS(9,12)	=	"SLAType"
	XLWKS.CELLS(9,13)	=	"SLA - Standard"
	XLWKS.CELLS(9,14)	=	"BillPeriod"
	XLWKS.CELLS(9,15)	=	"Monthly"
	XLWKS.CELLS(9,16)	=	"PaymentTerm"
	XLWKS.CELLS(9,17)	=	"Arrears-30 Days Terms"
	XLWKS.CELLS(9,18)	=	"ContractPeriod"
	XLWKS.CELLS(9,19)	=	"24"
	XLWKS.CELLS(9,20)	=	"Product"
	XLWKS.CELLS(9,21)	=	"ILL"
	XLWKS.CELLS(9,22)	=	"ProductOfr"
	XLWKS.CELLS(9,23)	=	"ILL Offering"
	XLWKS.CELLS(9,24)	=	"ProductNm"
	XLWKS.CELLS(9,25)	=	"ILL"
	XLWKS.CELLS(9,26)	=	"BillMode"
	XLWKS.CELLS(9,27)	=	"Postpaid"
	XLWKS.CELLS(9,28)	=	"CPEVendor1"
	XLWKS.CELLS(9,29)	=	"REL"
	XLWKS.CELLS(9,30)	=	"CPEModel1"
	XLWKS.CELLS(9,31)	=	"REL"
	XLWKS.CELLS(9,32)	=	"CPEType1"
	XLWKS.CELLS(9,33)	=	"Router"
	XLWKS.CELLS(9,34)	=	"ConnectionString"
	XLWKS.CELLS(9,35)	=	"10.64.61.86"
	XLWKS.CELLS(9,36)	=	"Client"
	XLWKS.CELLS(9,37)	=	"900"
	XLWKS.CELLS(9,38)	=	"User"
	XLWKS.CELLS(9,39)	=	"T37300175"
	XLWKS.CELLS(9,40)	=	"Password"
	XLWKS.CELLS(9,41)	=	"Karan@2026"
	XLWKS.CELLS(9,42)	=	"Language"
	XLWKS.CELLS(9,43)	=	"EN"
	XLWKS.CELLS(9,44)	=	"SystemNumber"
	XLWKS.CELLS(9,45)	=	"00"
	
'FILL CO APPROVE
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(10,2)	=	"COApproval"
	XLWKS.CELLS(10,3)	=	"Fn_Fast_COApproval()"
	XLWKS.CELLS(10,4)	=	"User"
	XLWKS.CELLS(10,5)	=	"T37300175"
	XLWKS.CELLS(10,6)	=	"Password"
	XLWKS.CELLS(10,7)	=	"Karan@2026"
	XLWKS.CELLS(10,8)	=	"ConnectionString"
	XLWKS.CELLS(10,9)	=	"10.64.61.86"
	XLWKS.CELLS(10,10)	=	"SystemNumber"
	XLWKS.CELLS(10,11)	=	"00"
	XLWKS.CELLS(10,12)	=	"Language"
	XLWKS.CELLS(10,13)	=	"EN"
	XLWKS.CELLS(10,14)	=	"Client"
	XLWKS.CELLS(10,15)	=	"900"
	XLWKS.CELLS(10,16)	=	"RFCCall"
	XLWKS.CELLS(10,17)	=	"ZCRM_ENT_ORD_STAT_UPDATE"
	
	
'FILL SAF
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(11,2)	=	"Jep_SAF_Creation"
	XLWKS.CELLS(11,3)	=	"Fn_Jep_SAF_Creation()"
	XLWKS.CELLS(11,4)	=	"CPERoutingProtocol"
	XLWKS.CELLS(11,5)	=	"STATIC"
	XLWKS.CELLS(11,6)	=	"SiteType"
	XLWKS.CELLS(11,7)	=	"Spoke"
	XLWKS.CELLS(11,8)	=	"LanIPAddrType"
	XLWKS.CELLS(11,9)	=	"IPv6"
	XLWKS.CELLS(11,10)	=	"LanIp"
	XLWKS.CELLS(11,11)	=	"2405:223:22:2:22:22:1:2/125"
	XLWKS.CELLS(11,12)	=	"WanIp"
	XLWKS.CELLS(11,13)	=	"10.2.3.22"
	XLWKS.CELLS(11,14)	=	"CpeIp1"
	XLWKS.CELLS(11,15)	=	"10.4.3.21"
	XLWKS.CELLS(11,16)	=	"WanIp1"
	XLWKS.CELLS(11,17)	=	"10.5.3.22"
	XLWKS.CELLS(11,18)	=	"IpAddrSelect"
	XLWKS.CELLS(11,19)	=	"IPv6"
	XLWKS.CELLS(11,20)	=	"TechEmail"
	XLWKS.CELLS(11,21)	=	"Tech@mail.jio.com"
	XLWKS.CELLS(11,22)	=	"CPEProvidedBy1"
	XLWKS.CELLS(11,23)	=	"Customer"
	XLWKS.CELLS(11,24)	=	"custHandOff"
	XLWKS.CELLS(11,25)	=	"GE_OPTICAL"
	XLWKS.CELLS(11,26)	=	"FirewallVersion"
	XLWKS.CELLS(11,27)	=	"9"
	XLWKS.CELLS(11,28)	=	"wanInterface"
	XLWKS.CELLS(11,29)	=	"GE_OPTICAL"
	XLWKS.CELLS(11,30)	=	"OtherWan"
	XLWKS.CELLS(11,31)	=	"10.4.2.1"
	XLWKS.CELLS(11,32)	=	"HouseNo"
	XLWKS.CELLS(11,33)	=	"8"
	XLWKS.CELLS(11,34)	=	"BldgName"
	XLWKS.CELLS(11,35)	=	"Bldg1"
	XLWKS.CELLS(11,36)	=	"Landmark"
	XLWKS.CELLS(11,37)	=	"LandM"
	XLWKS.CELLS(11,38)	=	"Street"
	XLWKS.CELLS(11,39)	=	"Street1"
	XLWKS.CELLS(11,40)	=	"Product"
	XLWKS.CELLS(11,41)	=	"ILL"
	XLWKS.CELLS(11,42)	=	"InstLocNm"
	XLWKS.CELLS(11,43)	=	"InstallationLocation"
	XLWKS.CELLS(11,44)	=	"Pincode"
	XLWKS.CELLS(11,45)	=	"401107"
	XLWKS.CELLS(11,46)	=	"TechFname"
	XLWKS.CELLS(11,47)	=	"Tfirst"
	XLWKS.CELLS(11,48)	=	"TechMname"
	XLWKS.CELLS(11,49)	=	"Tmiddle"
	XLWKS.CELLS(11,50)	=	"TechLname"
	XLWKS.CELLS(11,51)	=	"Tlast"
	XLWKS.CELLS(11,52)	=	"TechMob"
	XLWKS.CELLS(11,53)	=	"9478299383"
	
'JEP LOGOUT
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(12,2)	=	"Jep_Logout"
	XLWKS.CELLS(12,3)	=	"Fn_Jep_Logout()"
	
'JEP LOGIN
	'FILL LOGIN DATA
	XLWKS.CELLS(13,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(13,2)	=	"Jep_Login"
	XLWKS.CELLS(13,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(13,4)	=	"URL"
	XLWKS.CELLS(13,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(13,6)	=	"User"
	XLWKS.CELLS(13,7)	=	"sunil1.yadav"
	XLWKS.CELLS(13,8)	=	"Password"
	XLWKS.CELLS(13,9)	=	"Linus@1236"
	
	
'NPE APPROVAL
	XLWKS.CELLS(14,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(14,2)	=	"Jep_NEP_Creation"
	XLWKS.CELLS(14,3)	=	"Fn_Jep_NEP_Approval()"
	XLWKS.CELLS(14,4)	=	"Product"
	XLWKS.CELLS(14,5)	=	"ILL"
	XLWKS.CELLS(14,6)	=	"NEIDType"
	XLWKS.CELLS(14,7)	=	"L2 Switch"
	XLWKS.CELLS(14,8)	=	"NESiteName"
	XLWKS.CELLS(14,9)	=	"INMUMUMBXXXXNB0006"
	XLWKS.CELLS(14,10)	=	"NEDeviceName"
	XLWKS.CELLS(14,11)	=	"INMUMUMBXXXXNB0006ENBESS002"
	XLWKS.CELLS(14,12)	=	"NEPortAccess"
	XLWKS.CELLS(14,13)	=	""
	XLWKS.CELLS(14,14)	=	"PinCode"
	XLWKS.CELLS(14,15)	=	"421204"
	
	
'LOG OuT
	XLWKS.CELLS(15,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(15,2)	=	"Jep_Logout"
	XLWKS.CELLS(15,3)	=	"Fn_Jep_Logout()"
	
'TIBCO QUERY	
	XLWKS.CELLS(16,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(16,2)	=	"TIBCOCLE_JEP"
	XLWKS.CELLS(16,3)	=	"Fn_TIBCOCLE_JEP()"
	
'INSTALLATION WO
	XLWKS.CELLS(17,1)	=	"TestCase_Jep_ILL" & RAND_NO
	XLWKS.CELLS(17,2)	=	"RESQ_CycloBackEnd_JEP"
	XLWKS.CELLS(17,3)	=	"Fn_RESQ_CycloBackEnd_JEP()"
	XLWKS.CELLS(17,4)	=	"ConnectionString"
	XLWKS.CELLS(17,5)	=	"10.64.74.87"
	XLWKS.CELLS(17,6)	=	"Client"
	XLWKS.CELLS(17,7)	=	"444"
	XLWKS.CELLS(17,8)	=	"User"
	XLWKS.CELLS(17,9)	=	"P55013142"
	XLWKS.CELLS(17,10)	=	"Password"
	XLWKS.CELLS(17,11)	=	"Mumbai@2020"
	XLWKS.CELLS(17,12)	=	"Language"
	XLWKS.CELLS(17,13)	=	"EN"
	XLWKS.CELLS(17,14)	=	"SystemNumber"
	XLWKS.CELLS(17,15)	=	"00"
	XLWKS.CELLS(17,16)	=	"SAPGuiOKCode"
	XLWKS.CELLS(17,17)	=	"wui_sso"
	XLWKS.CELLS(17,18)	=	"Linkname"
	XLWKS.CELLS(17,19)	=	"ZRESQ_CASECO - ZResQ Case Coordinator"
	XLWKS.CELLS(17,20)	=	"LinkIndex"
	XLWKS.CELLS(17,21)	=	"0"
	XLWKS.CELLS(17,22)	=	"Serial_no"
	XLWKS.CELLS(17,23)	=	"1297040"
	XLWKS.CELLS(17,24)	=	"Case"
	
	


	
	

	
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

