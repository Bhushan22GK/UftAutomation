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


'CF
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(8,2)	=	"Jep_CF_Creation"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_CF_Creation()"
	XLWKS.CELLS(8,4)	=	"LastMile"
	XLWKS.CELLS(8,5)	=	"Ethernet"
	XLWKS.CELLS(8,6)	=	"TMN?"
	XLWKS.CELLS(8,7)	=	""
	XLWKS.CELLS(8,8)	=	"PRI?"
	XLWKS.CELLS(8,9)	=	""
	XLWKS.CELLS(8,10)	=	"Pilot"
	XLWKS.CELLS(8,11)	=	""
	XLWKS.CELLS(8,12)	=	"NoOfChannels"
	XLWKS.CELLS(8,13)	=	"20 Channels"
	XLWKS.CELLS(8,14)	=	"CPEProvidedBy1"
	XLWKS.CELLS(8,15)	=	"Customer"
	XLWKS.CELLS(8,16)	=	"CPEVendor1"
	XLWKS.CELLS(8,17)	=	"REL"
	XLWKS.CELLS(8,18)	=	"CPEModel1"
	XLWKS.CELLS(8,19)	=	"REL"
	XLWKS.CELLS(8,20)	=	"CPEType1"
	XLWKS.CELLS(8,21)	=	"IP-PBX"
	XLWKS.CELLS(8,22)	=	"Product"
	XLWKS.CELLS(8,23)	=	"SIP TRUNK"
	XLWKS.CELLS(8,24)	=	"ProductOfr"
	XLWKS.CELLS(8,25)	=	"SIP Trunk Offering"
	XLWKS.CELLS(8,26)	=	"ProductNm"
	XLWKS.CELLS(8,27)	=	"SIP Trunk"
	XLWKS.CELLS(8,28)	=	"Plan"
	XLWKS.CELLS(8,29)	=	"SIP trunk plan"
	XLWKS.CELLS(8,30)	=	"BillMode"
	XLWKS.CELLS(8,31)	=	"Postpaid"
	XLWKS.CELLS(8,32)	=	"BillPeriod"
	XLWKS.CELLS(8,33)	=	"Monthly"
	XLWKS.CELLS(8,34)	=	"PaymentTerm"
	XLWKS.CELLS(8,35)	=	"Arrears-30 Days Terms"
	XLWKS.CELLS(8,36)	=	"ContractPeriod"
	XLWKS.CELLS(8,37)	=	"24"


'SAF
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(9,2)	=	"Jep_SAF_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_SAF_Creation()"
	XLWKS.CELLS(9,4)	=	"State"
	XLWKS.CELLS(9,5)	=	"Maharashtra"
	XLWKS.CELLS(9,6)	=	"DTMF"
	XLWKS.CELLS(9,7)	=	"Out of Band"
	XLWKS.CELLS(9,8)	=	"District"
	XLWKS.CELLS(9,9)	=	"Mumbai (Suburban)"
	XLWKS.CELLS(9,10)	=	"City"
	XLWKS.CELLS(9,11)	=	"Mumbai"
	XLWKS.CELLS(9,12)	=	"Circle"
	XLWKS.CELLS(9,13)	=	"MUMBAI"
	XLWKS.CELLS(9,14)	=	"JioCentre"
	XLWKS.CELLS(9,15)	=	"I003"
	XLWKS.CELLS(9,16)	=	"wanInterface"
	XLWKS.CELLS(9,17)	=	"GE_OPTICAL"
	XLWKS.CELLS(9,18)	=	"HouseNo"
	XLWKS.CELLS(9,19)	=	"8"
	XLWKS.CELLS(9,20)	=	"BldgName"
	XLWKS.CELLS(9,21)	=	"Bldg1"
	XLWKS.CELLS(9,22)	=	"Landmark"
	XLWKS.CELLS(9,23)	=	"LandM"
	XLWKS.CELLS(9,24)	=	"Street"
	XLWKS.CELLS(9,25)	=	"Street1"
	XLWKS.CELLS(9,26)	=	"Product"
	XLWKS.CELLS(9,27)	=	"SIP TRUNK"
	XLWKS.CELLS(9,28)	=	"InstLocNm"
	XLWKS.CELLS(9,29)	=	"InstallationLocation"
	XLWKS.CELLS(9,30)	=	"Pincode"
	XLWKS.CELLS(9,31)	=	"401107"
	XLWKS.CELLS(9,32)	=	"TechFname"
	XLWKS.CELLS(9,33)	=	"Tfirst"
	XLWKS.CELLS(9,34)	=	"TechMname"
	XLWKS.CELLS(9,35)	=	"Tmiddle"
	XLWKS.CELLS(9,36)	=	"TechLname"
	XLWKS.CELLS(9,37)	=	"Tlast"
	XLWKS.CELLS(9,38)	=	"TechMob"
	XLWKS.CELLS(9,39)	=	"9478299383"
	XLWKS.CELLS(9,40)	=	"TechEmail"
	XLWKS.CELLS(9,41)	=	"Tech@mail.jio.com"
	XLWKS.CELLS(9,42)	=	"custHandOff"
	XLWKS.CELLS(9,43)	=	"GE_OPTICAL"
	XLWKS.CELLS(9,44)	=	"Pilot"
	XLWKS.CELLS(9,45)	=	"35042370"
	XLWKS.CELLS(9,46)	=	"NoOfFLN"
	XLWKS.CELLS(9,47)	=	"1"

	
'NPE
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(10,2)	=	"Jep_NEP_Creation"
	XLWKS.CELLS(10,3)	=	"Fn_Jep_NEP_Approval()"
	XLWKS.CELLS(10,4)	=	"Product"
	XLWKS.CELLS(10,5)	=	"SIP TRUNK"
	XLWKS.CELLS(10,6)	=	"NEIDType"
	XLWKS.CELLS(10,7)	=	"L2 Switch"
	XLWKS.CELLS(10,8)	=	"NESiteName"
	XLWKS.CELLS(10,9)	=	"INMUMUMBXXXXNB0006"
	XLWKS.CELLS(10,10)	=	"NEDeviceName"
	XLWKS.CELLS(10,11)	=	"INMUMUMBXXXXNB0006ENBESS002"
	XLWKS.CELLS(10,12)	=	"PinCode"
	XLWKS.CELLS(10,13)	=	"421204"
	XLWKS.CELLS(10,14)	=	"SignalingIP"
	XLWKS.CELLS(10,15)	=	"Single"
	XLWKS.CELLS(10,16)	=	"CacIncoming"
	XLWKS.CELLS(10,17)	=	"sdk"
	XLWKS.CELLS(10,18)	=	"CacOutgoing"
	XLWKS.CELLS(10,19)	=	"dfv"
	XLWKS.CELLS(10,20)	=	"Cps"
	XLWKS.CELLS(10,21)	=	"cfs"
	XLWKS.CELLS(10,22)	=	"CodecVal"
	XLWKS.CELLS(10,23)	=	"G.711"
	XLWKS.CELLS(10,24)	=	"SarSiteName"
	XLWKS.CELLS(10,25)	=	""
	XLWKS.CELLS(10,26)	=	"SarDeviceName"
	XLWKS.CELLS(10,27)	=	""
	XLWKS.CELLS(10,28)	=	"SbcDeviceName"
	XLWKS.CELLS(10,29)	=	""
	XLWKS.CELLS(10,30)	=	"SbcPortAccess"

	
'TIBCO QUERY	
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(11,2)	=	"TIBCOCLE_JEP"
	XLWKS.CELLS(11,3)	=	"Fn_TIBCOCLE_JEP()"
	
'INSTALLATION WO
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(12,2)	=	"RESQ_CycloBackEnd_JEP"
	XLWKS.CELLS(12,3)	=	"Fn_RESQ_CycloBackEnd_JEP()"
	XLWKS.CELLS(12,4)	=	"ConnectionString"
	XLWKS.CELLS(12,5)	=	"10.64.74.87"
	XLWKS.CELLS(12,6)	=	"Client"
	XLWKS.CELLS(12,7)	=	"444"
	XLWKS.CELLS(12,8)	=	"User"
	XLWKS.CELLS(12,9)	=	"P55013142"
	XLWKS.CELLS(12,10)	=	"Password"
	XLWKS.CELLS(12,11)	=	"Mumbai@2020"
	XLWKS.CELLS(12,12)	=	"Language"
	XLWKS.CELLS(12,13)	=	"EN"
	XLWKS.CELLS(12,14)	=	"SystemNumber"
	XLWKS.CELLS(12,15)	=	"00"
	XLWKS.CELLS(12,16)	=	"SAPGuiOKCode"
	XLWKS.CELLS(12,17)	=	"wui_sso"
	XLWKS.CELLS(12,18)	=	"Linkname"
	XLWKS.CELLS(12,19)	=	"ZRESQ_CASECO - ZResQ Case Coordinator"
	XLWKS.CELLS(12,20)	=	"LinkIndex"
	XLWKS.CELLS(12,21)	=	"0"
	XLWKS.CELLS(12,22)	=	"Serial_no"
	XLWKS.CELLS(12,23)	=	"1297040"
	XLWKS.CELLS(12,24)	=	"Case"
	
	

	
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