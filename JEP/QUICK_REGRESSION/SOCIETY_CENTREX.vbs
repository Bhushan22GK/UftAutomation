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
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_Society"
	XLWKS.CELLS(8,2)	=	"Jep_Login"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(8,4)	=	"URL"
	XLWKS.CELLS(8,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(8,6)	=	"User"
	XLWKS.CELLS(8,7)	=	"karan1.gupta"
	XLWKS.CELLS(8,8)	=	"Password"
	XLWKS.CELLS(8,9)	=	"Denmark100"
	
'JEP CAF
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_Society"
	XLWKS.CELLS(9,2)	=	"Jep_CAF_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_CAF_Creation()"
	XLWKS.CELLS(9,4)	=	"Circle"
	XLWKS.CELLS(9,5)	=	"MUMBAI"
	XLWKS.CELLS(9,6)	=	"ProductType"
	XLWKS.CELLS(9,7)	=	"SOCIETY CENTREX"
	XLWKS.CELLS(9,8)	=	"Email"
	XLWKS.CELLS(9,9)	=	""
	XLWKS.CELLS(9,10)	=	"AsName"


'JEP CF
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_Society"
	XLWKS.CELLS(10,2)	=	"Jep_CF_Creation"
	XLWKS.CELLS(10,3)	=	"Fn_Jep_CF_Creation()"
	XLWKS.CELLS(10,4)	=	"Product"
	XLWKS.CELLS(10,5)	=	"SOCIETY CENTREX"
	XLWKS.CELLS(10,6)	=	"ProductOfr"
	XLWKS.CELLS(10,7)	=	""
	XLWKS.CELLS(10,8)	=	"ProductNm"
	XLWKS.CELLS(10,9)	=	""
	XLWKS.CELLS(10,10)	=	"BillMode"
	XLWKS.CELLS(10,11)	=	"Postpaid"
	XLWKS.CELLS(10,12)	=	"BillPeriod"
	XLWKS.CELLS(10,13)	=	"Monthly"
	XLWKS.CELLS(10,14)	=	"PaymentTerm"
	XLWKS.CELLS(10,15)	=	"Arrears-30 Days Terms"
	XLWKS.CELLS(10,16)	=	"ContractPeriod"
	XLWKS.CELLS(10,17)	=	"24"
	XLWKS.CELLS(10,18)	=	"home-ONT"
	XLWKS.CELLS(10,19)	=	"1"
	XLWKS.CELLS(10,20)	=	"comArea-ONT"
	XLWKS.CELLS(10,21)	=	"0"
	XLWKS.CELLS(10,22)	=	"exist-ONT"
	XLWKS.CELLS(10,23)	=	"0"
	XLWKS.CELLS(10,24)	=	"free-ONT"
	XLWKS.CELLS(10,25)	=	"0"
	XLWKS.CELLS(10,26)	=	"CPEProvidedBy1"
	XLWKS.CELLS(10,27)	=	"Reliance"
	XLWKS.CELLS(10,28)	=	"CPEVendor1"
	XLWKS.CELLS(10,29)	=	""
	XLWKS.CELLS(10,30)	=	"CPEModel1"
	XLWKS.CELLS(10,31)	=	""
	XLWKS.CELLS(10,32)	=	"CPEType1"
	XLWKS.CELLS(10,33)	=	"ONT Gateway"
	XLWKS.CELLS(10,34)	=	"vpnTopology"
	XLWKS.CELLS(10,35)	=	""
	XLWKS.CELLS(10,36)	=	"Cos"
	XLWKS.CELLS(10,37)	=	""
	XLWKS.CELLS(10,38)	=	"CPE1Make_and_Model"
	XLWKS.CELLS(10,39)	=	"JioHubExpress for Society"
	
	
'CO APPROVE
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_Society"
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
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_Society"
	XLWKS.CELLS(12,2)	=	"Jep_SAF_Creation"
	XLWKS.CELLS(12,3)	=	"Fn_Jep_SAF_Creation()"
	XLWKS.CELLS(12,4)	=	"Product"
	XLWKS.CELLS(12,5)	=	"SOCIETY CENTREX"
	XLWKS.CELLS(12,6)	=	"citySel"
	XLWKS.CELLS(12,7)	=	""
	XLWKS.CELLS(12,8)	=	"stdCodeSel"
	XLWKS.CELLS(12,9)	=	""
	XLWKS.CELLS(12,10)	=	"Circle"
	XLWKS.CELLS(12,11)	=	"MUMBAI"
	XLWKS.CELLS(12,12)	=	"exist_flns"

'LOGOUT
	XLWKS.CELLS(13,1)	=	"TestCase_Jep_Society"
	XLWKS.CELLS(13,2)	=	"Jep_Logout"
	XLWKS.CELLS(13,3)	=	"Fn_Jep_Logout()"

	
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
