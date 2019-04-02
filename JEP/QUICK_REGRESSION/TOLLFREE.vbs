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
'LOGIN
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_TollFree"
	XLWKS.CELLS(8,2)	=	"Jep_Login"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(8,4)	=	"URL"
	XLWKS.CELLS(8,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(8,6)	=	"User"
	XLWKS.CELLS(8,7)	=	"karan1.gupta"
	XLWKS.CELLS(8,8)	=	"Password"
	XLWKS.CELLS(8,9)	=	"Denmark100"


'CAF
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_TollFree"
	XLWKS.CELLS(9,2)	=	"Jep_CAF_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_CAF_Creation()"
	XLWKS.CELLS(9,4)	=	"Circle"
	XLWKS.CELLS(9,5)	=	"MUMBAI"
	XLWKS.CELLS(9,6)	=	"ProductType"
	XLWKS.CELLS(9,7)	=	"Tollfree"
	XLWKS.CELLS(9,8)	=	"Email"
	XLWKS.CELLS(9,9)	=	""
	XLWKS.CELLS(9,10)	=	"AsName"


'CF
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_TollFree"
	XLWKS.CELLS(10,2)	=	"Jep_CF_Creation"
	XLWKS.CELLS(10,3)	=	"Fn_Jep_CF_Creation()"
	XLWKS.CELLS(10,4)	=	"Product"
	XLWKS.CELLS(10,5)	=	"Tollfree"
	XLWKS.CELLS(10,6)	=	"ProductOfr"
	XLWKS.CELLS(10,7)	=	"Tollfree Test Offer"
	XLWKS.CELLS(10,8)	=	"ProductNm"
	XLWKS.CELLS(10,9)	=	"Tollfree Test Plan"
	XLWKS.CELLS(10,10)	=	"BillMode"
	XLWKS.CELLS(10,11)	=	"Postpaid"
	XLWKS.CELLS(10,12)	=	"BillPeriod"
	XLWKS.CELLS(10,13)	=	"Monthly"
	XLWKS.CELLS(10,14)	=	"PaymentTerm"
	XLWKS.CELLS(10,15)	=	"Arrears-30 Days Terms"
	XLWKS.CELLS(10,16)	=	"ContractPeriod"
	XLWKS.CELLS(10,17)	=	"24"

'CO APPROVE
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_TollFree"
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


'SAF
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_TollFree"
	XLWKS.CELLS(12,2)	=	"Jep_SAF_Creation"
	XLWKS.CELLS(12,3)	=	"Fn_Jep_SAF_Creation()"
	XLWKS.CELLS(12,4)	=	"Product"
	XLWKS.CELLS(12,5)	=	"Tollfree"
	XLWKS.CELLS(12,6)	=	"Circle"
	XLWKS.CELLS(12,7)	=	"MUMBAI"
	XLWKS.CELLS(12,8)	=	"Term_no_type"
	XLWKS.CELLS(12,9)	=	"FLN"

'LOGOUT
	XLWKS.CELLS(13,1)	=	"TestCase_Jep_TollFree"
	XLWKS.CELLS(13,2)	=	"Jep_Logout"
	XLWKS.CELLS(13,3)	=	"Fn_Jep_Logout()"
	
	
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
