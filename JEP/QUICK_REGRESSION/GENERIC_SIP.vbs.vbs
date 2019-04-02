Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
Set xlwk = xlobj.Workbooks.Open("D:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")

for i = 8 to 150
	for j = 1 to 60
		xlwks.cells(i,j) = ""
	next
next


'JEP LOGIN
	XLWKS.CELLS(8,1)	=	"TestCase_Jep_GEN"
	XLWKS.CELLS(8,2)	=	"Jep_Login"
	XLWKS.CELLS(8,3)	=	"Fn_Jep_Login()"
	XLWKS.CELLS(8,4)	=	"URL"
	XLWKS.CELLS(8,5)	=	"https://jep.bss.jiolabs.com:9292/JEP/home"
	XLWKS.CELLS(8,6)	=	"User"
	XLWKS.CELLS(8,7)	=	"karan1.gupta"
	XLWKS.CELLS(8,8)	=	"Password"
	XLWKS.CELLS(8,9)	=	"Denmark100"


'HQ CreateObject
	XLWKS.CELLS(9,1)	=	"TestCase_Jep_GEN"
	XLWKS.CELLS(9,2)	=	"Jep_HQ_Creation"
	XLWKS.CELLS(9,3)	=	"Fn_Jep_HQ_Creation()"
	XLWKS.CELLS(9,4)	=	"Company Name"
	XLWKS.CELLS(9,5)	=	""
	XLWKS.CELLS(9,6)	=	"Email"
	XLWKS.CELLS(9,8)	=	"PANNumber"
	XLWKS.CELLS(9,9)	=	""
	XLWKS.CELLS(9,10)	=	"Pincode"
	XLWKS.CELLS(9,11)	=	"400049"
	XLWKS.CELLS(9,12)	=	"BuildingNumber"
	XLWKS.CELLS(9,13)	=	"12"
	XLWKS.CELLS(9,14)	=	"BuildingName"
	XLWKS.CELLS(9,15)	=	"RCP47678"
	XLWKS.CELLS(9,16)	=	"Industry_type"
	XLWKS.CELLS(9,17)	=	"Telecom"
	XLWKS.CELLS(9,18)	=	"Enterprise_segment"
	XLWKS.CELLS(9,19)	=	"Enterprise Govt"
	XLWKS.CELLS(9,20)	=	"Enterprise_sub_segment"
	XLWKS.CELLS(9,21)	=	"IT"
	XLWKS.CELLS(9,22)	=	"Enterprise_category"
	XLWKS.CELLS(9,23)	=	"GOVERNMENT DEPT"
	XLWKS.CELLS(9,24)	=	"CompanyStatus"
	XLWKS.CELLS(9,25)	=	"Govt"
	XLWKS.CELLS(9,26)	=	"Document_Type"
	XLWKS.CELLS(9,27)	=	"Reserve Bank Letter"
	XLWKS.CELLS(9,28)	=	"Annual_Turnover"
	XLWKS.CELLS(9,29)	=	"100-1000 Millions"
	XLWKS.CELLS(9,30)	=	"Total_Employees"
	XLWKS.CELLS(9,31)	=	"11-50"

	
	
'BL CreateObject
	XLWKS.CELLS(10,1)	=	"TestCase_Jep_GEN"
	XLWKS.CELLS(10,2)	=	"Jep_BL_Creation"
	XLWKS.CELLS(10,3)	=	"Fn_Jep_BL_Creation()"
	XLWKS.CELLS(10,4)	=	"Building Location Name"
	XLWKS.CELLS(10,5)	=	""
	XLWKS.CELLS(10,6)	=	"Department"
	XLWKS.CELLS(10,7)	=	"IT Software"
	XLWKS.CELLS(10,8)	=	"Designation"
	XLWKS.CELLS(10,9)	=	"Software"
	XLWKS.CELLS(10,10)	=	"Document Type"
	XLWKS.CELLS(10,11)	=	"HJ8787"
	XLWKS.CELLS(10,12)	=	"Issuing Authority"
	XLWKS.CELLS(10,13)	=	"GOV India"
	XLWKS.CELLS(10,14)	=	"Pincode"
	XLWKS.CELLS(10,15)	=	"421204"
	XLWKS.CELLS(10,16)	=	"BuildingNumber"
	XLWKS.CELLS(10,17)	=	"12"
	XLWKS.CELLS(10,18)	=	"BuildingName"
	XLWKS.CELLS(10,19)	=	"RCP47678"
	XLWKS.CELLS(10,20)	=	"POA_Type"
	XLWKS.CELLS(10,21)	=	"Reserve Bank Letter"
	XLWKS.CELLS(10,22)	=	"GST"
	XLWKS.CELLS(10,23)	=	"NO"
	XLWKS.CELLS(10,24)	=	"GST_NO"
	XLWKS.CELLS(10,25)	=	"27KANCQ2344R3Z5"
	XLWKS.CELLS(10,26)	=	"GST_REG_TYPE"
	XLWKS.CELLS(10,27)	=	"ACTIVE"
	XLWKS.CELLS(10,28)	=	"GST_STATUS"
	XLWKS.CELLS(10,29)	=	"Casual"
	XLWKS.CELLS(10,30)	=	"Product"
	XLWKS.CELLS(10,31)	=	"SIP Trunk"
	
'AS CreateObject
	XLWKS.CELLS(11,1)	=	"TestCase_Jep_GEN"
	XLWKS.CELLS(11,2)	=	"Jep_AS_Creation"
	XLWKS.CELLS(11,3)	=	"Fn_Jep_AS_Creation()"
	XLWKS.CELLS(11,4)	=	"Department"
	XLWKS.CELLS(11,5)	=	"IT Software"
	XLWKS.CELLS(11,6)	=	"Designation"
	XLWKS.CELLS(11,7)	=	"Software"
	XLWKS.CELLS(11,8)	=	"Email"
	XLWKS.CELLS(11,9)	=	""
	XLWKS.CELLS(11,10)	=	"AsName"
	XLWKS.CELLS(11,11)	=	"Karan"
	
	
'CAF
	XLWKS.CELLS(12,1)	=	"TestCase_Jep_SIPTrunk" &RAND_NO
	XLWKS.CELLS(12,2)	=	"Jep_CAF_Creation"
	XLWKS.CELLS(12,3)	=	"Fn_Jep_CAF_Creation()"
	XLWKS.CELLS(12,4)	=	"Circle"
	XLWKS.CELLS(12,5)	=	"MUMBAI"
	XLWKS.CELLS(12,6)	=	"ProductType"
	XLWKS.CELLS(12,7)	=	"SIP Trunk"
	XLWKS.CELLS(12,8)	=	"Email"
	XLWKS.CELLS(12,9)	=	""
	XLWKS.CELLS(12,10)	=	"AsName"



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


