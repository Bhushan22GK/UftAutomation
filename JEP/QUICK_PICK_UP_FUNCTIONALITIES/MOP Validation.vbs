


blnOutput = false
Set xlobjS = CreateObject("Excel.Application")
xlobjS.Application.Visible = False
Set xlwkS = xlobjS.Workbooks.Open("C:\RIL_TTAF_REPLICA\Datasheet\Data.xlsx")
Set xlwsS = xlobjS.ActiveWorkbook.Worksheets("QUICK_INPUT")
Rowcount = xlwsS.UsedRange.Rows.Count
Set xlws1S = xlobjS.ActiveWorkbook.Worksheets("BULK EXECUTION DATA")
Rowcount1 = xlws1S.UsedRange.Rows.Count

Set xlobj = CreateObject("Excel.Application")
xlobj.Application.Visible = False
xlobj.DisplayAlerts = False
Set xlwk = xlobj.Workbooks.Open("C:\RIL_TTAF_REPLICA\Input\AutomationInput_E2EOrders.xls")
set xlwks = xlwk.Worksheets("DetailedTestPlan")
set xlwks1 = xlwk.Worksheets("jep_lookup")
columncount = xlwks.UsedRange.Columns.Count
columncount1 = xlwks1.UsedRange.Columns.Count

Input_index = 8
for i = 2 to Rowcount1

                if xlws1S.cells(i,9) <> "" and UCase(Trim(xlws1S.cells(i,11))) = "INVOKEHPOO SUCCESS" then
                                
                                'STUFF PRE_HPOO
                                Pre_HPOO_Index = xlwsS.cells(21,6)
                                xlwks.cells(Input_index,1) = xlwks1.cells(Pre_HPOO_Index,1)
                                xlwks.cells(Input_index,2) = xlwks1.cells(Pre_HPOO_Index,2)
                                xlwks.cells(Input_index,3) = xlwks1.cells(Pre_HPOO_Index,3)
                                Input_index = Input_index + 1
                               

							   'STUFF OrderCare_Login
                                'OrderCare_Login_Index = xlwsS.cells(22,6)
                           '     for j = 1 to columncount1
                         '                       xlwks.cells(Input_index,j) = xlwks1.cells(OrderCare_Login_Index,j)
                        '        next
                        '        Input_index = Input_index + 1
                           

								'STUFF OrderCare_ValidateOrders_Jep
                           '     OrderCare_ValidateOrders_Jep_index = xlwsS.cells(23,6)
                         '       for j = 1 to columncount1
                        '                        xlwks.cells(Input_index,j) = xlwks1.cells(OrderCare_ValidateOrders_Jep_index,j)
                        '        next
                         '       Input_index = Input_index + 1
                            

								'STUFF JEP_HPNADB
                                JEP_HPNADB_Index = xlwsS.cells(24,6)
                                for j = 1 to columncount1
                                                xlwks.cells(Input_index,j) = xlwks1.cells(JEP_HPNADB_Index,j)
                                next
                                Input_index = Input_index + 1
                                'STUFF JEP_HPOO_Template
                                JEP_HPOO_Template_Index = xlwsS.cells(25,6)
                                for j = 1 to columncount1
                                                xlwks.cells(Input_index,j) = xlwks1.cells(JEP_HPOO_Template_Index,j)
                                next
                                Input_index = Input_index + 1
                                'STUFF Jep_Hpoo_WinScp_MatchTemplatesInWinScp
                                Jep_Hpoo_WinScp_MatchTemplatesInWinScp_Index = xlwsS.cells(26,6)
                                for j = 1 to columncount1
                                                xlwks.cells(Input_index,j) = xlwks1.cells(Jep_Hpoo_WinScp_MatchTemplatesInWinScp_Index,j)
                                next
                                Input_index = Input_index + 1
                                'STUFF Jep_Hpoo_WinScp
                                Jep_Hpoo_WinSc_Index = xlwsS.cells(27,6)
                                for j = 1 to columncount1
                                                xlwks.cells(Input_index,j) = xlwks1.cells(Jep_Hpoo_WinSc_Index,j)
                                next
                                Input_index = Input_index + 1
                                'STUFF Jep_Hpoo_MatchMOPParameters
                                Jep_Hpoo_MatchMOPParameters_Index = xlwsS.cells(28,6)
                                for j = 1 to columncount1
                                                xlwks.cells(Input_index,j) = xlwks1.cells(Jep_Hpoo_MatchMOPParameters_Index,j)
                                next
                                Input_index = Input_index + 1
                                'STUFF Post_HPOO
                                Post_HPOO_index = xlwsS.cells(29,6)
                                xlwks.cells(Input_index,1) = xlwks1.cells(Post_HPOO_index,1)
                                xlwks.cells(Input_index,2) = xlwks1.cells(Post_HPOO_index,2)
                                xlwks.cells(Input_index,3) = xlwks1.cells(Post_HPOO_index,3)
                                Input_index = Input_index + 1
                                
                                

                end if
Next

'Clearing Next Row
for j = 1 to columncount
                xlwks.cells(Input_index,j) = ""
next


xlwk.Save
xlwk.Close


xlwkS.Save
xlwkS.Close

msgbox "Done"
                
