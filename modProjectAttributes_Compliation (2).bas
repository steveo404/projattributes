Attribute VB_Name = "modProjectAttributes_Compliation (2)"
Option Compare Database
Option Explicit

Sub MasterAttributesTable()
'Script Name:       MasterAttributesTable
'Author:            Steve O'Neal
'Created:           2/16/2015
'Last Modified:
'Version:           1.0
'Dependency:        NONE
'
'Script used to combine the 24 attributes for evaluating projects into single file
'Source data from ProjectAttributes module and the generated tables
'First it removes the data from the previous master Attributes_Results table and then repopulates

'Updates:


Dim strSQL As String
Dim db As Database

Set db = CurrentDb()



strSQL = "UPDATE [Attribute_Results] "
strSQL = strSQL + "SET [Attribute_Results].[Con Fee Type]= NULL, "             '#1
strSQL = strSQL + "[Attribute_Results].[Current_YTD_ER_Ratio] = NULL, "        '#2
strSQL = strSQL + "[Attribute_Results].[Prior_YTD_ER_Ratio] = NULL, "          '#3
strSQL = strSQL + "[Attribute_Results].[Delinquent_Billing] = NULL, "          '#4
strSQL = strSQL + "[Attribute_Results].[AR_Over_90_Days] = NULL, "             '#5
strSQL = strSQL + "[Attribute_Results].[Work_At_Risk] = NULL, "                '#6
strSQL = strSQL + "[Attribute_Results].[Ratio_BudtoGross] = NULL, "            '#7
strSQL = strSQL + "[Attribute_Results].[Change_Order_Amount] = NULL, "         '#8
strSQL = strSQL + "[Attribute_Results].[Change_Order_Count] = NULL, "          '#9
strSQL = strSQL + "[Attribute_Results].[Percent_Complete] = NULL, "            '#10
strSQL = strSQL + "[Attribute_Results].[Project_Mgr_Tenure] = NULL, "          '#11
strSQL = strSQL + "[Attribute_Results].[PM_3yr_GM] = NULL, "                   '#13
strSQL = strSQL + "[Attribute_Results].[CountPMs] = NULL, "                    '#14
strSQL = strSQL + "[Attribute_Results].[CountOfTask Ofc] = NULL, "             '#15
strSQL = strSQL + "[Attribute_Results].[Turnover_YTD] = NULL, "                '#16
strSQL = strSQL + "[Attribute_Results].[Delivery Method] = NULL, "             '#17
strSQL = strSQL + "[Attribute_Results].[Program_Management] = NULL, "          '#18
strSQL = strSQL + "[Attribute_Results].[Net_Retain] = NULL, "                  '#19
strSQL = strSQL + "[Attribute_Results].[DB_ChangeOrder] = NULL, "              '#20
strSQL = strSQL + "[Attribute_Results].[Client_Create_Value] = NULL, "         '#21
strSQL = strSQL + "[Attribute_Results].[DM] = NULL, "                          '#22
strSQL = strSQL + "[Attribute_Results].[YTDLabor] = NULL, "                    '#23
strSQL = strSQL + "[Attribute_Results].[RemainRev] = NULL, "                   '#24
strSQL = strSQL + "[Attribute_Results].[SegPctComplete] = NULL, "              '#24
strSQL = strSQL + "[Attribute_Results].[ClosedJob] = NULL "                    '#25

db.Execute (strSQL)
strSQL = ""


strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#1_FeeType ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#1_FeeType].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Con Fee Type] = IIF([tbl_#1_FeeType].[Con Fee Type] is NULL, ""N/A"",[tbl_#1_FeeType].[Con Fee Type]) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#2_RatioER ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#2_RatioER].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Current_YTD_ER_Ratio] = IIF([tbl_#2_RatioER].[Ratio_YTDtoCon_ER] is NULL, 0,[tbl_#2_RatioER].[Ratio_YTDtoCon_ER]) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#3_PriorYear_RatioER ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#3_PriorYear_RatioER].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Prior_YTD_ER_Ratio] = [tbl_#3_PriorYear_RatioER].[Ratio_YTDtoCon_ER] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#4_DelBilling ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#4_DelBilling].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Delinquent_Billing] = IIF([tbl_#4_DelBilling].[SumOfTotal 04/24/15 Delinquent] is NULL, 0, [tbl_#4_DelBilling].[SumOfTotal 04/24/15 Delinquent]) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#5_ARover90days ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#5_ARover90days].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[AR_Over_90_Days] = IIF([tbl_#5_ARover90days].AR_over_90days is NULL, 0,Round([tbl_#5_ARover90days].AR_over_90days,2)) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#6_WAR ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#6_WAR].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Work_At_Risk] = IIF([tbl_#6_WAR].[SumOfWar] is NULL, 0, Round([tbl_#6_WAR].[SumOfWar],2)) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#7_RatioBudtoGross ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#7_RatioBudtoGross].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Ratio_BudtoGross] = IIf([tbl_#7_RatioBudtoGross].[Ratio_BudtoGross] is NULL, 0, [tbl_#7_RatioBudtoGross].[Ratio_BudtoGross]) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#8_ChangeOrder ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#8_ChangeOrder].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Change_Order_Amount] = IIF([tbl_#8_ChangeOrder].[SumOfBudgeted Fee Total] is NULL, 0, Round([tbl_#8_ChangeOrder].[SumOfBudgeted Fee Total],2)) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#9_ChangeOrderCount ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#9_ChangeOrderCount].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Change_Order_Count] = IIF([tbl_#9_ChangeOrderCount].CountOfJob is NULL, 0, Round([tbl_#9_ChangeOrderCount].CountOfJob, 2)) "

db.Execute (strSQL)
strSQL = ""
strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#10_PercComplete ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#10_PercComplete].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Percent_Complete] = [tbl_#10_PercComplete].[Percent_Complete] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#11_PMTenure ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#11_PMTenure].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Project_Mgr_Tenure] = [tbl_#11_PMTenure].[Tenure] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#13_PMPerformance ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#13_PMPerformance].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[PM_3yr_GM] = IIF([tbl_#13_PMPerformance].[PM_3yr_GM] is NULL, 0, [tbl_#13_PMPerformance].[PM_3yr_GM]) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#14_CountPMs ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#14_CountPMs].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[CountPMs] = [tbl_#14_CountPMs].[Total] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#15_TaskOfcCount ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#15_TaskOfcCount].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[CountOfTask Ofc] = [tbl_#15_TaskOfcCount].[CountOfJob] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#16_Turnover ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#16_Turnover].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Turnover_YTD] = IIF([tbl_#16_Turnover].[YTD Total %] is NULL, 0, Round([tbl_#16_Turnover].[YTD Total %], 3)) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#17_DeliveryMethod ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#17_DeliveryMethod].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Delivery Method] = [tbl_#17_DeliveryMethod].[Delivery_Method] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#18_ProgMgmt ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#18_ProgMgmt].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Program_Management] = [tbl_#18_ProgMgmt].[Program_Management] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#19_ClientRetain ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#19_ClientRetain].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Net_Retain] = [tbl_#19_ClientRetain].[SumOfNet Retain] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results " '#20
strSQL = strSQL + "SET [Attribute_Results].[DB_ChangeOrder] = IIf(([Delivery Method]=""DB"" Or [Delivery Method]=""DBF"" "
strSQL = strSQL + "Or [Delivery Method]=""DBOM"" Or [Delivery Method]=""DBFOM"") And "
strSQL = strSQL + "[Percent_Complete] Between 0.3 And 0.6,1,0) "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#21_ClientCreate ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#21_ClientCreate].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[Client_Create_Value] = [tbl_#21_ClientCreate].[Client_Create_Value] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#22_DeliveryMeth ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#22_DeliveryMeth].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[DM] = [tbl_#22_DeliveryMeth].[DM] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#23_YTDLabor ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#23_YTDLabor].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[YTDLabor] = [tbl_#23_YTDLabor].[YTD_Labor_Value] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#24_AWNCSeg ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#24_AWNCSeg].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[RemainRev] = [tbl_#24_AWNCSeg].[Remain], "
strSQL = strSQL + "[Attribute_Results].[SegPctComplete] = [tbl_#24_AWNCSeg].[SegPctComp] "

db.Execute (strSQL)
strSQL = ""

strSQL = "UPDATE  Attribute_Results "
strSQL = strSQL + "LEFT JOIN tbl_#25_ClosedJob ON "
strSQL = strSQL + "[Attribute_Results].[Job] = [tbl_#25_ClosedJob].[Job] "
strSQL = strSQL + "SET [Attribute_Results].[ClosedJob] = IIF([tbl_#25_ClosedJob].[Closed] is NULL, 0, [tbl_#25_ClosedJob].[Closed])  "

db.Execute (strSQL)
strSQL = ""

Application.RefreshDatabaseWindow

MsgBox ("Attribute_Results compliation complete")

End Sub
