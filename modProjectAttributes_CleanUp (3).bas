Attribute VB_Name = "modProjectAttributes_CleanUp (3)"
Option Compare Database
Option Explicit

Sub ProjectClean()
'Script Name:       ProjectClean
'Author:            Steve O'Neal
'Created:           2/12/2015
'Last Modified:
'Version:           1.0
'Dependency:        TableExists function to determine if tables exist before deleting
'
'Delete attribute tables that are no longer needed

If TableExists("tbl_#1_FeeType") Then
    DoCmd.DeleteObject acTable, "tbl_#1_FeeType"
End If

'Attribute #2 Ratio YTD to Con ER-
If TableExists("tbl_#2_RatioER") Then
    DoCmd.DeleteObject acTable, "tbl_#2_RatioER"
End If

'Attribute #3 Prior Year Ratio YTD to Con ER-
If TableExists("tbl_#3_PriorYear_RatioER") Then
    DoCmd.DeleteObject acTable, "tbl_#3_PriorYear_RatioER"
End If

'Attribute #4 Delinquent Billing-
If TableExists("tbl_#4_DelBilling") Then
    DoCmd.DeleteObject acTable, "tbl_#4_DelBilling"
End If

'Attribute #5 AR over 90 Days-
If TableExists("tbl_#5_ARover90days") Then
    DoCmd.DeleteObject acTable, "tbl_#5_ARover90days"
End If

'Attribute #6 Work At Risk-
If TableExists("tbl_#6_WAR") Then
    DoCmd.DeleteObject acTable, "tbl_#6_WAR"
End If


'Attribute #7 Budget to Gross-
If TableExists("tbl_#7_RatioBudtoGross") Then
    DoCmd.DeleteObject acTable, "tbl_#7_RatioBudtoGross"
End If


'Attribute #8 Change Order Amount-
If TableExists("tbl_#8_ChangeOrder") Then
    DoCmd.DeleteObject acTable, "tbl_#8_ChangeOrder"
End If

'Attribute #9 Count of Change Orders-
If TableExists("tbl_#9_ChangeOrderCount") Then
    DoCmd.DeleteObject acTable, "tbl_#9_ChangeOrderCount"
End If

'Attribute #10 Percent Complete-
If TableExists("tbl_#10_PercComplete") Then
    DoCmd.DeleteObject acTable, "tbl_#10_PercComplete"
End If

'Attribute #11-
If TableExists("tbl_#11_PMTenure") Then
    DoCmd.DeleteObject acTable, "tbl_#11_PMTenure"
End If

'Attribute #13 part 1-
If TableExists("temptbl_#13_part_1") Then
    DoCmd.DeleteObject acTable, "temptbl_#13_part_1"
End If

'Attribute #13 part 2
If TableExists("tbl_#13_PMPerformance") Then
    DoCmd.DeleteObject acTable, "tbl_#13_PMPerformance"
End If

'Attribute #14-
If TableExists("tbl_#14_CountPMs") Then
    DoCmd.DeleteObject acTable, "tbl_#14_CountPMs"
End If

'Attribute #15-
If TableExists("tbl_#15_TaskOfcCount") Then
    DoCmd.DeleteObject acTable, "tbl_#15_TaskOfcCount"
End If

'Attribute #16-
If TableExists("tbl_#16_Turnover") Then
    DoCmd.DeleteObject acTable, "tbl_#16_Turnover"
End If

'Attribute #17 part 1-
If TableExists("temptbl_#17_part_1") Then
    DoCmd.DeleteObject acTable, "temptbl_#17_part_1"
End If

'Attribute #17 part 2
If TableExists("tbl_#17_DeliveryMethod") Then
    DoCmd.DeleteObject acTable, "tbl_#17_DeliveryMethod"
End If


'Attribute #18 part 1-
If TableExists("temptbl_#18_part_1") Then
    DoCmd.DeleteObject acTable, "temptbl_#18_part_1"
End If

'Attribute #18 part 2
If TableExists("tbl_#18_ProgMgmt") Then
    DoCmd.DeleteObject acTable, "tbl_#18_ProgMgmt"
End If

'Attribute #19-
If TableExists("tbl_#19_ClientRetain") Then
    DoCmd.DeleteObject acTable, "tbl_#19_ClientRetain"
End If

'Attribute #21 Client Creation Date-
If TableExists("tbl_#21_ClientCreate") Then
    DoCmd.DeleteObject acTable, "tbl_#21_ClientCreate"
End If

'Attribute #22 DeliveryMethod
If TableExists("tbl_#22_DeliveryMeth") Then
    DoCmd.DeleteObject acTable, "tbl_#22_DeliveryMeth"
End If

'Attribute #23 YTD Labor-
If TableExists("tbl_#23_YTDLabor") Then
    DoCmd.DeleteObject acTable, "tbl_#23_YTDLabor"
End If

'Attribute #24 AW_NC Seg-
If TableExists("tbl_#24_AWNCSeg") Then
    DoCmd.DeleteObject acTable, "tbl_#24_AWNCSeg"
End If

'Attribute #25 Closed Job
If TableExists("tbl_#25_ClosedJob") Then
    DoCmd.DeleteObject acTable, "tbl_#25_ClosedJob"
End If

Application.RefreshDatabaseWindow

MsgBox ("CLEAN UP complete")

End Sub
