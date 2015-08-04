Attribute VB_Name = "modProjectAttributes (1)"
Option Compare Database
Option Explicit

Sub AttributeCalculation()
'Script Name:       AttributeCalculation
'Author:            Steve O'Neal
'Created:           1/27/2015
'Last Modified:     4/6/2015
'Version:           1.7
'Dependency:        TableExists function in the <functions> module
'
'Script used to calculate the 24 attributes for evaluating projects
'Source data is loaded into the database using the modImportXlsxfiles
'The script will generate a table for each attribute.  For some attributes, temporary
'tables are created and then deleted once the script has completed.
'Once complete, the script will create a final Attribute_Results table with all
'18 attributes matched to the in scope files.

'Updates:   1.1 Change the MASTER List to use AW and NC jobs with positive YTD Labor
'           1.2 Add additional attributes and update MASTER list to pull all jobs
'           1.3 Removed MASTER compliation and moved to separate module
'           1.4 Added table assignment section - easier table updates - updated all code
'           1.5 Updated for latest data pull
'           1.6 Updated for latest data pull - Q1 results
'           1.7 Added tblDataSources reference for easier updating of data sources
            

Dim strSQL As String
Dim db As Database

'********************************TABLE ASSIGNMENT************************************
Dim sourceTermsByOffice As String       'Termination data - typically from HR/Don Bender
Dim sourceARATBByJob As String          'DQ of AR ATB by Job
Dim sourceARReservesRetain As String    'DQ of AR Reservers Retainaing by Segment
Dim sourceDQEarnByJob12 As String       'DQ of Earnings by Job 3 years ago
Dim sourceDQEarnByJob13 As String       'DQ of Earnings by Job 2 years ago
Dim sourceDQEarnByJob14 As String       'DQ of Earnings by Job current year
Dim sourceEarnBySegment As String       'DQ of Earnings by Segment
Dim sourceEarnBySegOffice As String     'DQ of Earnings by Segment Office
Dim sourceFirmwideList As String        'DQ Firmwide Listing
Dim sourceJIPClosed As String           'DQ JIP Closed Jobs
Dim sourceJIPSegment As String          'DQ JIP Segment
Dim sourcePursuitRpts As String         'DQ Projections Pursuits Reports
Dim sourceJCClientCreate As String      'Client Create data - from Jeanie Chicoine
Dim sourceJCClientName As String        'Client Name Association - from Jeanie Chicoine
Dim sourceJCPMPerJob As String          'PMs per Job - from Jeanie Chicoine

'Use the tblDataSources to populate the list of tables used in this query
sourceTermsByOffice = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 1")) & "]"
sourceARATBByJob = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 2")) & "]"
sourceARReservesRetain = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 3")) & "]"
sourceDQEarnByJob12 = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 4")) & "]"
sourceDQEarnByJob13 = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 5")) & "]"
sourceDQEarnByJob14 = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 6")) & "]"
sourceEarnBySegment = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 7")) & "]"
sourceEarnBySegOffice = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 8")) & "]"
sourceFirmwideList = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 9")) & "]"
sourceJIPClosed = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 10")) & "]"
sourceJIPSegment = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 11")) & "]"
sourcePursuitRpts = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 12")) & "]"
sourceJCClientCreate = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 13")) & "]"
sourceJCClientName = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 14")) & "]"
sourceJCPMPerJob = "[" & (DLookup("[Current_Name]", "tblDataSources", "[ID] = 15")) & "]"

'************************************************************************************

Set db = CurrentDb()

strSQL = ""
strSQL = strSQL + ""

'Attribute #1 Fee Type-------------------------------------------------------------------------
'Script uses the DQ Earnings by Segment report to determine the Fee Type

'Delete old table if it still exists in the database
If TableExists("tbl_#1_FeeType") Then
    DoCmd.DeleteObject acTable, "tbl_#1_FeeType"
End If

strSQL = "SELECT B.Job, B.[Con Fee Type], B.[SumOfTotal Segment Budget] "
strSQL = strSQL + "INTO tbl_#1_FeeType "
strSQL = strSQL + "FROM (SELECT "
strSQL = strSQL + "    A.Job,"
strSQL = strSQL + "    A.[Con Fee Type], "
strSQL = strSQL + "    Sum(A.[Total Segment Budget]) As [SumOfTotal Segment Budget] "
strSQL = strSQL + "    FROM " & sourceEarnBySegment & " as A "
strSQL = strSQL + "    GROUP BY A.Job, "
strSQL = strSQL + "    A.[Con Fee Type] "
strSQL = strSQL + "    ORDER BY A.Job, "
strSQL = strSQL + "    Sum (A.[Total Segment Budget]) "
strSQL = strSQL + ""
strSQL = strSQL + ")  AS B INNER JOIN (SELECT D.Job, "
strSQL = strSQL + "    Max(D.[SumOfTotal Segment Budget]) As [MaxOfSumOfTotal Segment Budget] "
strSQL = strSQL + "    FROM "
strSQL = strSQL + "    ( "
strSQL = strSQL + "        SELECT"
strSQL = strSQL + "        C.Job,"
strSQL = strSQL + "        C.[Con Fee Type], "
strSQL = strSQL + "        Sum(C.[Total Segment Budget]) As [SumOfTotal Segment Budget] "
strSQL = strSQL + "        FROM " & sourceEarnBySegment & " as C"
strSQL = strSQL + "        GROUP BY C.Job, "
strSQL = strSQL + "        C.[Con Fee Type] "
strSQL = strSQL + "        ORDER BY C.Job, "
strSQL = strSQL + "        Sum (C.[Total Segment Budget]) "
strSQL = strSQL + "        ) as D "
strSQL = strSQL + "    GROUP BY D.Job "
strSQL = strSQL + ")  AS S ON (B.Job = S.Job) AND (B.[SumOfTotal Segment Budget] = S.[MaxOfSumOfTotal Segment Budget]) "

db.Execute (strSQL)
strSQL = ""

'Attribute #2 Ratio YTD to Con ER-------------------------------------------------------------------------
If TableExists("tbl_#2_RatioER") Then
    DoCmd.DeleteObject acTable, "tbl_#2_RatioER"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.[YTD ER], "
strSQL = strSQL + "A.[Con ER], "
strSQL = strSQL + "IIF([Con ER]=0,0,Round(Abs([YTD ER]/[Con ER]-1),3)) AS Ratio_YTDtoCon_ER "
strSQL = strSQL + "INTO tbl_#2_RatioER "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #3 Prior Year Ratio YTD to Con ER-------------------------------------------------------------------------
If TableExists("tbl_#3_PriorYear_RatioER") Then
    DoCmd.DeleteObject acTable, "tbl_#3_PriorYear_RatioER"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.[YTD ER], "
strSQL = strSQL + "A.[Con ER], "
strSQL = strSQL + "IIF([Con ER]=0,0, Round(Abs([YTD ER]/[Con ER]-1),3)) AS Ratio_YTDtoCon_ER "
strSQL = strSQL + "INTO tbl_#3_PriorYear_RatioER "
strSQL = strSQL + "FROM " & sourceDQEarnByJob13 & " As A "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #4 Delinquent Billing-------------------------------------------------------------------------
If TableExists("tbl_#4_DelBilling") Then
    DoCmd.DeleteObject acTable, "tbl_#4_DelBilling"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "Sum(A.[Total 04/24/15 Delinquent]) AS [SumOfTotal 04/24/15 Delinquent] "
strSQL = strSQL + "INTO tbl_#4_DelBilling "
strSQL = strSQL + "FROM " & sourceJIPSegment & " As A "
strSQL = strSQL + "GROUP BY A.Job  "

db.Execute (strSQL)
strSQL = ""

'Attribute #5 AR over 90 Days-------------------------------------------------------------------------
If TableExists("tbl_#5_ARover90days") Then
    DoCmd.DeleteObject acTable, "tbl_#5_ARover90days"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.[91-180 Days], "
strSQL = strSQL + "A.[Over 180 Days], "
strSQL = strSQL + "Abs(Round([91-180 Days]+[Over 180 Days],2)) AS AR_over_90days "
strSQL = strSQL + "INTO tbl_#5_ARover90days "
strSQL = strSQL + "FROM " & sourceARATBByJob & "As A "

db.Execute (strSQL)
strSQL = ""

'Attribute #6 Work At Risk-------------------------------------------------------------------------
If TableExists("tbl_#6_WAR") Then
    DoCmd.DeleteObject acTable, "tbl_#6_WAR"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "Sum(A.[Uncontracted War]+ A.[Unauthorized War]) AS [SumOfWar] "
strSQL = strSQL + "INTO tbl_#6_WAR "
strSQL = strSQL + "FROM " & sourceEarnBySegment & " AS A "
strSQL = strSQL + "GROUP BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #7 Budget to Gross-------------------------------------------------------------------------
If TableExists("tbl_#7_RatioBudtoGross") Then
    DoCmd.DeleteObject acTable, "tbl_#7_RatioBudtoGross"
End If

strSQL = "SELECT "
strSQL = strSQL + "T.Job, "
strSQL = strSQL + "T.[SumOfBudgeted Base Expense], "
strSQL = strSQL + "T.[SumOfBudgeted Fee Total], "
strSQL = strSQL + "IIf([SumOfBudgeted Fee Total]=0,-1,Round(([SumOfBudgeted Fee Total]-[SumOfBudgeted Base Expense])/[SumOfBudgeted Fee Total],3)) AS Ratio_BudtoGross "
strSQL = strSQL + "INTO tbl_#7_RatioBudtoGross "
strSQL = strSQL + "FROM "
strSQL = strSQL + "(SELECT "
strSQL = strSQL + "S.Job, "
strSQL = strSQL + "Sum(S.[Budgeted Base Expense]) AS [SumOfBudgeted Base Expense], "
strSQL = strSQL + "Sum(S.[Budgeted Fee Total]) AS [SumOfBudgeted Fee Total] "
strSQL = strSQL + "FROM " & sourceEarnBySegment & " AS S GROUP BY S.Job "
strSQL = strSQL + ")  AS T "

db.Execute (strSQL)
strSQL = ""

'Attribute #8 Change Order Amount-------------------------------------------------------------------------
If TableExists("tbl_#8_ChangeOrder") Then
    DoCmd.DeleteObject acTable, "tbl_#8_ChangeOrder"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.Ph, "
strSQL = strSQL + "Sum(A.[Budgeted Fee Total]) AS [SumOfBudgeted Fee Total] "
strSQL = strSQL + "INTO tbl_#8_ChangeOrder "
strSQL = strSQL + "FROM " & sourceEarnBySegment & " As A "
strSQL = strSQL + "GROUP BY "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.Ph "
strSQL = strSQL + "HAVING (((A.Ph)=""XW"")) "

db.Execute (strSQL)
strSQL = ""

'Attribute #9 Count of Change Orders-------------------------------------------------------------------------
If TableExists("tbl_#9_ChangeOrderCount") Then
    DoCmd.DeleteObject acTable, "tbl_#9_ChangeOrderCount"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "A.Ph, "
strSQL = strSQL + "Count(A.Job) AS CountOfJob "
strSQL = strSQL + "INTO tbl_#9_ChangeOrderCount "
strSQL = strSQL + "FROM " & sourceEarnBySegment & " As A "
strSQL = strSQL + "GROUP BY A.Job, "
strSQL = strSQL + "A.Ph "
strSQL = strSQL + "HAVING (((A.Ph)=""XW"")) "

db.Execute (strSQL)
strSQL = ""

'Attribute #10 Percent Complete-------------------------------------------------------------------------
If TableExists("tbl_#10_PercComplete") Then
    DoCmd.DeleteObject acTable, "tbl_#10_PercComplete"
End If

strSQL = "SELECT "
strSQL = strSQL + "K.Job, "
strSQL = strSQL + "Round([SumOfTotal JTD Gross Revenue]/[SumOfEst Fee],3) AS Percent_Complete "
strSQL = strSQL + "INTO tbl_#10_PercComplete "
strSQL = strSQL + "FROM (SELECT "
strSQL = strSQL + "G.Job, "
strSQL = strSQL + "Sum(G.[Est Fee]) AS [SumOfEst Fee], "
strSQL = strSQL + "Sum(G.[Total JTD Gross Revenue]) AS [SumOfTotal JTD Gross Revenue] "
strSQL = strSQL + "FROM " & sourceJIPSegment & " As G "
strSQL = strSQL + "GROUP BY G.Job "
strSQL = strSQL + ")  AS K "

db.Execute (strSQL)
strSQL = ""

'Attribute #11 PM Tenure---------------------------------------------------------------------
If TableExists("tbl_#11_PMTenure") Then
    DoCmd.DeleteObject acTable, "tbl_#11_PMTenure"
End If

Dim lowTen As Integer
Dim highTen As Integer

lowTen = 2
highTen = 20

strSQL = "SELECT "
strSQL = strSQL + "X.Job, "
strSQL = strSQL + "X.[Project Manager], "
strSQL = strSQL + "X.EMP_Number, "
strSQL = strSQL + "Y.[HNTB Service], "
strSQL = strSQL + "IIf([HNTB Service] Between 2 And 20,0,1) AS Tenure "
strSQL = strSQL + "INTO tbl_#11_PMTenure "
strSQL = strSQL + "FROM (SELECT "
strSQL = strSQL + " A.Job, "
strSQL = strSQL + " A.[Project Manager], "
strSQL = strSQL + " Left(Right([Project Manager],6),5) AS EMP_Number "
strSQL = strSQL + " FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + " GROUP BY "
strSQL = strSQL + " A.Job, "
strSQL = strSQL + " A.[Project Manager], "
strSQL = strSQL + " Left(Right([Project Manager],6),5) "
strSQL = strSQL + ")  AS X LEFT JOIN (SELECT "
strSQL = strSQL + "B.[Employee Number], "
strSQL = strSQL + "B.[HNTB Service] "
strSQL = strSQL + "FROM " & sourceFirmwideList & " As B "
strSQL = strSQL + ")  AS Y ON X.EMP_Number = Y.[Employee Number] "
strSQL = strSQL + "ORDER BY X.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #13 part 1-------------------------------------------------------------------------
'Complex section of code that requires several queries to obtain the list of PMs and their
'calculate the performance.  A few temporary tables will be created and some deleted.

'Create a temporary list of PMs derived from appending last 3 years list of PMs

If TableExists("temp_13_MasterPMs") Then
    DoCmd.DeleteObject acTable, "temp_13_MasterPMs"
End If

strSQL = "SELECT A.[Project Manager], "
strSQL = strSQL + "Sum(A.[YTD Earned Gross Revenue]) AS [SumOfYTD Earned Gross Revenue], "
strSQL = strSQL + "Sum(A.[YTD Earnings from Operations]) AS [SumOfYTD Earnings from Operations] "
strSQL = strSQL + "INTO temp_13_MasterPMs "
strSQL = strSQL + "FROM " & sourceDQEarnByJob12 & " As A "
strSQL = strSQL + "GROUP BY A.[Project Manager] "

db.Execute (strSQL)

strSQL = "INSERT INTO temp_13_MasterPMs ( [Project Manager], [SumOfYTD Earned Gross Revenue], [SumOfYTD Earnings from Operations]) "
strSQL = strSQL + "SELECT B.[Project Manager], "
strSQL = strSQL + "Sum(B.[YTD Earned Gross Revenue]) AS [SumOfYTD Earned Gross Revenue], "
strSQL = strSQL + "Sum(B.[YTD Earnings from Operations]) AS [SumOfYTD Earnings from Operations] "
strSQL = strSQL + "FROM " & sourceDQEarnByJob13 & " As B "
strSQL = strSQL + "GROUP BY B.[Project Manager] "

db.Execute (strSQL)

strSQL = "INSERT INTO temp_13_MasterPMs ( [Project Manager], [SumOfYTD Earned Gross Revenue], [SumOfYTD Earnings from Operations] ) "
strSQL = strSQL + "SELECT C.[Project Manager], "
strSQL = strSQL + "Sum(C.[YTD Earned Gross Revenue]) AS [SumOfYTD Earned Gross Revenue], "
strSQL = strSQL + "Sum(C.[YTD Earnings from Operations]) AS [SumOfYTD Earnings from Operations] "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As C "
strSQL = strSQL + "GROUP BY C.[Project Manager] "

db.Execute (strSQL)

If TableExists("tbl_#13_MasterPMs") Then
    DoCmd.DeleteObject acTable, "tbl_#13_MasterPMs"
End If

strSQL = "SELECT temp_13_MasterPMs.[Project Manager], "
strSQL = strSQL + "Sum(temp_13_MasterPMs.[SumOfYTD Earned Gross Revenue]) AS [SumOfSumOfYTD Earned Gross Revenue], "
strSQL = strSQL + "Sum(temp_13_MasterPMs.[SumOfYTD Earnings from Operations]) AS [SumOfSumOfYTD Earnings from Operations] "
strSQL = strSQL + "INTO tbl_#13_MasterPMs "
strSQL = strSQL + "FROM temp_13_MasterPMs "
strSQL = strSQL + "GROUP BY temp_13_MasterPMs.[Project Manager] "

db.Execute (strSQL)
strSQL = ""

'Attribute #13 part 2
'Creates the final table for this attribute
If TableExists("tbl_#13_PMPerformance") Then
    DoCmd.DeleteObject acTable, "tbl_#13_PMPerformance"
End If

strSQL = "SELECT A.Job, "
strSQL = strSQL + "A.[Project Manager], "
strSQL = strSQL + "[tbl_#13_MasterPMs].[SumOfSumOfYTD Earned Gross Revenue], "
strSQL = strSQL + "[tbl_#13_MasterPMs].[SumOfSumOfYTD Earnings from Operations], "
strSQL = strSQL + "IIf([tbl_#13_MasterPMs].[SumOfSumOfYTD Earned Gross Revenue] = 0, 0, "
strSQL = strSQL + "Round([SumOfSumOfYTD Earnings from Operations]/[SumOfSumOfYTD Earned Gross Revenue],3)) AS PM_3yr_GM "
strSQL = strSQL + "INTO tbl_#13_PMPerformance "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "LEFT JOIN [tbl_#13_MasterPMs] ON "
strSQL = strSQL + "A.[Project Manager] = [tbl_#13_MasterPMs].[Project Manager] "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Temporary table is deleted
DoCmd.DeleteObject acTable, "temp_13_MasterPMs"
DoCmd.DeleteObject acTable, "tbl_#13_MasterPMs"

'Attribute #14-------------------------------------------------------------------------
If TableExists("tbl_#14_CountPMs") Then
    DoCmd.DeleteObject acTable, "tbl_#14_CountPMs"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, " & sourceJCPMPerJob & ".Total "
strSQL = strSQL + "INTO tbl_#14_CountPMs "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "INNER JOIN " & sourceJCPMPerJob & " ON "
strSQL = strSQL + "A.Job = " & sourceJCPMPerJob & ".Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #15-------------------------------------------------------------------------
If TableExists("tbl_#15_TaskOfcCount") Then
    DoCmd.DeleteObject acTable, "tbl_#15_TaskOfcCount"
End If

strSQL = "SELECT B.Job, Count(B.Job) AS CountOfJob "
strSQL = strSQL + "INTO tbl_#15_TaskOfcCount "
strSQL = strSQL + "FROM ( "
strSQL = strSQL + "SELECT A.Job, A.[Task Ofc] "
strSQL = strSQL + "FROM " & sourceEarnBySegOffice & " As A "
strSQL = strSQL + "GROUP BY A.Job, A.[Task Ofc] "
strSQL = strSQL + ") AS B "
strSQL = strSQL + "GROUP BY B.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #16-------------------------------------------------------------------------
If TableExists("tbl_#16_Turnover") Then
    DoCmd.DeleteObject acTable, "tbl_#16_Turnover"
End If

strSQL = "SELECT X.Job, X.[Task Ofc], " & sourceTermsByOffice & ".[YTD Total %] "
strSQL = strSQL + "INTO tbl_#16_Turnover "
strSQL = strSQL + "FROM ((SELECT "
strSQL = strSQL + " A.Job, "
strSQL = strSQL + " A.[Task Ofc], "
strSQL = strSQL + " Sum(A.[YTD Labor]) AS [SumOfYTD Labor] "
strSQL = strSQL + " FROM " & sourceEarnBySegOffice & "As A  "
strSQL = strSQL + " GROUP BY "
strSQL = strSQL + " A.Job, "
strSQL = strSQL + " A.[Task Ofc] "
strSQL = strSQL + " HAVING (((Sum(A.[YTD Labor]))>0)) "
strSQL = strSQL + " ORDER BY A.Job "
strSQL = strSQL + ")  AS X INNER JOIN (SELECT "
strSQL = strSQL + " Z.Job, "
strSQL = strSQL + " Max(Z.[SumOfYTD Labor]) AS [MaxOfSumOfYTD Labor] "
strSQL = strSQL + " FROM "
strSQL = strSQL + " ( "
strSQL = strSQL + "     SELECT "
strSQL = strSQL + "     B.Job, "
strSQL = strSQL + "     B.[Task Ofc], "
strSQL = strSQL + "     Sum(B.[YTD Labor]) AS [SumOfYTD Labor] "
strSQL = strSQL + "     FROM " & sourceEarnBySegOffice & "As B "
strSQL = strSQL + "     GROUP BY "
strSQL = strSQL + "     B.Job, "
strSQL = strSQL + "     B.[Task Ofc] "
strSQL = strSQL + "     HAVING (((Sum(B.[YTD Labor]))>0)) "
strSQL = strSQL + "     ORDER BY B.Job "
strSQL = strSQL + " ) As Z "
strSQL = strSQL + " GROUP BY Z.Job "
strSQL = strSQL + ")  AS Y ON (X.Job = Y.Job) AND (X.[SumOfYTD Labor] = Y.[MaxOfSumOfYTD Labor])) "
strSQL = strSQL + "INNER JOIN " & sourceTermsByOffice & " ON X.[Task Ofc] = " & sourceTermsByOffice & ".Office "

db.Execute (strSQL)
strSQL = ""

'Attribute #17 part 1-------------------------------------------------------------------------
If TableExists("temptbl_#17_part_1") Then
    DoCmd.DeleteObject acTable, "temptbl_#17_part_1"
End If

strSQL = "SELECT X.[Job #], X.[Delivery Method], X.[SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "INTO temptbl_#17_part_1 "
strSQL = strSQL + "FROM (SELECT "
strSQL = strSQL + "  A.[Job #], "
strSQL = strSQL + "  A.[Delivery Method], "
strSQL = strSQL + "  Sum(A.[Anticipated Pursuit Gross Sales]) AS [SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "  FROM " & sourcePursuitRpts & " As A "
strSQL = strSQL + "  GROUP BY A.[Job #], A.[Delivery Method] "
strSQL = strSQL + "  HAVING (((A.[Job #])<>""     "")) "
strSQL = strSQL + ")  AS X INNER JOIN (SELECT "
strSQL = strSQL + " Z.[Job #], "
strSQL = strSQL + " Max(Z.[SumOfAnticipated Pursuit Gross Sales]) AS [MaxOfSumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + " FROM "
strSQL = strSQL + "     ( "
strSQL = strSQL + "     SELECT "
strSQL = strSQL + "    B.[Job #], "
strSQL = strSQL + "    B.[Delivery Method], "
strSQL = strSQL + "    Sum(B.[Anticipated Pursuit Gross Sales]) AS [SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "    FROM " & sourcePursuitRpts & " As B "
strSQL = strSQL + "    GROUP BY B.[Job #], B.[Delivery Method] "
strSQL = strSQL + "    HAVING (((B.[Job #])<>""     "")) "
strSQL = strSQL + "     )As Z "
strSQL = strSQL + "GROUP BY Z.[Job #] "
strSQL = strSQL + ")  AS Y ON (X.[Job #] = Y.[Job #]) "
strSQL = strSQL + "AND (X.[SumOfAnticipated Pursuit Gross Sales] = Y.[MaxOfSumOfAnticipated Pursuit Gross Sales]) "
strSQL = strSQL + "ORDER BY X.[Job #] "

db.Execute (strSQL)
strSQL = ""

'Attribute #17 part 2
If TableExists("tbl_#17_DeliveryMethod") Then
    DoCmd.DeleteObject acTable, "tbl_#17_DeliveryMethod"
End If

strSQL = "SELECT A.Job, "
strSQL = strSQL + "IIF([temptbl_#17_part_1].[Delivery Method] is NULL, ""N/A"", [temptbl_#17_part_1].[Delivery Method]) AS Delivery_Method "
strSQL = strSQL + "INTO tbl_#17_DeliveryMethod "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "LEFT JOIN [temptbl_#17_part_1] ON A.Job = [temptbl_#17_part_1].[Job #] "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Temporary table is deleted
DoCmd.DeleteObject acTable, "temptbl_#17_part_1"

'Attribute #18 part 1-------------------------------------------------------------------------
If TableExists("temptbl_#18_part_1") Then
    DoCmd.DeleteObject acTable, "temptbl_#18_part_1"
End If

strSQL = "SELECT X.[Job #], "
strSQL = strSQL + "IIf(X.[Program Delivery] Is Null,""N"",""Y"") AS ProgDelivery, "
strSQL = strSQL + "X.[SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "INTO temptbl_#18_part_1 "
strSQL = strSQL + "FROM ( "
strSQL = strSQL + "SELECT A.[Job #], "
strSQL = strSQL + "A.[Program Delivery], "
strSQL = strSQL + "Sum(A.[Anticipated Pursuit Gross Sales]) AS [SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "FROM " & sourcePursuitRpts & " As A "
strSQL = strSQL + "GROUP BY A.[Job #], A.[Program Delivery] "
strSQL = strSQL + "HAVING (((A.[Job #])<>""     "")))  AS X "
strSQL = strSQL + "INNER JOIN (SELECT Z.[Job #], Max(Z.[SumOfAnticipated Pursuit Gross Sales]) AS [MaxOfSumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "FROM (SELECT B.[Job #], B.[Program Delivery], "
strSQL = strSQL + "Sum(B.[Anticipated Pursuit Gross Sales]) AS [SumOfAnticipated Pursuit Gross Sales] "
strSQL = strSQL + "FROM " & sourcePursuitRpts & " As B "
strSQL = strSQL + "GROUP BY B.[Job #], "
strSQL = strSQL + "B.[Program Delivery] "
strSQL = strSQL + "HAVING (((B.[Job #])<>""     ""))) "
strSQL = strSQL + "AS Z GROUP BY Z.[Job #])  AS Y ON (X.[Job #] = Y.[Job #]) "
strSQL = strSQL + "AND (X.[SumOfAnticipated Pursuit Gross Sales] = Y.[MaxOfSumOfAnticipated Pursuit Gross Sales]) "
strSQL = strSQL + "ORDER BY X.[Job #] "

db.Execute (strSQL)
strSQL = ""

'Attribute #18 part 2
If TableExists("tbl_#18_ProgMgmt") Then
    DoCmd.DeleteObject acTable, "tbl_#18_ProgMgmt"
End If

strSQL = "SELECT A.Job, "
strSQL = strSQL + "IIf([temptbl_#18_part_1].[ProgDelivery]=""Y"",""Y"",IIf([temptbl_#18_part_1].[ProgDelivery]=""N"",""N"",""N/A"")) AS Program_Management "
strSQL = strSQL + "INTO tbl_#18_ProgMgmt "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "LEFT JOIN [temptbl_#18_part_1] ON A.Job = [temptbl_#18_part_1].[Job #] "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""


DoCmd.DeleteObject acTable, "temptbl_#18_part_1"

'Attribute #19-------------------------------------------------------------------------
If TableExists("tbl_#19_ClientRetain") Then
    DoCmd.DeleteObject acTable, "tbl_#19_ClientRetain"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "IIF(Sum(" & sourceARReservesRetain & ".[Net Retain]) is NULL, 0, "
strSQL = strSQL + "Sum(" & sourceARReservesRetain & ".[Net Retain])) AS [SumOfNet Retain] "
strSQL = strSQL + "INTO tbl_#19_ClientRetain "
strSQL = strSQL + "FROM (" & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "LEFT JOIN " & sourceJCClientName & " ON "
strSQL = strSQL + "A.[Primary Client Name] = " & sourceJCClientName & ".[Client Name]) "
strSQL = strSQL + "LEFT JOIN " & sourceARReservesRetain
strSQL = strSQL + "ON " & sourceJCClientName & ".[Client Abbr] = " & sourceARReservesRetain & ".Client "
strSQL = strSQL + "GROUP BY A.Job "
strSQL = strSQL + "ORDER BY A.Job "


db.Execute (strSQL)
strSQL = ""

'Attribute #21 Client Creation Date-------------------------------------------------------------------------
If TableExists("tbl_#21_ClientCreate") Then
    DoCmd.DeleteObject acTable, "tbl_#21_ClientCreate"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, " & sourceJCClientCreate & ".[Client Creation Date], "
strSQL = strSQL + "IIf([Client Creation Date]<=#12/31/2011#,0,1) AS Client_Create_Value "
strSQL = strSQL + "INTO tbl_#21_ClientCreate "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "LEFT JOIN " & sourceJCClientCreate & " ON "
strSQL = strSQL + "A.[Primary Client Name] = " & sourceJCClientCreate & ".Client "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #22 DeliveryMethod--------------------------------------------------------------------------
If TableExists("tbl_#22_DeliveryMeth") Then
    DoCmd.DeleteObject acTable, "tbl_#22_DeliveryMeth"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "IIf([HNTB Advantage]='Y','DB',IIf([Field2]='CM','CM',IIf([Field2]='PM','PM','N/A'))) AS DM "
strSQL = strSQL + "INTO tbl_#22_DeliveryMeth "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A LEFT JOIN 2014 ON A.Job = [2014].Project "
    
db.Execute (strSQL)
strSQL = ""


'Attribute #23 YTD Labor-------------------------------------------------------------------------
If TableExists("tbl_#23_YTDLabor") Then
    DoCmd.DeleteObject acTable, "tbl_#23_YTDLabor"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "IIf(A.[YTD Labor]>0,1,0) AS YTD_Labor_Value "
strSQL = strSQL + "INTO tbl_#23_YTDLabor "
strSQL = strSQL + "FROM " & sourceDQEarnByJob14 & " As A "
strSQL = strSQL + "ORDER BY A.Job "

db.Execute (strSQL)
strSQL = ""

'Attribute #24 AW_NC Seg-------------------------------------------------------------------------
If TableExists("tbl_#24_AWNCSeg") Then
    DoCmd.DeleteObject acTable, "tbl_#24_AWNCSeg"
End If

strSQL = "SELECT "
strSQL = strSQL + "B.Job, "
strSQL = strSQL + "Round([SumOfSumOfEst Fee]-[SumOfSumOfTotal JTD Gross Revenue],2) AS Remain, "
strSQL = strSQL + "IIF([SumOfSumOfEst Fee]=0,0,Round([SumOfSumOfTotal JTD Gross Revenue]/[SumOfSumOfEst Fee],3)) AS SegPctComp "
strSQL = strSQL + "INTO tbl_#24_AWNCSeg "
strSQL = strSQL + "FROM "
strSQL = strSQL + "( "
strSQL = strSQL + "SELECT "
strSQL = strSQL + " A.Job, "
strSQL = strSQL + " Sum(A.[SumOfEst Fee]) AS [SumOfSumOfEst Fee], "
strSQL = strSQL + " Sum(A.[SumOfTotal JTD Gross Revenue]) AS [SumOfSumOfTotal JTD Gross Revenue] "
strSQL = strSQL + " FROM "
strSQL = strSQL + " ( "
strSQL = strSQL + "     SELECT "
strSQL = strSQL + "     C.Job, "
strSQL = strSQL + "     Sum(C.[Est Fee]) AS [SumOfEst Fee], "
strSQL = strSQL + "     Sum(C.[Total JTD Gross Revenue]) AS [SumOfTotal JTD Gross Revenue] "
strSQL = strSQL + "     FROM " & sourceJIPSegment & " As C "
strSQL = strSQL + "     GROUP BY "
strSQL = strSQL + "     C.Job, "
strSQL = strSQL + "     C.[Seg Status] "
strSQL = strSQL + "     HAVING "
strSQL = strSQL + "     (((C.[Seg Status])=""AW"" Or "
strSQL = strSQL + "     (C.[Seg Status])=""NC"")) "
strSQL = strSQL + " ) as A "
strSQL = strSQL + " GROUP BY A.Job  "
strSQL = strSQL + ") As B "

db.Execute (strSQL)
strSQL = ""

'Attribute #25 Closed Job------------------------------------------------------------------------
If TableExists("tbl_#25_ClosedJob") Then
    DoCmd.DeleteObject acTable, "tbl_#25_ClosedJob"
End If

strSQL = "SELECT "
strSQL = strSQL + "A.Job, "
strSQL = strSQL + "1 AS Closed "
strSQL = strSQL + "INTO tbl_#25_ClosedJob "
strSQL = strSQL + "FROM " & sourceJIPClosed & " As A "

db.Execute (strSQL)
strSQL = ""

'Create the Master Job's list for the current list of active projects--------------------------
'This is the list of active jobs that will be matched against the 25 attributes
If TableExists("MASTER_Job_List") Then
    DoCmd.DeleteObject acTable, "MASTER_Job_List"
End If

strSQL = "SELECT A.Job "
strSQL = strSQL + "INTO MASTER_Job_List "
strSQL = strSQL + "FROM " & sourceJIPSegment & " As A "
strSQL = strSQL + "GROUP BY A.Job "

db.Execute (strSQL)
strSQL = ""

Application.RefreshDatabaseWindow

'Call MasterAttributesTable
'Call ProjectClean

MsgBox ("Script complete - Attribute_Results populated")

End Sub
