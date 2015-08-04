Attribute VB_Name = "modProjectAttribues_FinalScore (4)"
Option Compare Database

Sub ProjectAttributes_FinalScoring()
'Script Name:       ProjectAttributes_FinalScoring
'Author:            Steve O'Neal
'Created:           4/15/2015
'Last Modified:
'Version:           1.0
'Dependency:        TableExists function in the <functions> module
'
'Script used to

'Updates:   1.1

    Dim db As Database
    Dim strSQL As String
    Dim tableName As String
    
    Set db = CurrentDb
    
    tableName = "FINAL SCORING"
    
    If TableExists(tableName) Then
        DoCmd.DeleteObject acTable, tableName
    End If
    
    strSQL = "SELECT "
    strSQL = strSQL + "Attribute_Results.Job, "
    strSQL = strSQL + "Val(FeeType([Con Fee Type])) AS FeeType, "
    strSQL = strSQL + "C_ER([Current_YTD_ER_Ratio]) AS CurER, "
    strSQL = strSQL + "P_ER([Prior_YTD_ER_Ratio]) AS PriorER, "
    strSQL = strSQL + "DelB([Delinquent_Billing]) AS DelBill, "
    strSQL = strSQL + "AR90([AR_Over_90_Days]) AS AR90, "
    strSQL = strSQL + "WAR([Work_At_Risk]) AS War, "
    strSQL = strSQL + "RatioRev([Ratio_BudtoGross]) AS RatioRev, "
    strSQL = strSQL + "ChgOrd([Change_Order_Amount]) AS ChngOrd, "
    strSQL = strSQL + "ChgCnt([Change_Order_Count]) AS ChngCnt, "
    strSQL = strSQL + "PMTen([Project_Mgr_Tenure]) AS PMTen, "
    strSQL = strSQL + "PMPer([PM_3yr_GM]) AS PMPer, "
    strSQL = strSQL + "CountPM([CountPMs]) AS CntPMs, "
    strSQL = strSQL + "CountTask([CountOfTask Ofc]) AS CountTask, "
    strSQL = strSQL + "EmpTurn([Turnover_YTD]) AS EmpTurn, "
    strSQL = strSQL + "Val(MomDelMeth([Delivery Method])) AS MomDelMeth, "
    strSQL = strSQL + "Val(ProgMgmt([Program_Management])) AS ProgMgmt, "
    strSQL = strSQL + "ClientRt([Net_Retain]) AS ClientRt, "
    strSQL = strSQL + "DBChangeOrd([DB_ChangeOrder]) AS DBChange, "
    strSQL = strSQL + "NewClient([Client_Create_Value]) AS NewClient, "
    strSQL = strSQL + "Val(DelMeth([DM])) AS Delivery_Method, "
    strSQL = strSQL + "Attribute_Results.RemainRev, "
    strSQL = strSQL + "Attribute_Results.SegPctComplete, "
    strSQL = strSQL + "Attribute_Results.ClosedJob "
    strSQL = strSQL + "INTO [FINAL SCORING] "
    strSQL = strSQL + "FROM Attribute_Results "
    
    db.Execute (strSQL)
    strSQL = ""
    
    Call AttributeSort
    
    tableName = "FINAL SCORING_COMPOSITE"
    
    If TableExists(tableName) Then
        DoCmd.DeleteObject acTable, tableName
    End If
    
    strSQL = "SELECT [FINAL SCORING].Job, "
    strSQL = strSQL + "[DQ - Earnings by Job - 20150327].[Job Title], "
    strSQL = strSQL + "[FINAL SCORING].FeeType as [Fee Type], "
    strSQL = strSQL + "[FINAL SCORING].CurER as [Current YTD ER Ratio], "
    strSQL = strSQL + "[FINAL SCORING].PriorER as [Prior YTD ER Ratio], "
    strSQL = strSQL + "[FINAL SCORING].DelBill as [Delinquent Billing], "
    strSQL = strSQL + "[FINAL SCORING].AR90 as [AR Over 90 Days], "
    strSQL = strSQL + "[FINAL SCORING].War as [Work At Risk], "
    strSQL = strSQL + "[FINAL SCORING].RatioRev as [Ratio Budget to Gross], "
    strSQL = strSQL + "[FINAL SCORING].ChngOrd as [Change Order Amount], "
    strSQL = strSQL + "[FINAL SCORING].ChngCnt as [Change Order Count], "
    strSQL = strSQL + "[FINAL SCORING].PMTen as [Project Mgr Tenure], "
    strSQL = strSQL + "[FINAL SCORING].PMPer as [PM 3 Year GM], "
    strSQL = strSQL + "[FINAL SCORING].CntPMs as [Count of PMs], "
    strSQL = strSQL + "[FINAL SCORING].CountTask as [Count of Task Offices], "
    strSQL = strSQL + "[FINAL SCORING].EmpTurn as [Employee Turnover], "
    strSQL = strSQL + "[FINAL SCORING].MomDelMeth as [Delivery Method], "
    strSQL = strSQL + "[FINAL SCORING].ProgMgmt as [Program Management], "
    strSQL = strSQL + "[FINAL SCORING].ClientRt as [Client Net Retainage], "
    strSQL = strSQL + "[FINAL SCORING].DBChange as [DB Change Order], "
    strSQL = strSQL + "[FINAL SCORING].NewClient as [New Client], "
    strSQL = strSQL + "[FINAL SCORING].Delivery_Method as [Del Method], "
    strSQL = strSQL + "[FINAL SCORING].RemainRev as [Remaining Revenue], "
    strSQL = strSQL + "[FINAL SCORING].SegPctComplete as [Percent Complete], "
    strSQL = strSQL + "[FINAL SCORING].ClosedJob as [Closed Job], "
    strSQL = strSQL + "Round([FeeType]+[CurER]+[PriorER]+[DelBill]+[AR90]+[War]+[RatioRev]+ "
    strSQL = strSQL + "[ChngOrd]+[ChngCnt]+[PMTen]+[PMPer]+[CntPMs]+[CountTask]+[EmpTurn]+ "
    strSQL = strSQL + "[MomDelMeth]+[ProgMgmt]+[ClientRt]+[DBChange]+[NewClient]+[Delivery_Method],2) AS [Composite Score] "
    strSQL = strSQL + "INTO [FINAL SCORING_COMPOSITE REPORT]"
    strSQL = strSQL + "FROM [FINAL SCORING] "
    strSQL = strSQL + "INNER JOIN [DQ - Earnings by Job - 20150327] ON [FINAL SCORING].Job = [DQ - Earnings by Job - 20150327].Job "
    
    db.Execute (strSQL)
    
    Application.RefreshDatabaseWindow
    MsgBox ("Script complete!")
        
End Sub
