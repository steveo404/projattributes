Attribute VB_Name = "modOrderAttributes"
Option Compare Database

Sub AttributeSort()
'Script Name:       Attribute Sort
'Author:            Steve O'Neal
'Created:           4/9/2015
'Last Modified:     4/15/2015
'Version:           1.1
'Dependency:        NONE
'
'Script used to identify the top 3 attributes that contribute to the
'composite score.  If there is a tie, the first attribute from the left
'of the table is selected.
'
'Updates:           1.1 Added query to delete all attributes from prior run of script

    Dim db As Database
    Dim rs As Recordset
    Dim jobNumber As String
    Dim fieldIndex As String
    Dim Att(50) As Double
    Dim AttDescription(22) As String
    Dim strSQL As String

    Set db = CurrentDb()
    
    'Empty out the CURRENT table so it can be populated with the new data
    strSQL = "DELETE [tblPosition].[Job] "
    strSQL = strSQL + "FROM [tblPosition] "
    strSQL = strSQL + "WHERE ((([tblPosition].[Job]) Is Not Null)) "
    
    db.Execute (strSQL)
    strSQL = ""

    Set rs = db.OpenRecordset("FINAL SCORING_COMPOSITE", dbOpenDynaset)

    AttDescription(1) = "FeeType"
    AttDescription(2) = "CurER"
    AttDescription(3) = "PriorER"
    AttDescription(4) = "DelBill"
    AttDescription(5) = "AR90"
    AttDescription(6) = "War"
    AttDescription(7) = "RatioRev"
    AttDescription(8) = "ChngOrd"
    AttDescription(9) = "ChngCnt"
    AttDescription(11) = "PMTen"
    AttDescription(13) = "PMPer"
    AttDescription(14) = "CntPMs"
    AttDescription(15) = "CountTask"
    AttDescription(16) = "EmpTurn"
    AttDescription(17) = "MomDelMeth"
    AttDescription(18) = "ProgMgmt"
    AttDescription(19) = "ClientRt"
    AttDescription(20) = "DBChange"
    AttDescription(21) = "NewClient"
    AttDescription(22) = "Delivery_Method"


    Do Until rs.EOF
        jobNumber = rs![Job]
    
        Att(1) = Nz(DLookup("[FeeType]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(2) = Nz(DLookup("[CurER]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(3) = Nz(DLookup("[PriorER]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(4) = Nz(DLookup("[DelBill]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(5) = Nz(DLookup("[AR90]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(6) = Nz(DLookup("[War]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(7) = Nz(DLookup("[RatioRev]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(8) = Nz(DLookup("[ChngOrd]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(9) = Nz(DLookup("[ChngCnt]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(11) = Nz(DLookup("[PMTen]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(13) = Nz(DLookup("[PMPer]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(14) = Nz(DLookup("[CntPMs]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(15) = Nz(DLookup("[CountTask]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(16) = Nz(DLookup("[EmpTurn]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(17) = Nz(DLookup("[MomDelMeth]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(18) = Nz(DLookup("[ProgMgmt]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(19) = Nz(DLookup("[ClientRt]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(20) = Nz(DLookup("[DBChange]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(21) = Nz(DLookup("[NewClient]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        Att(22) = Nz(DLookup("[Delivery_Method]", "FINAL SCORING_COMPOSITE", "[Job] = '" & jobNumber & "'"))
        
        Dim pos1, pos2, pos3 As Integer
        Dim tempResult As String
        Dim i, j As Integer
        
        i = 1
        j = 2
        
        Do Until j = 23
            If Trim(Att(i)) >= Trim(Att(j)) Then  'Need to trim the values coming out of the Array
                pos1 = i
            Else
                pos1 = j
                i = j
            End If
            j = j + 1
        Loop
        
        i = 1
        j = 2
        
        'Do Until j = 23
        '    If i <> pos1 Then
        '        If j <> pos1 Then
        '            If Trim(Att(i)) >= Trim(Att(j)) Then
        '                pos2 = i
        '            Else
        '                pos2 = j
        '                i = j
        '            End If
        '            j = j + 1
        '        Else
        '            j = j + 1
        '        End If
        '    Else
        '        i = i + 1
        '        j = j + 1
        '    End If
        'Loop
        Att(pos1) = 0
        
        Do Until j = 23
            If Trim(Att(i)) >= Trim(Att(j)) Then  'Need to trim the values coming out of the Array
                pos2 = i
            Else
                pos2 = j
                i = j
            End If
            j = j + 1
        Loop
        
        Att(pos1) = 0
        Att(pos2) = 0
        
        i = 1
        j = 2
        
        Do Until j = 23
            If Trim(Att(i)) >= Trim(Att(j)) Then  'Need to trim the values coming out of the Array
                pos3 = i
            Else
                pos3 = j
                i = j
            End If
            j = j + 1
        Loop
                    
        
        '***********PUT RESULTS IN TABLE*************
        
        Dim attributeDesc1 As String
        Dim attributeDesc2 As String
        Dim attributeDesc3 As String
        Dim rst As Recordset
        
        attributeDesc1 = AttDescription(pos1)
        attributeDesc2 = AttDescription(pos2)
        attributeDesc3 = AttDescription(pos3)
        
        Set rst = db.OpenRecordset("tblPosition")
        rst.AddNew
        rst("Job").Value = jobNumber
        rst("Position_1").Value = attributeDesc1
        rst("Position_2").Value = attributeDesc2
        rst("Position_3").Value = attributeDesc3
        rst.Update
        
        rs.MoveNext
    Loop


    MsgBox ("Ordering script complete!")

End Sub
