Attribute VB_Name = "funcAttributeFunctions"
Option Compare Database
'Script Name:       funcFunctions
'Author:            Steve O'Neal
'Created:           2/5/2014
'Last Modified:     3/12/2015
'Version:           1.1
'Dependency:        tblScoring table
'
'Update:            1.1 Changed table names
'
'Script is a collection of functions that are used to determine score of each
'attributed used to evaluate jobs.  Each function can then be used in a query
'of those attributes.


Function FeeType(inputText As String) As String

    Dim result As String
    
    If inputText = "00" Then
        result = "0"
    ElseIf inputText = "N/A" Then
        result = "1"
    Else
        result = Nz(DLookup("[Score]", "tblDeliveryMethodScore", "[Value] = '" & inputText & "'"))
    End If
    
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 1
    
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    FeeType = Round(result * reDistWt, 3)

End Function


Function C_ER(inputVal As Double) As Double
'Round(IIf([Ratio_YTDtoCon_ER]<=0,0,IIf([Ratio_YTDtoCon_ER]>0.3,10,(([Ratio_YTDtoCon_ER]-0)/(0.3-0))*10)),3) AS TEST

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 2
    
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        C_ER = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        C_ER = scaleMax * reDistWt
    Else
        C_ER = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If

End Function

Function P_ER(inputVal As Double) As Double

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 3
    
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        P_ER = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        P_ER = scaleMax * reDistWt
    Else
        P_ER = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If

End Function

Function DelB(inputVal As Double) As Double
'=IF(E23<=tblScoring!$D$7,tblScoring!$F$7,IF(E23>=tblScoring!$E$7,tblScoring!$G$7,(E23-tblScoring!$D$7)/(tblScoring!$E$7-tblScoring!$D$7)*tblScoring!$G$7))

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 4
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        DelB = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        DelB = scaleMax * reDistWt
    Else
        DelB = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function


Function AR90(inputVal As Double) As Double
'=IF(F23<=tblScoring!$D$8,tblScoring!$F$8,IF(F23>=tblScoring!$E$8,tblScoring!$G$8,(F23-tblScoring!$D$8)/(tblScoring!$E$8-tblScoring!$D$8)*tblScoring!$G$8))

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 5
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        AR90 = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        AR90 = scaleMax * reDistWt
    Else
        AR90 = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function WAR(inputVal As Double) As Double
'=IF(G23<=tblScoring!$D$9,tblScoring!$F$9,IF(G23>=tblScoring!$E$9,tblScoring!$G$9,(G23-tblScoring!$D$9)/(tblScoring!$E$9-tblScoring!$D$9)*tblScoring!$G$9))
    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 6
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        WAR = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        WAR = scaleMax * reDistWt
    Else
        WAR = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
End Function

Function RatioRev(inputVal As Double) As Double
'=IF(H23<=tblScoring!$D$10,tblScoring!$F$10,IF(H23>=tblScoring!$E$10,tblScoring!$G$10,(H23-tblScoring!$E$10)/(tblScoring!$D$10-tblScoring!$E$10)*tblScoring!$F$10))

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 7
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal = 0 Then
        RatioRev = 0
    ElseIf inputVal <= minRange Then
        RatioRev = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        RatioRev = scaleMax * reDistWt
    Else
        RatioRev = Round((((inputVal - maxRange) / (minRange - maxRange)) * scaleMin) * reDistWt, 3)
    End If
    
End Function

Function ChgOrd(inputVal As Double) As Double
'=IF(I14<=tblScoring!$D$11,tblScoring!$F$11,IF(I14>=tblScoring!$E$11,tblScoring!$G$11,(I14-tblScoring!$D$11)/(tblScoring!$E$11-tblScoring!$D$11)*tblScoring!$G$11))

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 8
    
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        ChgOrd = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        ChgOrd = scaleMax * reDistWt
    Else
        ChgOrd = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
        
End Function

Function ChgCnt(inputVal As Double) As Double
'
    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 9
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        ChgCnt = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        ChgCnt = scaleMax * reDistWt
    Else
        ChgCnt = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function PMTen(inputVal As Double) As Double
'

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 11
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        PMTen = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        PMTen = scaleMax * reDistWt
    Else
        PMTen = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function PMPer(inputVal As Double) As Double
'

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    
    Dim test As Double
    
    catIndex = 13
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    test = Val(Format(maxRange, "Standard"))
    
    If inputVal <= Val(Format(minRange, "Standard")) Then
        PMPer = scaleMin * reDistWt
    ElseIf inputVal >= Val(Format(maxRange, "Standard")) Then
        PMPer = scaleMax * reDistWt
    Else
        PMPer = Round((((inputVal - Val(Format(maxRange, "Standard"))) / (Val(Format(minRange, "Standard")) - Val(Format(maxRange, "Standard")))) * scaleMin) * reDistWt, 3)
    End If
    
End Function


Function CountPM(inputVal As Double) As Double
'=IF(N2<=tblScoring!$D$17,tblScoring!$F$17,IF(N2>=tblScoring!$E$17,tblScoring!$G$17,(N2-tblScoring!$D$17)/(tblScoring!$E$17-tblScoring!$D$17)*tblScoring!$G$17))
    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 14
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        CountPM = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        CountPM = scaleMax * reDistWt
    Else
        CountPM = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function CountTask(inputVal As Double) As Double
'=IF(O3<=tblScoring!$D$18,tblScoring!$F$18,IF(O3>=tblScoring!$E$18,tblScoring!$G$18,(O3-tblScoring!$D$18)/(tblScoring!$E$18-tblScoring!$D$18)*tblScoring!$G$18))
    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 15
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        CountTask = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        CountTask = scaleMax * reDistWt
    Else
        CountTask = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function EmpTurn(inputVal As Double) As Double
'=IF(P2<=tblScoring!$D$19,tblScoring!$F$19,IF(P2>=tblScoring!$E$19,tblScoring!$G$19,(P2-tblScoring!$D$19)/(tblScoring!$E$19-tblScoring!$D$19)*tblScoring!$G$19))
    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    Dim test As Double
    
    catIndex = 16
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    test = Val(Format(maxRange, "Standard"))
    
    If inputVal <= Val(Format(minRange, "Standard")) Then
        EmpTurn = scaleMin * reDistWt
    ElseIf inputVal >= Val(Format(maxRange, "Standard")) Then
        EmpTurn = scaleMax * reDistWt
    Else
        EmpTurn = Round((((inputVal - Val(Format(minRange, "Standard"))) / (Val(Format(maxRange, "Standard")) - Val(Format(minRange, "Standard")))) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function MomDelMeth(inputText As String) As String
    Dim result As String
    
    result = Nz(DLookup("[Score]", "tblDeliveryMethodScore", "[Value] = '" & inputText & "'"))
    
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 17
    
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    MomDelMeth = result * reDistWt
    
End Function

Function ProgMgmt(inputText As String) As String
    Dim result As String
    
    result = Nz(DLookup("[Score]", "tblDeliveryMethodScore", "[Value] = '" & inputText & "'"))
    
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 18
    
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    ProgMgmt = result * reDistWt

End Function


Function ClientRt(inputVal As Double) As Double
'=IF(S2<=tblScoring!$D$22,tblScoring!$F$22,IF(S2>=tblScoring!$E$22,tblScoring!$G$22,(S2-tblScoring!$D$22)/(tblScoring!$E$22-tblScoring!$D$22)*tblScoring!$G$22))

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 19
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        ClientRt = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        ClientRt = scaleMax * reDistWt
    Else
        ClientRt = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function DBChangeOrd(inputVal As Double) As Double

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 20
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        DBChangeOrd = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        DBChangeOrd = scaleMax * reDistWt
    Else
        DBChangeOrd = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function

Function NewClient(inputVal As Double) As Double

    Dim minRange As String
    Dim maxRange As String
    Dim scaleMin As String
    Dim scaleMax As String
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 21
    minRange = Nz(DLookup("[Range Minimum]", "tblScoring", "[Index] = " & catIndex))
    maxRange = Nz(DLookup("[Range Maximum]", "tblScoring", "[Index] = " & catIndex))
    scaleMin = Nz(DLookup("[Scale Minimum]", "tblScoring", "[Index] = " & catIndex))
    scaleMax = Nz(DLookup("[Scale Maximum]", "tblScoring", "[Index] = " & catIndex))
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    If inputVal <= minRange Then
        NewClient = scaleMin * reDistWt
    ElseIf inputVal >= maxRange Then
        NewClient = scaleMax * reDistWt
    Else
        NewClient = Round((((inputVal - minRange) / (maxRange - minRange)) * scaleMax) * reDistWt, 3)
    End If
    
End Function
Function DelMeth(inputText As String) As String
    Dim result As String
    
    result = Nz(DLookup("[Score]", "tblDeliveryMethodScore", "[Value] = '" & inputText & "'"))
    
    Dim reDistWt As String
    Dim catIndex As Integer
    
    catIndex = 22
    
    reDistWt = Nz(DLookup("[Redistributed Weight]", "tblScoring", "[Index] = " & catIndex))
    
    DelMeth = result * reDistWt
    
End Function
