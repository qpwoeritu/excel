Attribute VB_Name = "modTopology"
'==========================
' Modul: modTopology
' Posledná úprava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

Public Sub FindIsolatedParts(ByVal nBuses As Long, ByVal nBranches As Long, ByRef FromBus() As Long, ByRef ToBus() As Long, ByRef BranchStatus() As Integer, _
                             ByVal nTrafo As Long, ByRef TrFrom() As Long, ByRef TrTo() As Long, _
                             ByVal nReaktory As Long, ByRef ReaktorFrom() As Long, ByRef ReaktorTo() As Long, _
                             ByVal nDifReaktory As Long, ByRef DifReaktorFrom() As Long, ByRef DifReaktorTo() As Long, _
                             ByVal nSwitches As Long, ByRef SwFrom() As Long, ByRef SwTo() As Long, ByRef SwStatus() As Integer, _
                             ByVal nComp As Long, ByRef CompBus() As Long, _
                             ByVal nMotors As Long, ByRef MotorBus() As Long, _
                             ByRef BusTypes() As BusType, ByRef IsBusIsolated() As Boolean, _
                             ByRef IsBranchIsolated() As Boolean, ByRef IsTrafoIsolated() As Boolean, ByRef IsReaktorIsolated() As Boolean, ByRef IsDifReaktorIsolated() As Boolean, ByRef IsSwitchIsolated() As Boolean, ByRef IsCompIsolated() As Boolean, ByRef IsMotorIsolated() As Boolean, _
                             ByRef isolatedCount As Long)
    Dim i As Long, u As Long, v As Long, slackIdx As Long
    Dim adjHead() As Long, adjNext() As Long, adjTo() As Long, adjIdx() As Long
    Dim adjType() As Integer ' 1=Branch, 2=Trafo, 3=Reaktor, 4=DifReaktor, 5=Switch
    Dim edgeCount As Long, totalEdges As Long, queue() As Long, qH As Long, qT As Long, visited() As Boolean
    
    If nBuses = 0 Then Exit Sub
    ReDim IsBusIsolated(1 To nBuses), visited(1 To nBuses)
    For i = 1 To nBuses: IsBusIsolated(i) = True: Next i
    If nBranches > 0 Then ReDim IsBranchIsolated(1 To nBranches): For i = 1 To nBranches: IsBranchIsolated(i) = True: Next i
    If nTrafo > 0 Then ReDim IsTrafoIsolated(1 To nTrafo): For i = 1 To nTrafo: IsTrafoIsolated(i) = True: Next i
    If nReaktory > 0 Then ReDim IsReaktorIsolated(1 To nReaktory): For i = 1 To nReaktory: IsReaktorIsolated(i) = True: Next i
    If nDifReaktory > 0 Then ReDim IsDifReaktorIsolated(1 To nDifReaktory): For i = 1 To nDifReaktory: IsDifReaktorIsolated(i) = True: Next i
    If nSwitches > 0 Then ReDim IsSwitchIsolated(1 To nSwitches): For i = 1 To nSwitches: IsSwitchIsolated(i) = True: Next i
    If nComp > 0 Then ReDim IsCompIsolated(1 To nComp): For i = 1 To nComp: IsCompIsolated(i) = True: Next i
    If nMotors > 0 Then ReDim IsMotorIsolated(1 To nMotors): For i = 1 To nMotors: IsMotorIsolated(i) = True: Next i
    
    For i = 1 To nBuses
        If BusTypes(i) = btSlack Then slackIdx = i: Exit For
    Next i
    If slackIdx = 0 Then Exit Sub
    
    totalEdges = 2 * (nBranches + nTrafo + nReaktory + nDifReaktory + nSwitches) + 1
    ReDim adjHead(1 To nBuses), adjNext(1 To totalEdges), adjTo(1 To totalEdges), adjIdx(1 To totalEdges), adjType(1 To totalEdges)
    
    edgeCount = 0
    ' Pridanie hrán (vedenia) - iba zapnuté
    For i = 1 To nBranches
        If BranchStatus(i) > 0 Then
            u = FromBus(i): v = ToBus(i)
            edgeCount = edgeCount + 1: adjTo(edgeCount) = v: adjIdx(edgeCount) = i: adjType(edgeCount) = 1: adjNext(edgeCount) = adjHead(u): adjHead(u) = edgeCount
            edgeCount = edgeCount + 1: adjTo(edgeCount) = u: adjIdx(edgeCount) = i: adjType(edgeCount) = 1: adjNext(edgeCount) = adjHead(v): adjHead(v) = edgeCount
        End If
    Next i
    ' Pridanie hrán (trafá)
    For i = 1 To nTrafo
        u = TrFrom(i): v = TrTo(i)
        edgeCount = edgeCount + 1: adjTo(edgeCount) = v: adjIdx(edgeCount) = i: adjType(edgeCount) = 2: adjNext(edgeCount) = adjHead(u): adjHead(u) = edgeCount
        edgeCount = edgeCount + 1: adjTo(edgeCount) = u: adjIdx(edgeCount) = i: adjType(edgeCount) = 2: adjNext(edgeCount) = adjHead(v): adjHead(v) = edgeCount
    Next i
    ' Pridanie hrán (reaktory)
    For i = 1 To nReaktory
        u = ReaktorFrom(i): v = ReaktorTo(i)
        edgeCount = edgeCount + 1: adjTo(edgeCount) = v: adjIdx(edgeCount) = i: adjType(edgeCount) = 3: adjNext(edgeCount) = adjHead(u): adjHead(u) = edgeCount
        edgeCount = edgeCount + 1: adjTo(edgeCount) = u: adjIdx(edgeCount) = i: adjType(edgeCount) = 3: adjNext(edgeCount) = adjHead(v): adjHead(v) = edgeCount
    Next i
    ' Pridanie hrán (dif. reaktory)
    For i = 1 To nDifReaktory
        u = DifReaktorFrom(i): v = DifReaktorTo(i)
        edgeCount = edgeCount + 1: adjTo(edgeCount) = v: adjIdx(edgeCount) = i: adjType(edgeCount) = 4: adjNext(edgeCount) = adjHead(u): adjHead(u) = edgeCount
        edgeCount = edgeCount + 1: adjTo(edgeCount) = u: adjIdx(edgeCount) = i: adjType(edgeCount) = 4: adjNext(edgeCount) = adjHead(v): adjHead(v) = edgeCount
    Next i
    ' Pridanie hrán (spínaèe) - iba zapnuté
    For i = 1 To nSwitches
        If SwStatus(i) > 0 Then
            u = SwFrom(i): v = SwTo(i)
            edgeCount = edgeCount + 1: adjTo(edgeCount) = v: adjIdx(edgeCount) = i: adjType(edgeCount) = 5: adjNext(edgeCount) = adjHead(u): adjHead(u) = edgeCount
            edgeCount = edgeCount + 1: adjTo(edgeCount) = u: adjIdx(edgeCount) = i: adjType(edgeCount) = 5: adjNext(edgeCount) = adjHead(v): adjHead(v) = edgeCount
        End If
    Next i
    
    ' BFS
    ReDim queue(1 To nBuses)
    qH = 1: qT = 1: queue(1) = slackIdx: visited(slackIdx) = True: IsBusIsolated(slackIdx) = False
    
    Do While qH <= qT
        u = queue(qH): qH = qH + 1
        i = adjHead(u)
        Do While i > 0
            v = adjTo(i)
            
            Select Case adjType(i)
                Case 1: IsBranchIsolated(adjIdx(i)) = False
                Case 2: IsTrafoIsolated(adjIdx(i)) = False
                Case 3: IsReaktorIsolated(adjIdx(i)) = False
                Case 4: IsDifReaktorIsolated(adjIdx(i)) = False
                Case 5: IsSwitchIsolated(adjIdx(i)) = False
            End Select
            
            If Not visited(v) Then
                visited(v) = True: IsBusIsolated(v) = False: qT = qT + 1: queue(qT) = v
            End If
            i = adjNext(i)
        Loop
    Loop
    
    ' Urèenie izolovanosti kompenzácie (pod¾a uzla)
    For i = 1 To nComp
        IsCompIsolated(i) = IsBusIsolated(CompBus(i))
    Next i
    
    ' Urèenie izolovanosti motorov (pod¾a uzla)
    For i = 1 To nMotors
        IsMotorIsolated(i) = IsBusIsolated(MotorBus(i))
    Next i
    
    isolatedCount = 0
    For i = 1 To nBuses: If IsBusIsolated(i) Then isolatedCount = isolatedCount + 1
    Next i
End Sub




