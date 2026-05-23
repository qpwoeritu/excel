Option Explicit

Public Sub BuildYBus(ByVal nBuses As Long, ByVal nBranches As Long, ByRef FromBus() As Long, ByRef ToBus() As Long, _
                     ByRef R() As Double, ByRef X() As Double, ByRef BranchStatus() As Integer, ByRef Bshunt() As Double, _
                     ByVal nSwitches As Long, ByRef SwFrom() As Long, ByRef SwTo() As Long, ByRef SwR() As Double, ByRef SwX() As Double, ByRef SwStatus() As Integer, _
                     ByVal nTrafo As Long, ByRef TrFrom() As Long, ByRef TrTo() As Long, _
                     ByRef TrR() As Double, ByRef TrX() As Double, ByRef TrG() As Double, ByRef TrB() As Double, ByRef TrRatio() As Double, _
                     ByVal nReaktory As Long, ByRef ReaktorFrom() As Long, ByRef ReaktorTo() As Long, ByRef ReaktorR() As Double, ByRef ReaktorX() As Double, _
                     ByVal nDifReaktory As Long, ByRef DifReaktorFrom() As Long, ByRef DifReaktorTo() As Long, ByRef DifReaktorR() As Double, ByRef DifReaktorX() As Double, _
                     ByVal nComp As Long, ByRef CompBus() As Long, ByRef CompB() As Double, ByRef CompStatus() As Integer, _
                     ByVal nMotors As Long, ByRef MotorBus() As Long, ByRef MotorG() As Double, ByRef MotorB() As Double, ByRef MotorStatus() As Integer, _
                     ByVal nGenBr As Long, ByRef GenBrFrom() As Long, ByRef GenBrTo() As Long, ByRef GenBrR() As Double, ByRef GenBrX() As Double, _
                     ByRef BusNames() As String, ByRef IsBusIsolated() As Boolean, ByRef IsBranchIsolated() As Boolean, ByRef IsTrafoIsolated() As Boolean, ByRef IsReaktorIsolated() As Boolean, ByRef IsDifReaktorIsolated() As Boolean, ByRef IsSwitchIsolated() As Boolean, _
                     ByRef Y() As Complex, ByRef G() As Double, ByRef B() As Double)
    
    Dim i As Long, J As Long, k As Long
    Dim Z As Complex, Ys As Complex, Ym As Complex, A As Double, t1 As Complex, t2 As Complex
    
    ReDim Y(1 To nBuses, 1 To nBuses), G(1 To nBuses, 1 To nBuses), B(1 To nBuses, 1 To nBuses)
    For i = 1 To nBuses: For J = 1 To nBuses: Y(i, J) = CCreate(0, 0): Next J: Next i
    
    ' Vedenia
    For k = 1 To nBranches
        ' Kontrola Statusu
        If BranchStatus(k) > 0 Then
            If Not IsBranchIsolated(k) And Not (R(k) = 0 And X(k) = 0) Then
                Z = CCreate(R(k), X(k)): Ys = CDiv(CCreate(1, 0), Z)
                i = FromBus(k): J = ToBus(k)
                Y(i, i) = CAdd(Y(i, i), Ys): Y(J, J) = CAdd(Y(J, J), Ys)
                Y(i, J) = CSub(Y(i, J), Ys): Y(J, i) = CSub(Y(J, i), Ys)
                
                ' Pridanie priečnej kapacity (PI článok: jB/2 na oboch koncoch)
                If Bshunt(k) <> 0# Then
                    Dim Ysh As Complex
                    Ysh = CCreate(0, Bshunt(k) / 2#)
                    Y(i, i) = CAdd(Y(i, i), Ysh)
                    Y(J, J) = CAdd(Y(J, J), Ysh)
                End If
            End If
        End If
    Next k
    
    ' Trafá
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            Z = CCreate(TrR(k), TrX(k)): Ys = CDiv(CCreate(1, 0), Z): Ym = CCreate(TrG(k), TrB(k))
            i = TrFrom(k): J = TrTo(k): A = TrRatio(k)
            t1 = CCreate(Ys.Re / (A * A), Ys.Im / (A * A)) ' ys/a^2
            Y(i, i) = CAdd(Y(i, i), CAdd(t1, Ym))
            Y(J, J) = CAdd(Y(J, J), Ys)
            t2 = CCreate(Ys.Re / A, Ys.Im / A) ' ys/a
            Y(i, J) = CSub(Y(i, J), t2): Y(J, i) = CSub(Y(J, i), t2)
        End If
    Next k
    
    ' Reaktory
    For k = 1 To nReaktory
        If Not IsReaktorIsolated(k) And Not (ReaktorR(k) = 0 And ReaktorX(k) = 0) Then
            Z = CCreate(ReaktorR(k), ReaktorX(k))
            Ys = CDiv(CCreate(1, 0), Z)
            i = ReaktorFrom(k): J = ReaktorTo(k)
            Y(i, i) = CAdd(Y(i, i), Ys)
            Y(J, J) = CAdd(Y(J, J), Ys)
            Y(i, J) = CSub(Y(i, J), Ys)
            Y(J, i) = CSub(Y(J, i), Ys)
        End If
    Next k
    
    ' Dif. Reaktory
    For k = 1 To nDifReaktory
        If Not IsDifReaktorIsolated(k) And Not (DifReaktorR(k) = 0 And DifReaktorX(k) = 0) Then
            Z = CCreate(DifReaktorR(k), DifReaktorX(k))
            Ys = CDiv(CCreate(1, 0), Z)
            i = DifReaktorFrom(k): J = DifReaktorTo(k)
            Y(i, i) = CAdd(Y(i, i), Ys)
            Y(J, J) = CAdd(Y(J, J), Ys)
            Y(i, J) = CSub(Y(i, J), Ys)
            Y(J, i) = CSub(Y(J, i), Ys)
        End If
    Next k
    
        ' Spínače
    For k = 1 To nSwitches
        If SwStatus(k) > 0 Then
            If Not IsSwitchIsolated(k) And Not (SwR(k) = 0 And SwX(k) = 0) Then
                Z = CCreate(SwR(k), SwX(k)): Ys = CDiv(CCreate(1, 0), Z)
                i = SwFrom(k): J = SwTo(k)
                Y(i, i) = CAdd(Y(i, i), Ys): Y(J, J) = CAdd(Y(J, J), Ys)
                Y(i, J) = CSub(Y(i, J), Ys): Y(J, i) = CSub(Y(J, i), Ys)
            End If
        End If
    Next k

    ' Kompenzácia
    ' Pridanie susceptancie (j*B) k diagonále
    For k = 1 To nComp
        ' Ak je kompenzácia zapnutá (Status=1)
        If CompStatus(k) = 1 Then
            i = CompBus(k)
            ' Y_ii = Y_ii + j*CompB
            Y(i, i) = CAdd(Y(i, i), CCreate(0, CompB(k)))
        End If
    Next k
    
    ' Motory VN (pre Load Flow)
    ' Pridanie admitancie (G + jB) k diagonále
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            i = MotorBus(k)
            ' Y_ii = Y_ii + (G + jB)
            Y(i, i) = CAdd(Y(i, i), CCreate(MotorG(k), MotorB(k)))
        End If
    Next k

    ' Generátorové vetvy (EMF model: fantóm -- Ra+jXs --> svorka)
    For k = 1 To nGenBr
        If Not (GenBrR(k) = 0 And GenBrX(k) = 0) Then
            Z = CCreate(GenBrR(k), GenBrX(k)): Ys = CDiv(CCreate(1, 0), Z)
            i = GenBrFrom(k): J = GenBrTo(k)
            Y(i, i) = CAdd(Y(i, i), Ys): Y(J, J) = CAdd(Y(J, J), Ys)
            Y(i, J) = CSub(Y(i, J), Ys): Y(J, i) = CSub(Y(J, i), Ys)
        End If
    Next k

    ' Izolované uzly (aby matica nebola singulárna)
    For i = 1 To nBuses
        If IsBusIsolated(i) Then Y(i, i) = CCreate(1, 0)
        For J = 1 To nBuses: G(i, J) = Y(i, J).Re: B(i, J) = Y(i, J).Im: Next J
    Next i
    
    Call WriteYMatrix(Y, G, B, BusNames)
End Sub




