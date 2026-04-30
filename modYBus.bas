Attribute VB_Name = "modYBus"
'==========================
' Modul: modYBus
' Posledná úprava: 29.04.2026 00:00 (Bratislava)
'==========================
Option Explicit

Private Const EPS_ZERO As Double = 0.000000000001

Private Sub AddSeriesElement(ByRef G() As Double, ByRef B() As Double, ByVal i As Long, ByVal j As Long, ByVal gSeries As Double, ByVal bSeries As Double)
    G(i, i) = G(i, i) + gSeries
    B(i, i) = B(i, i) + bSeries

    G(j, j) = G(j, j) + gSeries
    B(j, j) = B(j, j) + bSeries

    G(i, j) = G(i, j) - gSeries
    B(i, j) = B(i, j) - bSeries

    G(j, i) = G(j, i) - gSeries
    B(j, i) = B(j, i) - bSeries
End Sub

Private Sub AddShuntToDiagonal(ByRef G() As Double, ByRef B() As Double, ByVal i As Long, ByVal gShunt As Double, ByVal bShunt As Double)
    G(i, i) = G(i, i) + gShunt
    B(i, i) = B(i, i) + bShunt
End Sub

Private Function TrySeriesAdmittance(ByVal r As Double, ByVal x As Double, ByRef gSeries As Double, ByRef bSeries As Double) As Boolean
    Dim den As Double

    den = r * r + x * x
    If den < EPS_ZERO Then
        TrySeriesAdmittance = False
        Exit Function
    End If

    ' 1 / (r + jx) = (r - jx) / (r^2 + x^2)
    gSeries = r / den
    bSeries = -x / den
    TrySeriesAdmittance = True
End Function

Public Sub BuildYBus(ByVal nBuses As Long, ByVal nBranches As Long, ByRef FromBus() As Long, ByRef ToBus() As Long, _
                     ByRef R() As Double, ByRef X() As Double, ByRef BranchStatus() As Integer, ByRef Bshunt() As Double, _
                     ByVal nSwitches As Long, ByRef SwFrom() As Long, ByRef SwTo() As Long, ByRef SwR() As Double, ByRef SwX() As Double, ByRef SwStatus() As Integer, _
                     ByVal nTrafo As Long, ByRef TrFrom() As Long, ByRef TrTo() As Long, _
                     ByRef TrR() As Double, ByRef TrX() As Double, ByRef TrG() As Double, ByRef TrB() As Double, ByRef TrRatio() As Double, _
                     ByVal nReaktory As Long, ByRef ReaktorFrom() As Long, ByRef ReaktorTo() As Long, ByRef ReaktorR() As Double, ByRef ReaktorX() As Double, _
                     ByVal nDifReaktory As Long, ByRef DifReaktorFrom() As Long, ByRef DifReaktorTo() As Long, ByRef DifReaktorR() As Double, ByRef DifReaktorX() As Double, _
                     ByVal nComp As Long, ByRef CompBus() As Long, ByRef CompB() As Double, ByRef CompStatus() As Integer, _
                     ByVal nMotors As Long, ByRef MotorBus() As Long, ByRef MotorG() As Double, ByRef MotorB() As Double, ByRef MotorStatus() As Integer, _
                     ByRef BusNames() As String, ByRef IsBusIsolated() As Boolean, ByRef IsBranchIsolated() As Boolean, ByRef IsTrafoIsolated() As Boolean, ByRef IsReaktorIsolated() As Boolean, ByRef IsDifReaktorIsolated() As Boolean, ByRef IsSwitchIsolated() As Boolean, _
                     ByRef Y() As Complex, ByRef G() As Double, ByRef B() As Double)

    Dim i As Long, j As Long, k As Long
    Dim gi As Double, bi As Double
    Dim ratio As Double, invA As Double, invA2 As Double

    ReDim G(1 To nBuses, 1 To nBuses)
    ReDim B(1 To nBuses, 1 To nBuses)

    ' Vedenia
    For k = 1 To nBranches
        If BranchStatus(k) > 0 Then
            If Not IsBranchIsolated(k) Then
                If TrySeriesAdmittance(R(k), X(k), gi, bi) Then
                    i = FromBus(k): j = ToBus(k)
                    Call AddSeriesElement(G, B, i, j, gi, bi)

                    If Bshunt(k) <> 0# Then
                        Call AddShuntToDiagonal(G, B, i, 0#, Bshunt(k) / 2#)
                        Call AddShuntToDiagonal(G, B, j, 0#, Bshunt(k) / 2#)
                    End If
                End If
            End If
        End If
    Next k

    ' Trafá
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) Then
            If TrySeriesAdmittance(TrR(k), TrX(k), gi, bi) Then
                i = TrFrom(k): j = TrTo(k)
                ratio = TrRatio(k)
                If Abs(ratio) < EPS_ZERO Then ratio = 1#
                invA = 1# / ratio
                invA2 = invA * invA

                ' Yii += ys/a^2 + ym
                G(i, i) = G(i, i) + gi * invA2 + TrG(k)
                B(i, i) = B(i, i) + bi * invA2 + TrB(k)

                ' Yjj += ys
                G(j, j) = G(j, j) + gi
                B(j, j) = B(j, j) + bi

                ' Yij,Yji -= ys/a
                G(i, j) = G(i, j) - gi * invA
                B(i, j) = B(i, j) - bi * invA
                G(j, i) = G(j, i) - gi * invA
                B(j, i) = B(j, i) - bi * invA
            End If
        End If
    Next k

    ' Reaktory
    For k = 1 To nReaktory
        If Not IsReaktorIsolated(k) Then
            If TrySeriesAdmittance(ReaktorR(k), ReaktorX(k), gi, bi) Then
                i = ReaktorFrom(k): j = ReaktorTo(k)
                Call AddSeriesElement(G, B, i, j, gi, bi)
            End If
        End If
    Next k

    ' Dif. reaktory
    For k = 1 To nDifReaktory
        If Not IsDifReaktorIsolated(k) Then
            If TrySeriesAdmittance(DifReaktorR(k), DifReaktorX(k), gi, bi) Then
                i = DifReaktorFrom(k): j = DifReaktorTo(k)
                Call AddSeriesElement(G, B, i, j, gi, bi)
            End If
        End If
    Next k

    ' Spínače
    For k = 1 To nSwitches
        If SwStatus(k) > 0 Then
            If Not IsSwitchIsolated(k) Then
                If TrySeriesAdmittance(SwR(k), SwX(k), gi, bi) Then
                    i = SwFrom(k): j = SwTo(k)
                    Call AddSeriesElement(G, B, i, j, gi, bi)
                End If
            End If
        End If
    Next k

    ' Kompenzácia
    For k = 1 To nComp
        If CompStatus(k) = 1 Then
            i = CompBus(k)
            Call AddShuntToDiagonal(G, B, i, 0#, CompB(k))
        End If
    Next k

    ' Motory VN
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            i = MotorBus(k)
            Call AddShuntToDiagonal(G, B, i, MotorG(k), MotorB(k))
        End If
    Next k

    ' Izolované uzly + prevod do komplexnej matice pre existujúci export
    ReDim Y(1 To nBuses, 1 To nBuses)
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            G(i, i) = 1#
            B(i, i) = 0#
        End If

        For j = 1 To nBuses
            Y(i, j) = CCreate(G(i, j), B(i, j))
        Next j
    Next i

    Call WriteYMatrix(Y, G, B, BusNames)
End Sub
