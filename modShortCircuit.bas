Attribute VB_Name = "modShortCircuit"
'==========================
' Modul: modShortCircuit
' Posledn� �prava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

' V�po�et Ik3 (argumenty po�a s� Variant pre stabilitu)
Public Sub CalculateShortCircuit( _
    ByVal nBuses As Long, ByVal nBranches As Long, ByRef FromBus As Variant, ByRef ToBus As Variant, _
    ByRef R As Variant, ByRef X As Variant, ByRef BranchStatus As Variant, _
    ByVal nSwitches As Long, ByRef SwFrom As Variant, ByRef SwTo As Variant, ByRef SwR As Variant, ByRef SwX As Variant, ByRef SwStatus As Variant, _
    ByVal nTrafo As Long, ByRef TrFrom As Variant, ByRef TrTo As Variant, _
    ByRef TrR As Variant, ByRef TrX As Variant, ByRef TrRatio As Variant, _
    ByVal nReaktory As Long, ByRef ReaktorFrom As Variant, ByRef ReaktorTo As Variant, _
    ByRef ReaktorR As Variant, ByRef ReaktorX As Variant, _
    ByVal nDifReaktory As Long, ByRef DifReaktorFrom As Variant, ByRef DifReaktorTo As Variant, _
    ByRef DifReaktorR As Variant, ByRef DifReaktorX As Variant, _
    ByVal nMotors As Long, ByRef MotorBus As Variant, ByRef MotorXk As Variant, ByRef MotorStatus As Variant, _
    ByRef BusNames As Variant, ByRef BusTypes As Variant, ByRef BusBaseKV As Variant, _
    ByRef Ik_input As Variant, ByRef Ik_result As Variant, ByVal SBase_MVA As Double, _
    ByRef IsBusIsolated As Variant, ByRef IsBranchIsolated As Variant, ByRef IsTrafoIsolated As Variant, ByRef IsReaktorIsolated As Variant, ByRef IsDifReaktorIsolated As Variant, ByRef IsSwitchIsolated As Variant)

    Dim i As Long, J As Long, k As Long
    Dim Ysc() As Complex, MatBig() As Double, InvBig() As Double
    Dim Z As Complex, Ys As Complex, t1 As Complex, t2 As Complex, A As Double
    Dim slackIdx As Long, Ik_slack As Double, Z_grid_abs As Double, Un As Double
    Dim R_th As Double, X_th As Double, Z_th As Double
    
    ReDim Ysc(1 To nBuses, 1 To nBuses), Ik_result(1 To nBuses)
    
    ' Inicializ�cia Ysc (izolovan� = 1.0 na diagon�le)
    For i = 1 To nBuses
        For J = 1 To nBuses: Ysc(i, J) = CCreate(0, 0): Next J
        If IsBusIsolated(i) Then Ysc(i, i) = CCreate(1, 0)
    Next i
    
    ' Vedenia
    For k = 1 To nBranches
        If BranchStatus(k) > 0 Then
            If Not IsBranchIsolated(k) And Not (R(k) = 0 And X(k) = 0) Then
                Z = CCreate(CDbl(R(k)), CDbl(X(k))): Ys = CDiv(CCreate(1, 0), Z)
                i = FromBus(k): J = ToBus(k)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
                Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
            End If
        End If
    Next k
    
    ' Traf�
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            Z = CCreate(CDbl(TrR(k)), CDbl(TrX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = TrFrom(k): J = TrTo(k): A = TrRatio(k)
            t1 = CCreate(Ys.Re / (A * A), Ys.Im / (A * A))
            Ysc(i, i) = CAdd(Ysc(i, i), t1): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            t2 = CCreate(Ys.Re / A, Ys.Im / A)
            Ysc(i, J) = CSub(Ysc(i, J), t2): Ysc(J, i) = CSub(Ysc(J, i), t2)
        End If
    Next k
    
        ' Sp�na�e
    For k = 1 To nSwitches
        If SwStatus(k) > 0 Then
            If Not IsSwitchIsolated(k) And Not (SwR(k) = 0 And SwX(k) = 0) Then
                Z = CCreate(CDbl(SwR(k)), CDbl(SwX(k))): Ys = CDiv(CCreate(1, 0), Z)
                i = SwFrom(k): J = SwTo(k)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
                Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
            End If
        End If
    Next k

    ' Reaktory
    For k = 1 To nReaktory
        If Not IsReaktorIsolated(k) And Not (ReaktorR(k) = 0 And ReaktorX(k) = 0) Then
            Z = CCreate(CDbl(ReaktorR(k)), CDbl(ReaktorX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = ReaktorFrom(k): J = ReaktorTo(k)
            Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
        End If
    Next k
    
    ' Dif. Reaktory
    For k = 1 To nDifReaktory
        If Not IsDifReaktorIsolated(k) And Not (DifReaktorR(k) = 0 And DifReaktorX(k) = 0) Then
            Z = CCreate(CDbl(DifReaktorR(k)), CDbl(DifReaktorX(k))): Ys = CDiv(CCreate(1, 0), Z)
            i = DifReaktorFrom(k): J = DifReaktorTo(k)
            Ysc(i, i) = CAdd(Ysc(i, i), Ys): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            Ysc(i, J) = CSub(Ysc(i, J), Ys): Ysc(J, i) = CSub(Ysc(J, i), Ys)
        End If
    Next k
    
    ' Motory VN - pr�spevok do skratu
    ' Model: impedancia vo�i zemi Z = j*Xk (R sa zanedb�va alebo je v Xk zahrnut� ako impedancia)
    ' Y = 1 / (j*Xk) = -j / Xk
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            ' Kontrola na nenulov� reaktanciu
            If Abs(CDbl(MotorXk(k))) > 0.0000001 Then
                ' Ys = 1 / (j * Xk) = -j * (1/Xk)
                Ys = CCreate(0, -1# / CDbl(MotorXk(k)))
                i = MotorBus(k)
                ' Pridanie k diagon�le
                Ysc(i, i) = CAdd(Ysc(i, i), Ys)
            End If
        End If
    Next k
    
    ' Slack impedancia
    For i = 1 To nBuses
        If BusTypes(i) = 0 Then ' btSlack
            Ik_slack = Ik_input(i)
            If Ik_slack > 0 Then
                Un = BusBaseKV(i)
                Z_grid_abs = (1.1 * SBase_MVA) / (Sqr(3) * Un * Ik_slack)
                Ysc(i, i) = CAdd(Ysc(i, i), CDiv(CCreate(1, 0), CCreate(0, Z_grid_abs)))
            End If
            slackIdx = i: Exit For
        End If
    Next i
    
    ' Debug v�pis skratovej matice
    Call WriteSCMatrix(Ysc, BusNames)
    
    ' Inverzia (Double matica pre Excel MInverse)
    ReDim MatBig(1 To 2 * nBuses, 1 To 2 * nBuses)
    For i = 1 To nBuses
        For J = 1 To nBuses
            MatBig(i, J) = Ysc(i, J).Re: MatBig(nBuses + i, nBuses + J) = Ysc(i, J).Re
            MatBig(i, nBuses + J) = -Ysc(i, J).Im: MatBig(nBuses + i, J) = Ysc(i, J).Im
        Next J
    Next i
    
    On Error GoTo 0
    ' Pou�itie vlastnej funkcie na inverziu matice (namiesto excelovsk�ho MInverse)
    Call MatrixInverse_Gauss(MatBig, InvBig)
    
    ' V�po�et Ik
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Ik_result(i) = 0
        Else
            R_th = InvBig(i, i): X_th = InvBig(nBuses + i, i)
            Z_th = Sqr(R_th * R_th + X_th * X_th)
            Un = BusBaseKV(i)
            If Z_th > 0.0000001 Then
                Ik_result(i) = (1.1 / Z_th) * (SBase_MVA / (Sqr(3) * Un))
            Else
                Ik_result(i) = 0
            End If
        End If
    Next i
End Sub

Public Sub WriteShortCircuitResults(ByRef Ik_result As Variant, ByVal nBuses As Long)
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("uzly")
    ws.Cells(2, 10).Value = "Ik3'' [kA]"
    For i = 1 To nBuses
        ws.Cells(2 + i, 10).Value = Round(Ik_result(i), 2)
    Next i
End Sub

' Z�pis skratovej admitan�nej matice pre kontrolu
Private Sub WriteSCMatrix(ByRef Ysc() As Complex, ByRef BusNames As Variant)
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long, J As Long
    Dim row0 As Long, col0 As Long
    Dim txt As String
    
    Set ws = GetOrCreateSheet("SC_matica")
    ws.Cells.Clear
    
    n = UBound(Ysc, 1)

    row0 = 1
    col0 = 1

    Dim arr As Variant
    Dim startRowX As Long

    ' Blok Re(Ysc) - hlavi�ka + matica v jednom Variant poli, jeden Range.Value z�pis
    ws.Cells(row0, col0).Value = "Re(Ysc)"
    ReDim arr(1 To n + 1, 1 To n + 1)
    arr(1, 1) = ""
    For J = 1 To n
        arr(1, J + 1) = BusNames(J)
    Next J
    For i = 1 To n
        arr(i + 1, 1) = BusNames(i)
        For J = 1 To n
            arr(i + 1, J + 1) = Ysc(i, J).Re
        Next J
    Next i
    ws.Range(ws.Cells(row0 + 1, col0), ws.Cells(row0 + 1 + n, col0 + n)).Value = arr

    ' Blok Im(Ysc)
    startRowX = row0 + n + 3
    ws.Cells(startRowX, col0).Value = "Im(Ysc)"
    ReDim arr(1 To n + 1, 1 To n + 1)
    arr(1, 1) = ""
    For J = 1 To n
        arr(1, J + 1) = BusNames(J)
    Next J
    For i = 1 To n
        arr(i + 1, 1) = BusNames(i)
        For J = 1 To n
            arr(i + 1, J + 1) = Ysc(i, J).Im
        Next J
    Next i
    ws.Range(ws.Cells(startRowX + 1, col0), ws.Cells(startRowX + 1 + n, col0 + n)).Value = arr
End Sub



'--------------------------------------
' V�po�et inverznej matice pomocou Gaussovej elimin�cie
'--------------------------------------
Private Sub MatrixInverse_Gauss(ByRef A_in() As Double, ByRef A_inv() As Double)
    Dim n As Long, i As Long, J As Long, k As Long
    Dim maxRow As Long, maxValue As Double
    Dim temp As Double, factor As Double
    Dim A() As Double
    
    n = UBound(A_in, 1)
    ReDim A(1 To n, 1 To 2 * n)
    
    ' Pr�prava matice (A | I)
    For i = 1 To n
        For J = 1 To n
            A(i, J) = A_in(i, J)
            If i = J Then A(i, n + J) = 1# Else A(i, n + J) = 0#
        Next J
    Next i
    
    ' Gaussova elimin�cia
    For i = 1 To n
        maxRow = i
        maxValue = Abs(A(i, i))
        For k = i + 1 To n
            If Abs(A(k, i)) > maxValue Then
                maxValue = Abs(A(k, i))
                maxRow = k
            End If
        Next k
        
        If maxRow <> i Then
            For k = 1 To 2 * n
                temp = A(i, k): A(i, k) = A(maxRow, k): A(maxRow, k) = temp
            Next k
        End If
        
        If Abs(A(i, i)) < 1E-18 Then Err.Raise vbObjectError + 102, , "Matica je singul�rna."
        
        temp = A(i, i)
        For k = 1 To 2 * n: A(i, k) = A(i, k) / temp: Next k
        
        For k = 1 To n
            If k <> i Then
                factor = A(k, i)
                For J = 1 To 2 * n
                    A(k, J) = A(k, J) - factor * A(i, J)
                Next J
            End If
        Next k
    Next i
    
    ' Extrakcia inverznej matice
    ReDim A_inv(1 To n, 1 To n)
    For i = 1 To n: For J = 1 To n: A_inv(i, J) = A(i, n + J): Next J: Next i
End Sub


