Option Explicit

'==========================
' Modul: modNR
' Newton-Raphsonov load-flow. Verejný vstupný bod je RunNRPhase – volá ho runCALC
' v rámci fázy "Výpočet load-flow" (index!I5/J5/I6).
' Žiadne StatusBar progressy – časovač a iteráciu zapisuje runCALC do buniek
' I6/J5 priamo, heartbeat z dlhej slučky cez PhaseYield z modProgress.
'==========================

'--------------------------------------
' Výpočet činných a jalových výkonov P, Q
' Optimalizácia A4: predpočítané Cos/Sin pre každý uhol uzla;
' v dvojnásobnom cykle použijeme súčtové vzorce namiesto opätovných
' volaní Cos/Sin (znižuje počet transcendentných volaní z 2·n² na 2·n).
'--------------------------------------
Private Sub CalcPower(ByVal nBuses As Long, _
                      ByRef G() As Double, _
                      ByRef B() As Double, _
                      ByRef Vmag() As Double, _
                      ByRef Vang() As Double, _
                      ByRef Pcalc() As Double, _
                      ByRef Qcalc() As Double)
    Dim i As Long, k As Long
    Dim costh As Double, sinth As Double, ViVk As Double

    Dim cosV() As Double, sinV() As Double
    ReDim cosV(1 To nBuses)
    ReDim sinV(1 To nBuses)
    For i = 1 To nBuses
        cosV(i) = Cos(Vang(i))
        sinV(i) = Sin(Vang(i))
    Next i

    Dim Gik As Double, Bik As Double
    For i = 1 To nBuses
        Pcalc(i) = 0#
        Qcalc(i) = 0#
        For k = 1 To nBuses
            Gik = G(i, k)
            Bik = B(i, k)
            If Gik <> 0# Or Bik <> 0# Then
                costh = cosV(i) * cosV(k) + sinV(i) * sinV(k)
                sinth = sinV(i) * cosV(k) - cosV(i) * sinV(k)
                ViVk = Vmag(i) * Vmag(k)
                Pcalc(i) = Pcalc(i) + ViVk * (Gik * costh + Bik * sinth)
                Qcalc(i) = Qcalc(i) + ViVk * (Gik * sinth - Bik * costh)
            End If
        Next k
    Next i
End Sub

'--------------------------------------
' Zostavenie vektora nesúladu ?P, ?Q (alebo ?V pre PV)
'--------------------------------------
Private Sub BuildMismatchVectors(ByVal nBuses As Long, _
                                 ByRef BusTypes() As BusType, _
                                 ByRef Pspec() As Double, _
                                 ByRef Qspec() As Double, _
                                 ByRef Vmag() As Double, _
                                 ByRef Pcalc() As Double, _
                                 ByRef Qcalc() As Double, _
                                 ByRef PQIndex() As Long, _
                                 ByVal nPQ As Long, _
                                 ByRef mismatch() As Double, _
                                 ByRef maxDP As Double, _
                                 ByRef maxDQ As Double, _
                                 ByRef epsilon As Double, _
                                 ByRef BusBaseKV() As Double)
    Dim i As Long, idx As Long
    Dim dP As Double, dQ As Double

    maxDP = 0#
    maxDQ = 0#

    For i = 1 To nPQ
        idx = PQIndex(i)
        dP = Pspec(idx) - Pcalc(idx)
        mismatch(i) = dP
        If Abs(dP) > maxDP Then maxDP = Abs(dP)
    Next i

    For i = 1 To nPQ
        idx = PQIndex(i)
        If BusTypes(idx) = btPQ Then
            dQ = Qspec(idx) - Qcalc(idx)
            mismatch(nPQ + i) = dQ
            If Abs(dQ) > maxDQ Then maxDQ = Abs(dQ)
        ElseIf BusTypes(idx) = btPV Then
            ' Pre PV uzly je napätie konštantné – rovnica ?Q sa nahrádza ?V = 0.
            mismatch(nPQ + i) = 0#
        End If
    Next i

    epsilon = IIf(maxDP > maxDQ, maxDP, maxDQ)
End Sub

'--------------------------------------
' Zostavenie Jakobiho matice pre neznáme uzly (PQ aj PV)
'--------------------------------------
Private Sub BuildJacobian(ByVal nBuses As Long, _
                          ByRef BusTypes() As BusType, _
                          ByRef G() As Double, _
                          ByRef B() As Double, _
                          ByRef Vmag() As Double, _
                          ByRef Vang() As Double, _
                          ByRef Pcalc() As Double, _
                          ByRef Qcalc() As Double, _
                          ByRef PQIndex() As Long, _
                          ByVal nPQ As Long, _
                          ByRef J() As Double)
    Dim rowPQ As Long, colPQ As Long
    Dim i As Long, k As Long
    Dim costh As Double, sinth As Double
    Dim H As Double, n As Double, m As Double, L As Double
    Dim Vi As Double
    Dim Gik As Double, Bik As Double
    Dim sz As Long, r As Long, c As Long

    Dim cosV() As Double, sinV() As Double
    ReDim cosV(1 To nBuses)
    ReDim sinV(1 To nBuses)
    For i = 1 To nBuses
        cosV(i) = Cos(Vang(i))
        sinV(i) = Sin(Vang(i))
    Next i

    sz = 2 * nPQ
    For r = 1 To sz
        For c = 1 To sz
            J(r, c) = 0#
        Next c
    Next r

    For rowPQ = 1 To nPQ
        i = PQIndex(rowPQ)
        Vi = Vmag(i)
        If Abs(Vi) < 0.000000001 Then Vi = 0.000000001

        H = -Qcalc(i) - B(i, i) * Vi * Vi
        n = Pcalc(i) / Vi + G(i, i) * Vi
        m = Pcalc(i) - G(i, i) * Vi * Vi
        L = Qcalc(i) / Vi - B(i, i) * Vi

        J(rowPQ, rowPQ) = H
        J(rowPQ, nPQ + rowPQ) = n

        If BusTypes(i) = btPQ Then
            J(nPQ + rowPQ, rowPQ) = m
            J(nPQ + rowPQ, nPQ + rowPQ) = L
        ElseIf BusTypes(i) = btPV Then
            J(nPQ + rowPQ, rowPQ) = 0#
            J(nPQ + rowPQ, nPQ + rowPQ) = 1#
        End If

        For colPQ = 1 To nPQ
            If colPQ <> rowPQ Then
                k = PQIndex(colPQ)
                Gik = G(i, k)
                Bik = B(i, k)
                If Gik <> 0# Or Bik <> 0# Then
                    costh = cosV(i) * cosV(k) + sinV(i) * sinV(k)
                    sinth = sinV(i) * cosV(k) - cosV(i) * sinV(k)
                    H = Vmag(i) * Vmag(k) * (Gik * sinth - Bik * costh)
                    n = Vmag(i) * (Gik * costh + Bik * sinth)

                    J(rowPQ, colPQ) = H
                    J(rowPQ, nPQ + colPQ) = n

                    If BusTypes(i) = btPQ Then
                        m = -Vmag(i) * Vmag(k) * (Gik * costh + Bik * sinth)
                        L = Vmag(i) * (Gik * sinth - Bik * costh)
                        J(nPQ + rowPQ, colPQ) = m
                        J(nPQ + rowPQ, nPQ + colPQ) = L
                    End If
                End If
            End If
        Next colPQ
    Next rowPQ
End Sub

'--------------------------------------
' Riešenie lineárnej sústavy J * x = rhs pomocou Gaussovej eliminácie
' s čiastočným pivotovaním. POZOR: J aj rhs sú po návrate prepísané.
' Volá PhaseYield (max raz za ~200 ms) kvôli aktualizácii J5 a responzívnosti UI.
'--------------------------------------
Private Sub SolveLinearSystem_Gauss(ByRef J() As Double, _
                                    ByRef rhs() As Double, _
                                    ByRef solution() As Double)
    Dim n As Long
    Dim i As Long, m As Long, k As Long
    Dim maxRow As Long
    Dim maxValue As Double
    Dim temp As Double
    Dim factor As Double
    Dim Jki As Double, Jii_inv As Double, rhs_i As Double

    On Error GoTo ErrHandler

    n = UBound(J, 1)

    Dim rowi() As Double
    ReDim rowi(1 To n)

    For i = 1 To n
        ' Heartbeat: max raz za ~200 ms aktualizuje časovú bunku J5 a urobí DoEvents
        Call PhaseYield

        maxRow = i
        maxValue = Abs(J(i, i))
        For k = i + 1 To n
            If Abs(J(k, i)) > maxValue Then
                maxValue = Abs(J(k, i))
                maxRow = k
            End If
        Next k

        If maxRow <> i Then
            For k = i To n
                temp = J(i, k)
                J(i, k) = J(maxRow, k)
                J(maxRow, k) = temp
            Next k
            temp = rhs(i)
            rhs(i) = rhs(maxRow)
            rhs(maxRow) = temp
        End If

        If Abs(J(i, i)) < 0.000000000000001 Then
            Err.Raise vbObjectError + 10, "SolveLinearSystem_Gauss", _
                      "Matica je singulárna, sústava nemá riešenie."
        End If

        Jii_inv = 1# / J(i, i)
        rhs_i = rhs(i)
        For m = i + 1 To n
            rowi(m) = J(i, m)
        Next m

        For k = i + 1 To n
            Jki = J(k, i)
            If Jki <> 0# Then
                factor = Jki * Jii_inv
                rhs(k) = rhs(k) - factor * rhs_i
                For m = i + 1 To n
                    J(k, m) = J(k, m) - factor * rowi(m)
                Next m
            End If
        Next k
    Next i

    ReDim solution(1 To n)
    For i = n To 1 Step -1
        temp = rhs(i)
        For m = i + 1 To n
            temp = temp - J(i, m) * solution(m)
        Next m
        solution(i) = temp / J(i, i)
    Next i

    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 10, , "Chyba pri riešení sústavy: " & Err.Description
End Sub

'--------------------------------------
' Aktualizácia stavu napätí
'--------------------------------------
Private Sub UpdateState(ByRef Vmag() As Double, _
                        ByRef Vang() As Double, _
                        ByRef PQIndex() As Long, _
                        ByVal nPQ As Long, _
                        ByRef deltaX() As Double)
    Dim i As Long, idx As Long

    For i = 1 To nPQ
        idx = PQIndex(i)
        Vang(idx) = Vang(idx) + deltaX(i)
    Next i

    For i = 1 To nPQ
        idx = PQIndex(i)
        Vmag(idx) = Vmag(idx) + deltaX(nPQ + i)
    Next i
End Sub

'--------------------------------------
' Hlavná fáza load-flow: NR slučka + všetky post-NR zápisy.
' Predpokladá:
'   - Načítané a v p.u. prepočítané vstupné dáta
'   - Vybudované G(), B() (z BuildYBus)
'   - Vyhodnotenú topológiu (IsBusIsolated, atď.)
'   - Vynulované Pspec/Qspec/Vmag pre izolované uzly (urobí runCALC)
'   - Spustený BeginPhaseTimer(J5) – PhaseYield bude tikať J5
'   - Bunka iterCell (index!I6) na zápis čísla iterácie
'
' maxIter a epsLimit číta z index!B3, B4 (ako predtým).
'--------------------------------------
Public Sub RunNRPhase( _
    ByVal SBase_MVA As Double, ByVal nBuses As Long, ByVal nBusReal As Long, ByRef BusNames() As String, _
    ByRef BusTypes() As BusType, ByRef BusBaseKV() As Double, _
    ByRef Vmag() As Double, ByRef Vang() As Double, ByRef Pspec() As Double, ByRef Qspec() As Double, _
    ByRef G() As Double, ByRef B() As Double, _
    ByVal nBranches As Long, ByRef FromBus() As Long, ByRef ToBus() As Long, _
    ByRef R() As Double, ByRef X() As Double, ByRef BranchStatus() As Integer, ByRef Bshunt() As Double, _
    ByVal nSwitches As Long, ByRef SwFrom() As Long, ByRef SwTo() As Long, _
    ByRef SwR() As Double, ByRef SwX() As Double, ByRef SwStatus() As Integer, _
    ByVal nTrafo As Long, ByRef TrFrom() As Long, ByRef TrTo() As Long, _
    ByRef TrR() As Double, ByRef TrX() As Double, ByRef TrG() As Double, ByRef TrB() As Double, ByRef TrRatio() As Double, _
    ByVal nReaktory As Long, ByRef ReaktorFrom() As Long, ByRef ReaktorTo() As Long, ByRef ReaktorR() As Double, ByRef ReaktorX() As Double, _
    ByVal nDifReaktory As Long, ByRef DifReaktorFrom() As Long, ByRef DifReaktorTo() As Long, ByRef DifReaktorR() As Double, ByRef DifReaktorX() As Double, _
    ByVal nComp As Long, ByRef CompBus() As Long, ByRef CompB() As Double, ByRef CompStatus() As Integer, _
    ByVal nMotors As Long, ByRef MotorBus() As Long, ByRef MotorR() As Double, ByRef MotorG() As Double, ByRef MotorB() As Double, ByRef MotorStatus() As Integer, _
    ByRef IsBusIsolated() As Boolean, ByVal iterCell As Range)

    Dim Pcalc() As Double, Qcalc() As Double
    Dim PQIndex() As Long
    Dim nPQ As Long, slackIndex As Long
    Dim i As Long, k As Long, idx1 As Long, idx2 As Long
    Dim maxIter As Long
    Dim epsLimit As Double
    Dim iter As Long, iterUsed As Long
    Dim mismatch() As Double
    Dim J() As Double
    Dim deltaX() As Double
    Dim maxDP As Double, maxDQ As Double, eps As Double
    Dim converged As Boolean
    Dim Vi As Complex, Vj As Complex, Zs As Complex, I_pu As Complex
    Dim Ubase As Double, Ibase_A As Double, SwCurrent_A() As Double

    ' Parametre z listu index: B3 = max iter, B4 = epsilon limit
    With ThisWorkbook.Worksheets("index")
        maxIter = CLng(ParseDouble(.Range("B3").Value))
        If maxIter <= 0 Then maxIter = 20
        epsLimit = ParseDouble(.Range("B4").Value)
        If epsLimit <= 0 Then epsLimit = 0.000001
    End With

    ' Identifikácia slack a počet neznámych uzlov
    slackIndex = 0
    For i = 1 To nBuses
        If BusTypes(i) = btSlack Then
            If slackIndex <> 0 Then
                Err.Raise vbObjectError + 11, , "Viac ako jeden slack uzol."
            End If
            slackIndex = i
        End If
    Next i

    If slackIndex = 0 Then
        Err.Raise vbObjectError + 12, , "Nenájdený slack uzol."
    End If

    nPQ = 0
    For i = 1 To nBuses
        If BusTypes(i) <> btSlack And Not IsBusIsolated(i) Then
            nPQ = nPQ + 1
        End If
    Next i

    converged = False
    iterUsed = 0
    eps = 0#

    If nPQ = 0 Then
        ' Len slack uzol – netreba iterovať
        converged = True
        GoTo SkipNR
    End If

    ReDim PQIndex(1 To nPQ)
    k = 0
    For i = 1 To nBuses
        If BusTypes(i) <> btSlack And Not IsBusIsolated(i) Then
            k = k + 1
            PQIndex(k) = i
        End If
    Next i

    ReDim Pcalc(1 To nBuses)
    ReDim Qcalc(1 To nBuses)
    ReDim mismatch(1 To 2 * nPQ)
    ReDim J(1 To 2 * nPQ, 1 To 2 * nPQ)

    ' Príprava výsledkových listov (napatia, epsilon)
    Call ClearResultsSheets

    ' Buffre pre log napätí a epsilon (flushneme jedným Range.Value zápisom po slučke)
    Dim VBuf As Variant, EBuf As Variant
    Dim VRow As Long, ERow As Long
    Dim bIdx As Long
    ReDim VBuf(1 To maxIter * nBuses, 1 To 4)
    ReDim EBuf(1 To maxIter, 1 To 4)
    VRow = 0: ERow = 0

    ' NR iteračný cyklus
    For iter = 1 To maxIter
        iterUsed = iter

        ' Aktuálna iterácia do panelu (index!I6)
        Call WritePhaseIter(iterCell, iter)

        ' P, Q
        Call CalcPower(nBuses, G, B, Vmag, Vang, Pcalc, Qcalc)

        ' Mismatch + epsilon
        Call BuildMismatchVectors(nBuses, BusTypes, Pspec, Qspec, Vmag, Pcalc, Qcalc, PQIndex, nPQ, mismatch, maxDP, maxDQ, eps, BusBaseKV)

        ' Tick časovej bunky (J5) na hranici iterácie
        Call PhaseYield

        ' Log iterácie do buffrov
        For bIdx = 1 To nBuses
            VRow = VRow + 1
            VBuf(VRow, 1) = iter
            VBuf(VRow, 2) = BusNames(bIdx)
            VBuf(VRow, 3) = Vmag(bIdx)
            VBuf(VRow, 4) = Vang(bIdx) * RAD2DEG
        Next bIdx
        ERow = ERow + 1
        EBuf(ERow, 1) = iter
        EBuf(ERow, 2) = maxDP
        EBuf(ERow, 3) = maxDQ
        EBuf(ERow, 4) = eps

        ' Konvergencia?
        If eps < epsLimit Then
            converged = True
            Exit For
        End If

        ' Jakobián + riešenie + update
        Call BuildJacobian(nBuses, BusTypes, G, B, Vmag, Vang, Pcalc, Qcalc, PQIndex, nPQ, J)
        Call SolveLinearSystem_Gauss(J, mismatch, deltaX)
        Call UpdateState(Vmag, Vang, PQIndex, nPQ, deltaX)
    Next iter

    ' Flush log buffrov (jeden Range.Value zápis na list)
    Call FlushVoltageLog(VBuf, VRow)
    Call FlushEpsilonLog(EBuf, ERow)

    ' Ak NR neskonvergovalo, hlásime to volajúcemu ako chybu (runCALC označí I5 ako "chyba")
    If Not converged Then
        Err.Raise vbObjectError + 13, , _
            "Newton-Raphson neskonvergoval (iter=" & iterUsed & _
            ", eps=" & Format(eps, "0.000000") & ", limit=" & Format(epsLimit, "0.000000") & ")."
    End If

SkipNR:

    ' Post-NR zápisy na listy
    Call WriteFinalVoltagesToUzly(Vmag, Vang, BusBaseKV, nBusReal)

    ' Pre izolované uzly prepíšeme stĺpce H/I na "izolovane"/"-"
    ' (WriteFinalVoltagesToUzly inak pre nich zapíše 0; report sheet sa už vyplnil v I3 fáze).
    Dim wsUz As Worksheet
    Set wsUz = ThisWorkbook.Worksheets("uzly")
    For i = 1 To nBusReal
        If IsBusIsolated(i) Then
            wsUz.Cells(2 + i, 8).Value = "izolovane"
            wsUz.Cells(2 + i, 9).Value = "-"
        End If
    Next i

    Call WriteBranchCurrents(nBranches, FromBus, ToBus, R, X, BranchStatus, Vmag, Vang, SBase_MVA, BusBaseKV, Bshunt)

    ' Prúdy spínačmi
    If nSwitches > 0 Then
        ReDim SwCurrent_A(1 To nSwitches)
        For k = 1 To nSwitches
            If SwStatus(k) > 0 Then
                idx1 = SwFrom(k): idx2 = SwTo(k)
                Vi = CFromPolarRad(Vmag(idx1), Vang(idx1))
                Vj = CFromPolarRad(Vmag(idx2), Vang(idx2))
                Zs = CCreate(SwR(k), SwX(k))
                If Abs(Zs.Re) > 0.000000001 Or Abs(Zs.Im) > 0.000000001 Then
                    I_pu = CDiv(CSub(Vi, Vj), Zs)
                Else
                    I_pu = CCreate(0, 0)
                End If
                Ubase = BusBaseKV(idx1)
                If Ubase <> 0# Then
                    Ibase_A = (SBase_MVA * 1000#) / (Sqr(3) * Ubase)
                Else
                    Ibase_A = 0#
                End If
                SwCurrent_A(k) = CAbs(I_pu) * Ibase_A
            Else
                SwCurrent_A(k) = 0#
            End If
        Next k
        Call WriteSwitchResults(nSwitches, SwCurrent_A)
    End If

    Call WriteTransformerFlows(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, Vmag, Vang, BusBaseKV, SBase_MVA)
    Call WriteReactorResults(nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)
    Call WriteDifReactorResults(nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)
    Call WriteCompResults(nComp, CompBus, Vmag, BusBaseKV)
    Call WriteMotorResults(nMotors, MotorBus, MotorR, MotorG, MotorB, MotorStatus, Vmag, BusBaseKV, SBase_MVA)

    ' Celkové zaťaženie uzlov (P, Q, I) do stĺpcov K, L, M na liste uzly.
    ' Pcalc/Qcalc môže byť ReDim-nuté na 0 keď nPQ=0 (len slack) – pre WriteNodeThroughput
    ' potrebujeme pole 1..nBuses. V takom prípade ho doplníme nulami.
    If nPQ = 0 Then
        ReDim Pcalc(1 To nBuses)
        ReDim Qcalc(1 To nBuses)
    End If
    Call WriteNodeThroughput(nBusReal, BusNames, BusBaseKV, SBase_MVA, _
                             nBranches, FromBus, ToBus, R, X, BranchStatus, _
                             nTrafo, TrFrom, TrTo, TrR, TrX, TrRatio, TrG, TrB, _
                             nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                             nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                             nComp, CompBus, CompB, CompStatus, _
                             nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                             Vmag, Vang, Pcalc, Qcalc)
End Sub

'--------------------------------------
' Výpočet a zápis zaťaženia uzlov (P, Q, I) do stĺpcov K, L, M
'--------------------------------------
Private Sub WriteNodeThroughput( _
    ByVal nBuses As Long, ByRef BusNames() As String, ByRef BusBaseKV() As Double, ByVal SBase_MVA As Double, _
    ByVal nBranches As Long, ByRef FromBus() As Long, ByRef ToBus() As Long, ByRef R() As Double, ByRef X() As Double, ByRef BranchStatus() As Integer, _
    ByVal nTrafo As Long, ByRef TrFrom() As Long, ByRef TrTo() As Long, ByRef TrR() As Double, ByRef TrX() As Double, ByRef TrRatio() As Double, ByRef TrG() As Double, ByRef TrB() As Double, _
    ByVal nReaktory As Long, ByRef ReaktorFrom() As Long, ByRef ReaktorTo() As Long, ByRef ReaktorR() As Double, ByRef ReaktorX() As Double, _
    ByVal nDifReaktory As Long, ByRef DifReaktorFrom() As Long, ByRef DifReaktorTo() As Long, ByRef DifReaktorR() As Double, ByRef DifReaktorX() As Double, _
    ByVal nComp As Long, ByRef CompBus() As Long, ByRef CompB() As Double, ByRef CompStatus() As Integer, _
    ByVal nMotors As Long, ByRef MotorBus() As Long, ByRef MotorG() As Double, ByRef MotorB() As Double, ByRef MotorStatus() As Integer, _
    ByRef Vmag() As Double, ByRef Vang() As Double, ByRef Pcalc() As Double, ByRef Qcalc() As Double)

    Dim i As Long, k As Long
    Dim SumP() As Double, SumQ() As Double, SumI() As Double
    Dim ws As Worksheet

    ReDim SumP(1 To nBuses), SumQ(1 To nBuses), SumI(1 To nBuses)

    Dim Vi As Complex, Vj As Complex, Z As Complex, Ys As Complex
    Dim I_pu As Complex, S_pu As Complex
    Dim I_abs_pu As Double
    Dim P_flow As Double, Q_flow As Double
    Dim Ubase As Double, Ibase_A As Double

    ' 1. Vedenia
    For k = 1 To nBranches
        If BranchStatus(k) > 0 Then
            Call CalcBranchFlow(k, FromBus(k), ToBus(k), R(k), X(k), Vmag, Vang, Vi, Vj, Z, Ys, I_pu, S_pu)

            P_flow = -S_pu.Re: Q_flow = -S_pu.Im
            I_abs_pu = CAbs(I_pu)
            If P_flow > 0 Then SumP(FromBus(k)) = SumP(FromBus(k)) + P_flow
            If Q_flow > 0 Then SumQ(FromBus(k)) = SumQ(FromBus(k)) + Q_flow
            SumI(FromBus(k)) = SumI(FromBus(k)) + I_abs_pu

            Dim I_ji As Complex, S_ji As Complex
            I_ji = CCreate(-I_pu.Re, -I_pu.Im)
            S_ji = CMul(Vj, CConj(I_ji))
            P_flow = -S_ji.Re: Q_flow = -S_ji.Im
            I_abs_pu = CAbs(I_ji)
            If P_flow > 0 Then SumP(ToBus(k)) = SumP(ToBus(k)) + P_flow
            If Q_flow > 0 Then SumQ(ToBus(k)) = SumQ(ToBus(k)) + Q_flow
            SumI(ToBus(k)) = SumI(ToBus(k)) + I_abs_pu
        End If
    Next k

    ' 2. Trafá
    For k = 1 To nTrafo
        Call CalcTrafoFlow(k, TrFrom(k), TrTo(k), TrR(k), TrX(k), TrRatio(k), TrG(k), TrB(k), Vmag, Vang, _
                           Vi, Vj, I_pu, S_pu)

        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(TrFrom(k)) = SumP(TrFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(TrFrom(k)) = SumQ(TrFrom(k)) + Q_flow
        SumI(TrFrom(k)) = SumI(TrFrom(k)) + I_abs_pu

        Dim Zs As Complex, ys_t As Complex, Yseries_a As Complex
        Zs = CCreate(TrR(k), TrX(k)): ys_t = CDiv(CCreate(1, 0), Zs)
        Yseries_a = CCreate(ys_t.Re / TrRatio(k), ys_t.Im / TrRatio(k))

        Dim term1 As Complex, term2 As Complex, I_sec As Complex, S_sec As Complex
        term1 = CMul(Vj, ys_t)
        term2 = CMul(Vi, Yseries_a)
        I_sec = CSub(term1, term2)
        S_sec = CMul(Vj, CConj(I_sec))

        P_flow = -S_sec.Re: Q_flow = -S_sec.Im
        I_abs_pu = CAbs(I_sec)
        If P_flow > 0 Then SumP(TrTo(k)) = SumP(TrTo(k)) + P_flow
        If Q_flow > 0 Then SumQ(TrTo(k)) = SumQ(TrTo(k)) + Q_flow
        SumI(TrTo(k)) = SumI(TrTo(k)) + I_abs_pu
    Next k

    ' 3. Reaktory
    For k = 1 To nReaktory
        Call CalcBranchFlow(k, ReaktorFrom(k), ReaktorTo(k), ReaktorR(k), ReaktorX(k), Vmag, Vang, Vi, Vj, Z, Ys, I_pu, S_pu)

        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(ReaktorFrom(k)) = SumP(ReaktorFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(ReaktorFrom(k)) = SumQ(ReaktorFrom(k)) + Q_flow
        SumI(ReaktorFrom(k)) = SumI(ReaktorFrom(k)) + I_abs_pu

        I_ji = CCreate(-I_pu.Re, -I_pu.Im)
        S_ji = CMul(Vj, CConj(I_ji))
        P_flow = -S_ji.Re: Q_flow = -S_ji.Im
        I_abs_pu = CAbs(I_ji)
        If P_flow > 0 Then SumP(ReaktorTo(k)) = SumP(ReaktorTo(k)) + P_flow
        If Q_flow > 0 Then SumQ(ReaktorTo(k)) = SumQ(ReaktorTo(k)) + Q_flow
        SumI(ReaktorTo(k)) = SumI(ReaktorTo(k)) + I_abs_pu
    Next k

    ' 4. Dif Reaktory
    For k = 1 To nDifReaktory
        Call CalcBranchFlow(k, DifReaktorFrom(k), DifReaktorTo(k), DifReaktorR(k), DifReaktorX(k), Vmag, Vang, Vi, Vj, Z, Ys, I_pu, S_pu)

        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(DifReaktorFrom(k)) = SumP(DifReaktorFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(DifReaktorFrom(k)) = SumQ(DifReaktorFrom(k)) + Q_flow
        SumI(DifReaktorFrom(k)) = SumI(DifReaktorFrom(k)) + I_abs_pu

        I_ji = CCreate(-I_pu.Re, -I_pu.Im)
        S_ji = CMul(Vj, CConj(I_ji))
        P_flow = -S_ji.Re: Q_flow = -S_ji.Im
        I_abs_pu = CAbs(I_ji)
        If P_flow > 0 Then SumP(DifReaktorTo(k)) = SumP(DifReaktorTo(k)) + P_flow
        If Q_flow > 0 Then SumQ(DifReaktorTo(k)) = SumQ(DifReaktorTo(k)) + Q_flow
        SumI(DifReaktorTo(k)) = SumI(DifReaktorTo(k)) + I_abs_pu
    Next k

    ' 5. Kompenzácia (Shunt)
    For k = 1 To nComp
        If CompStatus(k) = 1 Then
            i = CompBus(k)
            Dim V_sq As Double
            V_sq = Vmag(i) * Vmag(i)
            P_flow = 0
            Q_flow = V_sq * CompB(k)
            I_abs_pu = Vmag(i) * Abs(CompB(k))

            If P_flow > 0 Then SumP(i) = SumP(i) + P_flow
            If Q_flow > 0 Then SumQ(i) = SumQ(i) + Q_flow
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next k

    ' 6. Motory (Shunt)
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            i = MotorBus(k)
            V_sq = Vmag(i) * Vmag(i)
            P_flow = -V_sq * MotorG(k)
            Q_flow = V_sq * MotorB(k)
            I_abs_pu = Vmag(i) * Sqr(MotorG(k) * MotorG(k) + MotorB(k) * MotorB(k))

            If P_flow > 0 Then SumP(i) = SumP(i) + P_flow
            If Q_flow > 0 Then SumQ(i) = SumQ(i) + Q_flow
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next k

    ' 7. Injekcia do uzla (Generátory / Odbery)
    For i = 1 To nBuses
        If Pcalc(i) > 0 Then SumP(i) = SumP(i) + Pcalc(i)
        If Qcalc(i) > 0 Then SumQ(i) = SumQ(i) + Qcalc(i)

        If Vmag(i) > 0.0000001 Then
            I_abs_pu = Sqr(Pcalc(i) * Pcalc(i) + Qcalc(i) * Qcalc(i)) / Vmag(i)
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next i

    Set ws = ThisWorkbook.Worksheets("uzly")
    ws.Cells(2, 11).Value = "Sum P_in [MW]"
    ws.Cells(2, 12).Value = "Sum Q_in [Mvar]"
    ws.Cells(2, 13).Value = "Sum I [A]"

    For i = 1 To nBuses
        Dim P_real As Double, Q_real As Double, I_real As Double
        P_real = SumP(i) * SBase_MVA
        Q_real = SumQ(i) * SBase_MVA

        Ubase = BusBaseKV(i)
        If Ubase <> 0 Then
            Ibase_A = (SBase_MVA * 1000#) / (Sqr(3) * Ubase)
        Else
            Ibase_A = 0
        End If
        I_real = SumI(i) * Ibase_A

        ws.Cells(2 + i, 11).Value = Round(P_real, 2)
        ws.Cells(2 + i, 12).Value = Round(Q_real, 2)
        ws.Cells(2 + i, 13).Value = Round(I_real, 2)
    Next i
End Sub

' Výpočet toku na začiatku vetvy (I_ij, S_ij)
Private Sub CalcBranchFlow(ByVal k As Long, ByVal i As Long, ByVal J As Long, _
                           ByVal R As Double, ByVal X As Double, _
                           ByRef Vmag() As Double, ByRef Vang() As Double, _
                           ByRef Vi As Complex, ByRef Vj As Complex, _
                           ByRef Z As Complex, ByRef Ys As Complex, _
                           ByRef I_pu As Complex, ByRef S_pu As Complex)
    Vi = CFromPolarRad(Vmag(i), Vang(i))
    Vj = CFromPolarRad(Vmag(J), Vang(J))
    Z = CCreate(R, X)
    I_pu = CDiv(CSub(Vi, Vj), Z)
    S_pu = CMul(Vi, CConj(I_pu))
End Sub

' Výpočet toku trafa (primár)
Private Sub CalcTrafoFlow(ByVal k As Long, ByVal i As Long, ByVal J As Long, _
                          ByVal R As Double, ByVal X As Double, ByVal Ratio As Double, _
                          ByVal G As Double, ByVal B As Double, _
                          ByRef Vmag() As Double, ByRef Vang() As Double, _
                          ByRef Vi As Complex, ByRef Vj As Complex, _
                          ByRef I_prim As Complex, ByRef S_prim As Complex)
    Dim Zs As Complex, Ys As Complex, Ym As Complex
    Dim term1 As Complex, term2 As Complex
    Dim Yseries_a2 As Complex, Yseries_a As Complex

    Vi = CFromPolarRad(Vmag(i), Vang(i))
    Vj = CFromPolarRad(Vmag(J), Vang(J))
    Zs = CCreate(R, X)
    Ys = CDiv(CCreate(1, 0), Zs)
    Ym = CCreate(G, B)

    Yseries_a2 = CCreate(Ys.Re / (Ratio * Ratio), Ys.Im / (Ratio * Ratio))
    term1 = CMul(Vi, CAdd(Yseries_a2, Ym))

    Yseries_a = CCreate(Ys.Re / Ratio, Ys.Im / Ratio)
    term2 = CMul(Vj, Yseries_a)

    I_prim = CSub(term1, term2)
    S_prim = CMul(Vi, CConj(I_prim))
End Sub
