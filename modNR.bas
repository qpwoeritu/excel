Option Explicit

'--------------------------------------
' Throttled progress (Možnosť A): Solve a iné horúce slučky volajú
' MaybeYield ktorý max raz za ~200 ms aktualizuje StatusBar a urobí
' DoEvents. Tým sa zabráni stavu "nereaguje" aj keď jedna NR iterácia
' trvá > 5 sekúnd. Stav sa nastavuje v NewtonRaphsonLoadFlow pred slučkou.
'--------------------------------------
Private m_nrIter As Long
Private m_nrMaxIter As Long
Private m_nrSBase As Double
Private m_nrStartTime As Double
Private m_nrLastEps As Double
Private m_lastYield As Double

Private Sub MaybeYield()
    Dim t As Double
    t = Timer
    If t - m_lastYield > 0.2 Then
        Application.StatusBar = NRProgressStr(m_nrIter, m_nrMaxIter, m_nrLastEps, m_nrSBase, t - m_nrStartTime)
        DoEvents
        m_lastYield = t
    End If
End Sub

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

    ' Predpočet trigonometrie raz na uzol (A4):
    '   Cos(Vang(i) - Vang(k)) = cosV(i)*cosV(k) + sinV(i)*sinV(k)
    '   Sin(Vang(i) - Vang(k)) = sinV(i)*cosV(k) - cosV(i)*sinV(k)
    Dim cosV() As Double, sinV() As Double
    ReDim cosV(1 To nBuses)
    ReDim sinV(1 To nBuses)
    For i = 1 To nBuses
        cosV(i) = Cos(Vang(i))
        sinV(i) = Sin(Vang(i))
    Next i

    For i = 1 To nBuses
        Pcalc(i) = 0#
        Qcalc(i) = 0#
        For k = 1 To nBuses
            costh = cosV(i) * cosV(k) + sinV(i) * sinV(k)
            sinth = sinV(i) * cosV(k) - cosV(i) * sinV(k)
            ViVk = Vmag(i) * Vmag(k)
            Pcalc(i) = Pcalc(i) + ViVk * (G(i, k) * costh + B(i, k) * sinth)
            Qcalc(i) = Qcalc(i) + ViVk * (G(i, k) * sinth - B(i, k) * costh)
        Next k
    Next i
End Sub

'--------------------------------------
' Zostavenie vektora nesúladu ?P, ?Q (alebo ?V pre PV)
' ?P = Pspec - Pcalc
' ?Q = Qspec - Qcalc (pre PQ)
' ?V = Vspec - Vcalc (pre PV) - tu používam trik, že ?Q rovnica je nahradená rovnicou pre napätie
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
    Dim Vspec As Double

    maxDP = 0#
    maxDQ = 0#

    ' prvá polovica vektora – ?P (pre všetky neznáme uzly: PQ aj PV)
    For i = 1 To nPQ
        idx = PQIndex(i)
        dP = Pspec(idx) - Pcalc(idx)
        mismatch(i) = dP
        If Abs(dP) > maxDP Then maxDP = Abs(dP)
    Next i

    ' druhá polovica – ?Q (pre PQ) alebo ?V (pre PV)
    For i = 1 To nPQ
        idx = PQIndex(i)
        If BusTypes(idx) = btPQ Then
            ' PQ uzol: ?Q
            dQ = Qspec(idx) - Qcalc(idx)
            mismatch(nPQ + i) = dQ
            If Abs(dQ) > maxDQ Then maxDQ = Abs(dQ)
        ElseIf BusTypes(idx) = btPV Then
            ' Pre PV uzly (generátorové uzly) je napätie konštantné.
            ' Namiesto rovnice pre jalový výkon (Q) použijeme podmienku fixného napätia.
            ' Nastavením mismatch na 0 a jednotkového riadku v Jakobiáne (dV=0)
            ' zabezpečíme, že veľkosť napätia ostane na počiatočnej špecifikovanej hodnote.
            mismatch(nPQ + i) = 0#
        End If
    Next i

    epsilon = IIf(maxDP > maxDQ, maxDP, maxDQ)
End Sub

'--------------------------------------
' Zostavenie Jakobiho matice pre neznáme uzly (PQ aj PV)
' J má rozmery (2*nPQ) x (2*nPQ)
' Pre PV uzly sa riadky M a L menia (M ostáva ak rátame P, L sa nahrádza rovnicou pre dV)
'
' Optimalizácia A4: predpočítané Cos/Sin pre každý uhol uzla;
' súčtové vzorce namiesto opätovných volaní Cos/Sin v inom cykle.
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

    ' Predpočet trigonometrie raz na uzol (A4)
    Dim cosV() As Double, sinV() As Double
    ReDim cosV(1 To nBuses)
    ReDim sinV(1 To nBuses)
    For i = 1 To nBuses
        cosV(i) = Cos(Vang(i))
        sinV(i) = Sin(Vang(i))
    Next i

    For rowPQ = 1 To nPQ
        i = PQIndex(rowPQ)
        For colPQ = 1 To nPQ
            k = PQIndex(colPQ)
            ' Výpočet derivácií (H, N, M, L) je rovnaký pre všetky typy (závisí od fyziky siete)
            ' Až pri zápise do J rozhodneme, či ich použijeme
            If i = k Then
                Vi = Vmag(i)
                If Abs(Vi) < 0.000000001 Then Vi = 0.000000001
                H = -Qcalc(i) - B(i, i) * Vi * Vi
                n = Pcalc(i) / Vi + G(i, i) * Vi
                m = Pcalc(i) - G(i, i) * Vi * Vi
                L = Qcalc(i) / Vi - B(i, i) * Vi
            Else
                ' Cos(Vang(i) - Vang(k)) a Sin(Vang(i) - Vang(k)) cez súčtové vzorce (A4)
                costh = cosV(i) * cosV(k) + sinV(i) * sinV(k)
                sinth = sinV(i) * cosV(k) - cosV(i) * sinV(k)
                H = Vmag(i) * Vmag(k) * (G(i, k) * sinth - B(i, k) * costh)
                n = Vmag(i) * (G(i, k) * costh + B(i, k) * sinth)
                m = -Vmag(i) * Vmag(k) * (G(i, k) * costh + B(i, k) * sinth)
                L = Vmag(i) * (G(i, k) * sinth - B(i, k) * costh)
            End If

            ' H a N bloky (dP/dTheta, dP/dV) sú platné pre PQ aj PV (lebo P je špecifikované pre oba)
            J(rowPQ, colPQ) = H
            J(rowPQ, nPQ + colPQ) = n

            ' M a L bloky (dQ/dTheta, dQ/dV)
            If BusTypes(i) = btPQ Then
                ' Pre PQ uzol: Použijeme štandardné M a L
                J(nPQ + rowPQ, colPQ) = m
                J(nPQ + rowPQ, nPQ + colPQ) = L
            ElseIf BusTypes(i) = btPV Then
                ' Spracovanie PV uzla v Jakobiáne:
                ' Rovnica pre odchýlku jalového výkonu je nahradená podmienkou konštantného napätia (dV = 0).
                ' To dosiahne riadkom s nulami a jednotkou na diagonále v časti dP/dV.
                If i = k Then
                    J(nPQ + rowPQ, colPQ) = 0#         ' dV/dTheta = 0
                    J(nPQ + rowPQ, nPQ + colPQ) = 1#   ' dV/dV    = 1
                Else
                    J(nPQ + rowPQ, colPQ) = 0#
                    J(nPQ + rowPQ, nPQ + colPQ) = 0#
                End If
            End If
        Next colPQ
    Next rowPQ
End Sub

'--------------------------------------
' Riešenie lineárnej sústavy J * x = rhs pomocou Gaussovej eliminácie
' s čiastočným pivotovaním.
'
' Optimalizácia A3: oproti pôvodnej verzii sa NEKOPÍRUJE J do A a rhs do B.
' Pracuje sa priamo s J a rhs in-place. Volajúci ich v ďalšej iterácii znova
' zostavuje cez BuildJacobian / BuildMismatchVectors, takže to nevadí.
' POZOR: po návrate sú J aj rhs prepísané (sú v podstate „odpadové").
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

    On Error GoTo ErrHandler

    n = UBound(J, 1)

    ' Priama eliminácia s čiastočným pivotovaním (in-place na J, rhs)
    For i = 1 To n
        ' Throttled progress + DoEvents (max raz za ~200 ms).
        ' Pre n=1000 sa v outer slučke zavolá 1000-krát, ale len ~50 z toho
        ' fyzicky aktualizuje StatusBar – overhead je zanedbateľný.
        Call MaybeYield

        maxRow = i
        maxValue = Abs(J(i, i))
        For k = i + 1 To n
            If Abs(J(k, i)) > maxValue Then
                maxValue = Abs(J(k, i))
                maxRow = k
            End If
        Next k

        ' Zámena riadkov
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

        ' Kontrola singularity
        If Abs(J(i, i)) < 0.000000000000001 Then
            Err.Raise vbObjectError + 10, "SolveLinearSystem_Gauss", _
                      "Matica je singulárna, sústava nemá riešenie."
        End If

        ' Eliminácia pod pivotom
        For k = i + 1 To n
            factor = J(k, i) / J(i, i)
            rhs(k) = rhs(k) - factor * rhs(i)
            For m = i + 1 To n
                J(k, m) = J(k, m) - factor * J(i, m)
            Next m
        Next k
    Next i

    ' Spätná substitúcia
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

    ' prvá polovica vektora – ??
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vang(idx) = Vang(idx) + deltaX(i)
    Next i

    ' druhá polovica – ?|V|
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vmag(idx) = Vmag(idx) + deltaX(nPQ + i)
    Next i
End Sub
'--------------------------------------
' Hlavný Newton-Raphson load-flow
' Optimalizácia A1: vypnutie automatického prepočtu, obnovy obrazovky
'                   a udalostí počas behu (pôvodný stav sa obnoví).
' Optimalizácia A2: použitie CFromPolarRad (uhol v radiánoch) pre výpočet
'                   prúdov spínačov – eliminuje zbytočnú konverziu rad↔deg.
'--------------------------------------
Public Sub NewtonRaphsonLoadFlow()
    ' --- A1: uchovanie pôvodných nastavení Excelu kvôli zrýchleniu ---
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim appSettingsSaved As Boolean
    appSettingsSaved = False
    ' ----------------------------------------------------------------

    Dim SBase_MVA As Double
    Dim VLevels() As Double  ' Napäťové hladiny zo data!K3:K8
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusBaseKV() As Double  ' Nové pole báz pre uzly
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim BranchName() As String
    Dim R() As Double, X() As Double
    Dim Bshunt() As Double
    Dim BranchStatus() As Integer

    ' Transformátory
    Dim nTrafo As Long
    Dim TrFrom() As Long, TrTo() As Long
    Dim TrR() As Double, TrX() As Double
    Dim TrG() As Double, TrB() As Double
    Dim TrRatio() As Double

    ' Reaktory
    Dim nReaktory As Long
    Dim ReaktorName() As String
    Dim ReaktorFrom() As Long, ReaktorTo() As Long
    Dim ReaktorR() As Double, ReaktorX() As Double

    ' Dif. Reaktory
    Dim nDifReaktory As Long
    Dim DifReaktorName() As String
    Dim DifReaktorFrom() As Long, DifReaktorTo() As Long
    Dim DifReaktorR() As Double, DifReaktorX() As Double

    ' Spínače
    Dim nSwitches As Long
    Dim SwitchName() As String
    Dim SwFrom() As Long, SwTo() As Long
    Dim SwR() As Double, SwX() As Double
    Dim SwStatus() As Integer

    ' Kompenzácia
    Dim nComp As Long
    Dim CompName() As String
    Dim CompBus() As Long
    Dim CompB() As Double
    Dim CompStatus() As Integer

    ' Motory VN
    Dim nMotors As Long
    Dim MotorName() As String
    Dim MotorBus() As Long
    Dim MotorR() As Double
    Dim MotorXk() As Double
    Dim MotorG() As Double
    Dim MotorB() As Double
    Dim MotorStatus() As Integer

    ' Topológia - Izolované časti
    Dim IsBusIsolated() As Boolean
    Dim IsBranchIsolated() As Boolean
    Dim IsTrafoIsolated() As Boolean
    Dim IsReaktorIsolated() As Boolean
    Dim IsDifReaktorIsolated() As Boolean
    Dim IsCompIsolated() As Boolean
    Dim IsMotorIsolated() As Boolean
    Dim isolatedCount As Long

    Dim Y() As Complex
    Dim G() As Double, B() As Double
    Dim Pcalc() As Double, Qcalc() As Double
    Dim PQIndex() As Long  ' Indexy uzlov, pre ktoré rátame rovnice
    Dim nPQ As Long

    Dim slackIndex As Long
    Dim i As Long, k As Long, idx1 As Long, idx2 As Long
    Dim Vi As Complex, Vj As Complex, Zs As Complex, I_pu As Complex
    Dim Ubase As Double, Ibase_A As Double, SwCurrent_A() As Double

    Dim maxIter As Long
    Dim epsLimit As Double
    Dim iter As Long, iterUsed As Long

    Dim mismatch() As Double
    Dim J() As Double
    Dim deltaX() As Double
    Dim maxDP As Double, maxDQ As Double, eps As Double

    Dim startTime As Double, totalTime As Double
    Dim converged As Boolean

    On Error GoTo ErrHandler

    ' --- A1: vypnutie automatiky Excelu počas výpočtu ---
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    appSettingsSaved = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' ----------------------------------------------------

    '--------------------------
    ' Načítanie parametrov z listu index
    ' B3 – max. počet iterácií
    ' B4 – epsilon limit
    '--------------------------
    With ThisWorkbook.Worksheets("index")
        maxIter = CLng(ParseDouble(.Range("B3").Value))
        If maxIter <= 0 Then maxIter = 20
        epsLimit = ParseDouble(.Range("B4").Value)
        If epsLimit <= 0 Then epsLimit = 0.000001
    End With

    '--------------------------
    ' Načítanie uzlov a vedení
    '--------------------------
    ' načítanie bázových hodnôt
    Call GetBaseValues(SBase_MVA, VLevels)

    Dim busDict As Object

    ' načítanie uzlov (skutočné -> p.u.)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, VLevels, busDict)

    ' načítanie vedení (ohm -> p.u.)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, Bshunt, busDict)

    ' načítanie transformátorov (ohm/siemens -> p.u.)
    Call LoadTransformerData(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, BusNames, BusBaseKV, SBase_MVA, busDict)

    ' načítanie reaktorov (ohm -> p.u.)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)

    ' načítanie dif. reaktorov
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)

    ' načítanie spínačov
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA, busDict)

    ' načítanie kompenzácie
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA, busDict)

    ' načítanie motorov
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA, busDict)

    '--------------------------
    ' Identifikácia izolovaných častí
    '--------------------------
    Dim IsSwitchIsolated() As Boolean
    Call FindIsolatedParts(nBuses, nBranches, FromBus, ToBus, BranchStatus, _
                           nTrafo, TrFrom, TrTo, _
                           nReaktory, ReaktorFrom, ReaktorTo, _
                           nDifReaktory, DifReaktorFrom, DifReaktorTo, _
                           nSwitches, SwFrom, SwTo, SwStatus, _
                           nComp, CompBus, _
                           nMotors, MotorBus, _
                           BusTypes, _
                           IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, IsCompIsolated, IsMotorIsolated, isolatedCount)

    ' Ak je izolovaný uzol, vynuluj jeho Pspec a Qspec, aby nerobil problémy v mismatch vektore
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Pspec(i) = 0#
            Qspec(i) = 0#
            ' Vynulujeme aj počiatočné napätie
            Vmag(i) = 0#
        End If
    Next i

    ' tvorba Y-matice v p.u. (s ignorovaním izolovaných)
    Call BuildYBus(nBuses, nBranches, FromBus, ToBus, R, X, BranchStatus, Bshunt, _
                   nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                   nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, _
                   nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                   nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                   nComp, CompBus, CompB, CompStatus, _
                   nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                   BusNames, IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
                   Y, G, B)

    '--------------------------
    ' Identifikácia slack a PQ uzlov
    '--------------------------
    slackIndex = 0
    nPQ = 0
    For i = 1 To nBuses
        If BusTypes(i) = btSlack Then
            If slackIndex <> 0 Then
                Err.Raise vbObjectError + 11, , "Viac ako jeden slack uzol."
            End If
            slackIndex = i
        ElseIf BusTypes(i) = btPQ Then
            nPQ = nPQ + 1
        End If
    Next i

    If slackIndex = 0 Then
        Err.Raise vbObjectError + 12, , "Nenájdený slack uzol."
    End If

    ' Zrátame všetky non-Slack uzly
    nPQ = 0
    For i = 1 To nBuses
        If BusTypes(i) <> btSlack And Not IsBusIsolated(i) Then
            nPQ = nPQ + 1
        End If
    Next i

    If nPQ = 0 Then
        ' Len slack - skončené
        converged = True
        GoTo SkipNR
    End If

    ReDim PQIndex(1 To nPQ)
    k = 0
    For i = 1 To nBuses
        If BusTypes(i) <> btSlack Then
            If Not IsBusIsolated(i) Then
                k = k + 1
                PQIndex(k) = i
            End If
        End If
    Next i

    ' Aktualizácia nPQ (skutočný počet počítaných uzlov)
    nPQ = k
    If nPQ > 0 Then ReDim Preserve PQIndex(1 To nPQ)

    ReDim Pcalc(1 To nBuses)
    ReDim Qcalc(1 To nBuses)

    If nPQ > 0 Then
        ReDim mismatch(1 To 2 * nPQ)
        ReDim J(1 To 2 * nPQ, 1 To 2 * nPQ)
    End If

    '--------------------------
    ' Príprava výsledkových listov
    '--------------------------
    Call ClearResultsSheets

    ' Bufre pre log napätí a epsilon (flushneme jedným Range.Value zápisom po slučke)
    Dim VBuf As Variant, EBuf As Variant
    Dim VRow As Long, ERow As Long
    Dim bIdx As Long
    ReDim VBuf(1 To maxIter * nBuses, 1 To 4)
    ReDim EBuf(1 To maxIter, 1 To 4)
    VRow = 0: ERow = 0

    startTime = Timer
    converged = False
    iterUsed = 0

    Application.DisplayStatusBar = True
    Application.StatusBar = "Load Flow NR  [--------------------]  čakajte..."

    ' Setup throttled progress – Solve volá MaybeYield ktorý čítä tieto premenné
    m_nrMaxIter = maxIter
    m_nrSBase = SBase_MVA
    m_nrStartTime = startTime
    m_nrLastEps = 0#
    m_lastYield = Timer

    ' ScreenUpdating = True počas NR slučky, aby StatusBar fyzicky prekreslil.
    ' V slučke sa nečíta ani nepíše do buniek, takže to nezpomalí výpočet.
    ' Pred post-NR zápismi sa znova vypne.
    Application.ScreenUpdating = True

    '--------------------------
    ' Hlavný NR iteračný cyklus
    '--------------------------
    If nPQ > 0 Then
        For iter = 1 To maxIter
            iterUsed = iter
            m_nrIter = iter

            ' výpočet P, Q
            Call CalcPower(nBuses, G, B, Vmag, Vang, Pcalc, Qcalc)

            ' vektor nesúladu a epsilon (upravený pre PV)
            Call BuildMismatchVectors(nBuses, BusTypes, Pspec, Qspec, Vmag, Pcalc, Qcalc, PQIndex, nPQ, mismatch, maxDP, maxDQ, eps, BusBaseKV)
            m_nrLastEps = eps

            ' Progress bar v StatusBar + DoEvents na hranici iterácie (garantovaný update)
            Application.StatusBar = NRProgressStr(iter, maxIter, eps, SBase_MVA, Timer - startTime)
            DoEvents
            m_lastYield = Timer

            ' logovanie napätí a epsilon do bufrov (flush po slučke)
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

            ' kontrola konvergencie
            If eps < epsLimit Then
                converged = True
                Exit For
            End If

            ' Jakobiho matica
            Call BuildJacobian(nBuses, BusTypes, G, B, Vmag, Vang, Pcalc, Qcalc, PQIndex, nPQ, J)

            ' riešenie J * ?x = mismatch  (POZOR: J a mismatch sú po návrate prepísané)
            Call SolveLinearSystem_Gauss(J, mismatch, deltaX)

            ' aktualizácia napätí
            Call UpdateState(Vmag, Vang, PQIndex, nPQ, deltaX)
        Next iter
    Else
        converged = True
    End If

    ' Vypneme ScreenUpdating pre post-NR zápisy (veľa buniek – výrazne zrýchli zápis).
    Application.ScreenUpdating = False

    ' Flush log bufrov (jeden Range.Value zápis na list)
    Call FlushVoltageLog(VBuf, VRow)
    Call FlushEpsilonLog(EBuf, ERow)

SkipNR:
    totalTime = Timer - startTime
    Call WriteSummaryToIndex(totalTime, iterUsed, eps, converged)

    ' zapíš výsledné napätia na list "uzly" v skutočných hodnotách [kV]
    Call WriteFinalVoltagesToUzly(Vmag, Vang, BusBaseKV)

    ' vypočítaj a zapíš prúdy vo vedeniach v reálnych hodnotách
    Call WriteBranchCurrents(nBranches, FromBus, ToBus, R, X, BranchStatus, Vmag, Vang, SBase_MVA, BusBaseKV, Bshunt)

    ' Výpočet prúdov spínačmi
    If nSwitches > 0 Then
        ReDim SwCurrent_A(1 To nSwitches)
        For k = 1 To nSwitches
            If SwStatus(k) > 0 Then
                idx1 = SwFrom(k): idx2 = SwTo(k)
                ' A2: uhly sú v radiánoch, používame CFromPolarRad (bez zbytočnej rad->deg->rad konverzie)
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

    ' vypočítaj a zapíš toky v transformátoroch
    Call WriteTransformerFlows(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, Vmag, Vang, BusBaseKV, SBase_MVA)

    ' vypočítaj a zapíš toky v reaktoroch
    Call WriteReactorResults(nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)

    ' vypočítaj a zapíš toky v dif. reaktoroch
    Call WriteDifReactorResults(nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)

    ' zapíš výsledky kompenzácie
    Call WriteCompResults(nComp, CompBus, Vmag, BusBaseKV)

    ' zapíš výsledky motorov
    Call WriteMotorResults(nMotors, MotorBus, MotorR, MotorG, MotorB, MotorStatus, Vmag, BusBaseKV, SBase_MVA)

    ' Výpočet a zápis celkového zaťaženia uzlov (P, Q, I)
    Call WriteNodeThroughput(nBuses, BusNames, BusBaseKV, SBase_MVA, _
                             nBranches, FromBus, ToBus, R, X, BranchStatus, _
                             nTrafo, TrFrom, TrTo, TrR, TrX, TrRatio, TrG, TrB, _
                             nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                             nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                             nComp, CompBus, CompB, CompStatus, _
                             nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                             Vmag, Vang, Pcalc, Qcalc)

    ' Report izolovaných (volaný až na konci, aby prepísal stĺpec H na "izolovane")
    Call WriteIsolationReport(nBuses, BusNames, IsBusIsolated, _
                              nBranches, FromBus, ToBus, IsBranchIsolated, _
                              nTrafo, TrFrom, TrTo, IsTrafoIsolated, _
                              nComp, CompBus, IsCompIsolated)

    ' Aktualizácia SLD
    Call UpdateSLD

    ' --- A1: obnova pôvodných nastavení Excelu ---
    If appSettingsSaved Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreen
        Application.EnableEvents = prevEvents
    End If
    Application.StatusBar = False
    ' --------------------------------------------

    Exit Sub

ErrHandler:
    ' --- A1: obnova nastavení aj v prípade chyby, aby Excel neostal v manuálnom režime ---
    If appSettingsSaved Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreen
        Application.EnableEvents = prevEvents
    End If
    Application.StatusBar = False
    ' --------------------------------------------------------------------------------------
    MsgBox "Chyba v Newton-Raphson výpočte: " & Err.Description, vbCritical
End Sub
'--------------------------------------
' Textový progress bar pre StatusBar
' Príklad: "Load Flow NR  [=========-----------]  iter 4/20  |  eps = 0.023 MW  |  12.3 s"
'--------------------------------------
Private Function NRProgressStr(ByVal iter As Long, ByVal maxIter As Long, _
                               ByVal eps As Double, ByVal SBase_MVA As Double, _
                               ByVal elapsed As Double) As String
    Const BAR_LEN As Long = 20
    Dim filled As Long
    filled = CLng(CDbl(iter) / CDbl(maxIter) * BAR_LEN)
    If filled > BAR_LEN Then filled = BAR_LEN
    Dim bar As String
    bar = "[" & String(filled, "=") & String(BAR_LEN - filled, "-") & "]"
    NRProgressStr = "Load Flow NR  " & bar & _
                    "  iter " & iter & "/" & maxIter & _
                    "  |  eps = " & Format(eps * SBase_MVA, "0.000") & " MW" & _
                    "  |  " & Format(elapsed, "0.0") & " s"
End Function

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

    ' Pomocné pre výpočty
    Dim Vi As Complex, Vj As Complex, Z As Complex, Ys As Complex
    Dim I_pu As Complex, S_pu As Complex
    Dim I_abs_pu As Double
    Dim P_flow As Double, Q_flow As Double
    Dim Ubase As Double, Ibase_A As Double

    ' 1. Vedenia
    For k = 1 To nBranches
        ' Len zapnuté vedenia
        If BranchStatus(k) > 0 Then
            ' Uzol i -> j
            Call CalcBranchFlow(k, FromBus(k), ToBus(k), R(k), X(k), Vmag, Vang, Vi, Vj, Z, Ys, I_pu, S_pu)

            ' Tok z i do vedenia (S_pu)
            P_flow = -S_pu.Re: Q_flow = -S_pu.Im
            I_abs_pu = CAbs(I_pu)

            If P_flow > 0 Then SumP(FromBus(k)) = SumP(FromBus(k)) + P_flow
            If Q_flow > 0 Then SumQ(FromBus(k)) = SumQ(FromBus(k)) + Q_flow
            SumI(FromBus(k)) = SumI(FromBus(k)) + I_abs_pu

            ' Uzol j -> i (opačný tok)
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

        ' Primár (i)
        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(TrFrom(k)) = SumP(TrFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(TrFrom(k)) = SumQ(TrFrom(k)) + Q_flow
        SumI(TrFrom(k)) = SumI(TrFrom(k)) + I_abs_pu

        ' Sekundár (j)
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

        ' Uzol i
        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(ReaktorFrom(k)) = SumP(ReaktorFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(ReaktorFrom(k)) = SumQ(ReaktorFrom(k)) + Q_flow
        SumI(ReaktorFrom(k)) = SumI(ReaktorFrom(k)) + I_abs_pu

        ' Uzol j
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

        ' Uzol i
        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(DifReaktorFrom(k)) = SumP(DifReaktorFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(DifReaktorFrom(k)) = SumQ(DifReaktorFrom(k)) + Q_flow
        SumI(DifReaktorFrom(k)) = SumI(DifReaktorFrom(k)) + I_abs_pu

        ' Uzol j
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
        If Pcalc(i) > 0 Then
            SumP(i) = SumP(i) + Pcalc(i)
        End If
        If Qcalc(i) > 0 Then
            SumQ(i) = SumQ(i) + Qcalc(i)
        End If

        If Vmag(i) > 0.0000001 Then
            I_abs_pu = Sqr(Pcalc(i) * Pcalc(i) + Qcalc(i) * Qcalc(i)) / Vmag(i)
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next i

    ' Zápis do listu "uzly"
    Set ws = ThisWorkbook.Worksheets("uzly")

    ' Hlavičky
    ws.Cells(2, 11).Value = "Sum P_in [MW]"     ' K
    ws.Cells(2, 12).Value = "Sum Q_in [Mvar]"   ' L
    ws.Cells(2, 13).Value = "Sum I [A]"         ' M

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

' Pomocná: Výpočet toku na začiatku vetvy (I_ij, S_ij)
' A2: použitý CFromPolarRad – odpadá konverzia uhla radiány -> stupne -> radiány.
Private Sub CalcBranchFlow(ByVal k As Long, ByVal i As Long, ByVal J As Long, _
                           ByVal R As Double, ByVal X As Double, _
                           ByRef Vmag() As Double, ByRef Vang() As Double, _
                           ByRef Vi As Complex, ByRef Vj As Complex, _
                           ByRef Z As Complex, ByRef Ys As Complex, _
                           ByRef I_pu As Complex, ByRef S_pu As Complex)
    Vi = CFromPolarRad(Vmag(i), Vang(i))
    Vj = CFromPolarRad(Vmag(J), Vang(J))
    Z = CCreate(R, X)
    ' I_ij = (Vi - Vj) / Z
    I_pu = CDiv(CSub(Vi, Vj), Z)
    ' S_ij = Vi * conj(I_ij)
    S_pu = CMul(Vi, CConj(I_pu))
End Sub

' Pomocná: Výpočet toku trafa (primár)
' A2: použitý CFromPolarRad – odpadá konverzia uhla radiány -> stupne -> radiány.
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

    ' I_prim (z i)
    Yseries_a2 = CCreate(Ys.Re / (Ratio * Ratio), Ys.Im / (Ratio * Ratio))
    term1 = CMul(Vi, CAdd(Yseries_a2, Ym))

    Yseries_a = CCreate(Ys.Re / Ratio, Ys.Im / Ratio)
    term2 = CMul(Vj, Yseries_a)

    I_prim = CSub(term1, term2)
    S_prim = CMul(Vi, CConj(I_prim))
End Sub
