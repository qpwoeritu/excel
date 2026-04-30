Attribute VB_Name = "modNR"
'==========================
' Modul: modNR
' Posledná úprava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

'--------------------------------------
' Vpoet innch a jalovch vkonov P, Q
'--------------------------------------
Private Sub CalcPower(ByVal nBuses As Long, _
                      ByRef G() As Double, _
                      ByRef B() As Double, _
                      ByRef Vmag() As Double, _
                      ByRef Vang() As Double, _
                      ByRef Pcalc() As Double, _
                      ByRef Qcalc() As Double)

    Dim i As Long, k As Long
    Dim theta As Double
    
    For i = 1 To nBuses
        Pcalc(i) = 0#
        Qcalc(i) = 0#
        For k = 1 To nBuses
            theta = Vang(i) - Vang(k)
            Pcalc(i) = Pcalc(i) + Vmag(i) * Vmag(k) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
            Qcalc(i) = Qcalc(i) + Vmag(i) * Vmag(k) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
        Next k
    Next i
End Sub

'--------------------------------------
' Zostavenie vektora nesladu ?P, ?Q (alebo ?V pre PV)
' ?P = Pspec - Pcalc
' ?Q = Qspec - Qcalc (pre PQ)
' ?V = Vspec - Vcalc (pre PV) - tu pouvam trik, e ?Q rovnica je nahraden rovnicou pre naptie
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
    
    ' prv polovica vektora  ?P (pre vetky neznme uzly: PQ aj PV)
    For i = 1 To nPQ
        idx = PQIndex(i)
        dP = Pspec(idx) - Pcalc(idx)
        mismatch(i) = dP
        If Abs(dP) > maxDP Then maxDP = Abs(dP)
    Next i
    
    ' druh polovica  ?Q (pre PQ) alebo ?V (pre PV)
    For i = 1 To nPQ
        idx = PQIndex(i)
        If BusTypes(idx) = btPQ Then
            ' PQ uzol: ?Q
            dQ = Qspec(idx) - Qcalc(idx)
            mismatch(nPQ + i) = dQ
            If Abs(dQ) > maxDQ Then maxDQ = Abs(dQ)
        ElseIf BusTypes(idx) = btPV Then
            ' Pre PV uzly (genertorov uzly) je naptie kontantn.
            ' Namiesto rovnice pre jalov vkon (Q) pouijeme podmienku fixnho naptia.
            ' Nastavenm mismatch na 0 a jednotkovho riadku v Jakobine (dV=0)
            ' zabezpeme, e vekos naptia ostane na poiatonej pecifikovanej hodnote.
            mismatch(nPQ + i) = 0#
        End If
    Next i
    
    epsilon = IIf(maxDP > maxDQ, maxDP, maxDQ)
End Sub

'--------------------------------------
' Zostavenie Jakobiho matice pre neznme uzly (PQ aj PV)
' J m rozmery (2*nPQ) x (2*nPQ)
' Pre PV uzly sa riadky M a L menia (M ostva ak rtame P, L sa nahrdza rovnicou pre dV)
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
    Dim theta As Double
    Dim H As Double, n As Double, m As Double, L As Double
    Dim Vi As Double
    
    For rowPQ = 1 To nPQ
        i = PQIndex(rowPQ)
        For colPQ = 1 To nPQ
            k = PQIndex(colPQ)
            
            ' Vpoet derivci (H, N, M, L) je rovnak pre vetky typy (zvis od fyziky siete)
            ' A pri zpise do J rozhodneme, i ich pouijeme
            
            If i = k Then
                Vi = Vmag(i)
                If Abs(Vi) < 0.000000001 Then Vi = 0.000000001
                
                H = -Qcalc(i) - B(i, i) * Vi * Vi
                n = Pcalc(i) / Vi + G(i, i) * Vi
                m = Pcalc(i) - G(i, i) * Vi * Vi
                L = Qcalc(i) / Vi - B(i, i) * Vi
            Else
                theta = Vang(i) - Vang(k)
                H = Vmag(i) * Vmag(k) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
                n = Vmag(i) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
                m = -Vmag(i) * Vmag(k) * (G(i, k) * Cos(theta) + B(i, k) * Sin(theta))
                L = Vmag(i) * (G(i, k) * Sin(theta) - B(i, k) * Cos(theta))
            End If
            
            ' H a N bloky (dP/dTheta, dP/dV) s platn pre PQ aj PV (lebo P je pecifikovan pre oba)
            J(rowPQ, colPQ) = H
            J(rowPQ, nPQ + colPQ) = n
            
            ' M a L bloky (dQ/dTheta, dQ/dV)
            If BusTypes(i) = btPQ Then
                ' Pre PQ uzol: Pouijeme tandardn M a L
                J(nPQ + rowPQ, colPQ) = m
                J(nPQ + rowPQ, nPQ + colPQ) = L
            ElseIf BusTypes(i) = btPV Then
                ' Spracovanie PV uzla v Jakobine:
                ' Rovnica pre odchlku jalovho vkonu je nahraden podmienkou kontantnho naptia (dV = 0).
                ' To dosiahne riadkom s nulami a jednotkou na diagonle v asti dP/dV.
                If i = k Then
                    J(nPQ + rowPQ, colPQ) = 0#        ' dV/dTheta = 0
                    J(nPQ + rowPQ, nPQ + colPQ) = 1#  ' dV/dV = 1
                Else
                    J(nPQ + rowPQ, colPQ) = 0#
                    J(nPQ + rowPQ, nPQ + colPQ) = 0#
                End If
            End If
        Next colPQ
    Next rowPQ
End Sub

'--------------------------------------
' Rieenie linerneho systmu J * x = rhs
' Rieenie J * x = rhs pomocou MINVERSE/MMULT
'--------------------------------------
' Rieenie linernej sstavy J * x = rhs pomocou Gaussovej elimincie
' Tto metda je vpotovo efektvnejia a stabilnejia ako inverzia matice.
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
    Dim A() As Double
    Dim B() As Double

    On Error GoTo ErrHandler

    n = UBound(J, 1)
    ReDim A(1 To n, 1 To n)
    ReDim B(1 To n)

    ' Kpia do loklnych pol, aby sme neovplyvnili pvodn maticu
    For i = 1 To n
        For m = 1 To n
            A(i, m) = J(i, m)
        Next m
        B(i) = rhs(i)
    Next i

    ' Priama elimincia s iastonm pivotovanm
    For i = 1 To n
        maxRow = i
        maxValue = Abs(A(i, i))
        For k = i + 1 To n
            If Abs(A(k, i)) > maxValue Then
                maxValue = Abs(A(k, i))
                maxRow = k
            End If
        Next k

        ' Zmena riadkov
        If maxRow <> i Then
            For k = i To n
                temp = A(i, k)
                A(i, k) = A(maxRow, k)
                A(maxRow, k) = temp
            Next k
            temp = B(i)
            B(i) = B(maxRow)
            B(maxRow) = temp
        End If

        ' Kontrola singularity
        If Abs(A(i, i)) < 0.000000000000001 Then
            Err.Raise vbObjectError + 10, "SolveLinearSystem_Gauss", "Matica je singulárna, sústava nemá riešenie."
        End If

        ' Elimincia pod pivotom
        For k = i + 1 To n
            factor = A(k, i) / A(i, i)
            B(k) = B(k) - factor * B(i)
            For m = i + 1 To n
                A(k, m) = A(k, m) - factor * A(i, m)
            Next m
        Next k
    Next i

    ' Sptn substitcia
    ReDim solution(1 To n)
    For i = n To 1 Step -1
        temp = B(i)
        For m = i + 1 To n
            temp = temp - A(i, m) * solution(m)
        Next m
        solution(i) = temp / A(i, i)
    Next i
    
    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 10, , "Chyba pri riešení sústavy: " & Err.Description
End Sub


'--------------------------------------
' Aktualizcia stavu napt
'--------------------------------------
Private Sub UpdateState(ByRef Vmag() As Double, _
                        ByRef Vang() As Double, _
                        ByRef PQIndex() As Long, _
                        ByVal nPQ As Long, _
                        ByRef deltaX() As Double)

    Dim i As Long, idx As Long
    
    ' prv polovica vektora  ??
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vang(idx) = Vang(idx) + deltaX(i)
    Next i
    
    ' druh polovica  ?|V|
    For i = 1 To nPQ
        idx = PQIndex(i)
        Vmag(idx) = Vmag(idx) + deltaX(nPQ + i)
    Next i
End Sub

'--------------------------------------
' Hlavn Newton-Raphson load-flow
'--------------------------------------
Public Sub NewtonRaphsonLoadFlow()
    Dim SBase_MVA As Double
    Dim BaseVoltages() As Double ' Nov bzy
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusBaseKV() As Double ' Nov pole bz pre uzly
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim BranchName() As String
    Dim R() As Double, X() As Double
    Dim Bshunt() As Double
    Dim BranchStatus() As Integer
    
    ' Transformtory
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
    
    ' Spnae
    Dim nSwitches As Long
    Dim SwitchName() As String
    Dim SwFrom() As Long, SwTo() As Long
    Dim SwR() As Double, SwX() As Double
    Dim SwStatus() As Integer
    
    ' Kompenzcia
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
    
    ' Topolgia - Izolovan asti
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
    Dim PQIndex() As Long ' Indexy uzlov, pre ktor rtame rovnice
    Dim nPQ As Long ' Poet rovnc (pre PQ uzly 2, pre PV uzly 1) - ALEBO poet aktvnych uzlov?
    ' Upresnenie: Newton-Raphson riei pre kad uzol (okrem Slack) rovnice.
    ' Pre PQ: P a Q. Pre PV: P.
    ' Moja implementcia predtm predpokladala len PQ uzly v zozname PQIndex a Slack.
    ' Teraz musme rozli PQ a PV.
    ' Pre jednoduchos: PQIndex bude obsahova vetky uzly okrem Slacku.
    ' A budeme dynamicky zostavova mismatch vektor a Jacobian.
    ' Ale pre zachovanie kompatibility s existujcimi funkciami (BuildMismatchVectors, BuildJacobian)
    ' musme by opatrn. Tie funkcie predpokladaj 2*nPQ vekos.
    
    ' Nov prstup:
    ' PQIndex bude zoznam VETKCH neznmych uzlov (PQ aj PV).
    ' nPQ bude poet tchto uzlov.
    ' Mismatch vektor bude ma vekos 2*nPQ? Nie.
    ' Pre PV uzol vynechme Q rovnicu?
    ' no. Ale to zmen truktru Jacobianu.
    
    ' Zjednoduenie pre PV uzol v tejto implementcii:
    ' PV uzol sa bude sprva ako PQ uzol, ale v kadej itercii resetujeme |V| na predpsan hodnotu?
    ' To je "Type Switching" metda (nie presne NR, ale funguje).
    ' Alebo Q-limit checking.
    ' Ak je to PV uzol, Q je neznma, |V| je znma.
    ' V NR formulcii: Neznme s dTheta a dV.
    ' Pre PV uzol je dV = 0.
    ' Teda stpec dV v Jacobiane pre PV uzol meme vynecha?
    ' A riadok dQ tie vynecha?
    
    ' Implementcia "Dummy Equation" pre PV uzol v plnej matici:
    ' Riadok dQ pre PV uzol nahradme rovnicou: dV = 0 (alebo V - Vspec = 0).
    ' V Jacobiane:
    ' Riadok zodpovedajci Q rovnici (nPQ + i) bude ma:
    ' dQ/dTheta = 0
    ' dQ/dV = 1 (alebo vek slo pre vyntenie, ale 1 sta ak rhs je dV)
    ' Prav strana (mismatch): Vspec - Vcalc.
    ' Tm pdom solver vypota dV tak, aby Vcalc + dV = Vspec.
    
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
    
    '--------------------------
    ' Natanie parametrov z listu index
    ' B3  max. poet iterci
    ' B4  epsilon limit
    '--------------------------
    With ThisWorkbook.Worksheets("index")
        maxIter = CLng(ParseDouble(.Range("B3").Value))
        If maxIter <= 0 Then maxIter = 20
        
        epsLimit = ParseDouble(.Range("B4").Value)
        If epsLimit <= 0 Then epsLimit = 0.000001
    End With
    
    '--------------------------
    ' Natanie uzlov a veden
    '--------------------------
    ' natanie bzovch hodnt
    Call GetBaseValues(SBase_MVA, BaseVoltages)
    
    ' natanie uzlov (skuton -> p.u.)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, BaseVoltages)
    
    ' natanie veden (ohm -> p.u.)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, Bshunt)
    
    ' natanie transformtorov (ohm/siemens -> p.u.)
    Call LoadTransformerData(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, BusNames, BusBaseKV, SBase_MVA)
    
    ' natanie reaktorov (ohm -> p.u.)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA)
    
    ' natanie dif. reaktorov
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA)
    
    ' natanie spnaov
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA)
    
    ' natanie kompenzcie
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA)
    
    ' natanie motorov
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA)
    
    '--------------------------
    ' Identifikcia izolovanch ast
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
                           
    ' Ak je izolovan uzol, vynuluj jeho Pspec a Qspec, aby nerobil problmy v mismatch vektore
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Pspec(i) = 0#
            Qspec(i) = 0#
            ' Vynulujeme aj poiaton naptie
            Vmag(i) = 0#
        End If
    Next i
    
    ' tvorba Y-matice v p.u. (s ignorovanm izolovanch)
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
    ' Identifikcia slack a PQ uzlov
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
        Err.Raise vbObjectError + 12, , "Nenjden slack uzol."
    End If
    ' Potame aj PV uzly do zoznamu neznmych
    If nPQ = 0 Then
        ' Err.Raise vbObjectError + 13, , "Nie s iadne PQ uzly." ' PV mu by
    End If
    
    ' Zrtame vetky non-Slack uzly
    nPQ = 0
    For i = 1 To nBuses
        If BusTypes(i) <> btSlack And Not IsBusIsolated(i) Then
            nPQ = nPQ + 1
        End If
    Next i
    
    If nPQ = 0 Then
         ' Len slack - skonen
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
    
    ' Aktualizcia nPQ (skuton poet potanch uzlov)
    nPQ = k
    ' Musme zmeni pole, ak sme nejak vynechali
    If nPQ > 0 Then ReDim Preserve PQIndex(1 To nPQ)
    
    ReDim Pcalc(1 To nBuses)
    ReDim Qcalc(1 To nBuses)
    If nPQ > 0 Then
        ReDim mismatch(1 To 2 * nPQ)
        ReDim J(1 To 2 * nPQ, 1 To 2 * nPQ)
    End If
    
    '--------------------------
    ' Prprava vsledkovch listov
    '--------------------------
    Call ClearResultsSheets
    
    startTime = Timer
    converged = False
    iterUsed = 0
    
    '--------------------------
    ' Hlavn NR iteran cyklus
    '--------------------------
    If nPQ > 0 Then
        For iter = 1 To maxIter
            iterUsed = iter
            
            ' vpoet P, Q
            Call CalcPower(nBuses, G, B, Vmag, Vang, Pcalc, Qcalc)
            
            ' vektor nesladu a epsilon (upraven pre PV)
            ' Pre PV uzol: Mismatch Q (druh polka) nahradme (Vspec - Vcalc)
            ' Alebo upravme BuildMismatchVectors
            Call BuildMismatchVectors(nBuses, BusTypes, Pspec, Qspec, Vmag, Pcalc, Qcalc, PQIndex, nPQ, mismatch, maxDP, maxDQ, eps, BusBaseKV) ' Pridane Vmag, BusBaseKV (pre Vspec ak treba, ale Vspec je asi konstanta)
            ' Pozor: Vspec pre PV uzol. Kde je uloen?
            ' Predpoklad: Vmag na zaiatku itercie pre PV uzol je nastaven na Vspec.
            ' A my chceme aby ostalo.
            
            ' logovanie napt a epsilon
            Call LogVoltages(iter, BusNames, Vmag, Vang)
            Call LogEpsilon(iter, maxDP, maxDQ, eps)
            
            ' kontrola konvergencie
            If eps < epsLimit Then
                converged = True
                Exit For
            End If
            
            ' Jakobiho matica
            Call BuildJacobian(nBuses, BusTypes, G, B, Vmag, Vang, Pcalc, Qcalc, PQIndex, nPQ, J)
            
            ' rieenie J * ?x = mismatch
            Call SolveLinearSystem_Gauss(J, mismatch, deltaX)
            
            ' aktualizcia napt
            Call UpdateState(Vmag, Vang, PQIndex, nPQ, deltaX)
        Next iter
    Else
        converged = True
    End If
    
SkipNR:
       totalTime = Timer - startTime
    Call WriteSummaryToIndex(totalTime, iterUsed, eps, converged)
    
    ' zap vsledn naptia na list "uzly" v skutonch hodnotch [kV]
    Call WriteFinalVoltagesToUzly(Vmag, Vang, BusBaseKV)
    
    ' vypotaj a zap prdy vo vedeniach v relnych hodnotch
    Call WriteBranchCurrents(nBranches, FromBus, ToBus, R, X, BranchStatus, Vmag, Vang, SBase_MVA, BusBaseKV, Bshunt)
    ' Vpoet prdov spnami
    If nSwitches > 0 Then
        ReDim SwCurrent_A(1 To nSwitches)
        For k = 1 To nSwitches
            If SwStatus(k) > 0 Then
                idx1 = SwFrom(k): idx2 = SwTo(k)
                Vi = CFromPolar(Vmag(idx1), Vang(idx1) * RAD2DEG)
                Vj = CFromPolar(Vmag(idx2), Vang(idx2) * RAD2DEG)
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

    
    ' vypotaj a zap toky v transformtoroch
    Call WriteTransformerFlows(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, Vmag, Vang, BusBaseKV, SBase_MVA)
    
    ' vypotaj a zap toky v reaktoroch
    Call WriteReactorResults(nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)
    
    ' vypotaj a zap toky v dif. reaktoroch
    Call WriteDifReactorResults(nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, Vmag, Vang, BusBaseKV, SBase_MVA)
    
    ' zap vsledky kompenzcie
    Call WriteCompResults(nComp, CompBus, Vmag, BusBaseKV)
    
    ' zap vsledky motorov
    Call WriteMotorResults(nMotors, MotorBus, MotorR, MotorG, MotorB, MotorStatus, Vmag, BusBaseKV, SBase_MVA)
    
    ' Vpoet a zpis celkovho zaaenia uzlov (P, Q, I)
    ' Posielame Pcalc a Qcalc (vsledn injekciu) namiesto Pspec/Qspec
    Call WriteNodeThroughput(nBuses, BusNames, BusBaseKV, SBase_MVA, _
                             nBranches, FromBus, ToBus, R, X, BranchStatus, _
                             nTrafo, TrFrom, TrTo, TrR, TrX, TrRatio, TrG, TrB, _
                             nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                             nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                             nComp, CompBus, CompB, CompStatus, _
                             nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                             Vmag, Vang, Pcalc, Qcalc)
    
    ' Report izolovanch (volan a na konci, aby prepsal stpec H na "izolovane")
    Call WriteIsolationReport(nBuses, BusNames, IsBusIsolated, _
                              nBranches, FromBus, ToBus, IsBranchIsolated, _
                              nTrafo, TrFrom, TrTo, IsTrafoIsolated, _
                              nComp, CompBus, IsCompIsolated)

    ' Aktualizcia SLD
    Call UpdateSLD

    Exit Sub

ErrHandler:
    MsgBox "Chyba vo výpočte Newton-Raphson: " & Err.Description, vbCritical
End Sub

'--------------------------------------
' Vpoet a zpis zaaenia uzlov (P, Q, I) do stpcov K, L, M
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
    
    ' Pomocn pre vpoty
    Dim Vi As Complex, Vj As Complex, Z As Complex, Ys As Complex
    Dim I_pu As Complex, S_pu As Complex
    Dim I_abs_pu As Double
    Dim P_flow As Double, Q_flow As Double
    Dim Ubase As Double, Ibase_A As Double
    
    ' 1. Vedenia
    For k = 1 To nBranches
        ' Len zapnut vedenia
        If BranchStatus(k) > 0 Then
            ' Uzol i -> j
            Call CalcBranchFlow(k, FromBus(k), ToBus(k), R(k), X(k), Vmag, Vang, Vi, Vj, Z, Ys, I_pu, S_pu)
            ' Tok z i do vedenia (S_pu)
            ' Ak P teie DO uzla i, S_pu.Re < 0 (lebo S_pu je tok i->j).
            ' Teda P_in = -S_pu.Re. Ak P_in > 0 -> zapota.
            P_flow = -S_pu.Re: Q_flow = -S_pu.Im
            I_abs_pu = CAbs(I_pu)
            
            If P_flow > 0 Then SumP(FromBus(k)) = SumP(FromBus(k)) + P_flow
            If Q_flow > 0 Then SumQ(FromBus(k)) = SumQ(FromBus(k)) + Q_flow
            ' I zapotame vdy (zaaenie zbernice pripojenou vetvou)
            SumI(FromBus(k)) = SumI(FromBus(k)) + I_abs_pu
            
            ' Uzol j -> i (opan tok)
            ' I_ji = -I_ij (pribline, ak zanedbme shunt, ale model vedenia tu nem shunt)
            ' S_ji = Vj * conj(-I_ij)
            Dim I_ji As Complex, S_ji As Complex
            I_ji = CCreate(-I_pu.Re, -I_pu.Im)
            S_ji = CMul(Vj, CConj(I_ji))
            
            P_flow = -S_ji.Re: Q_flow = -S_ji.Im ' Tok DO uzla j
            I_abs_pu = CAbs(I_ji)
            
            If P_flow > 0 Then SumP(ToBus(k)) = SumP(ToBus(k)) + P_flow
            If Q_flow > 0 Then SumQ(ToBus(k)) = SumQ(ToBus(k)) + Q_flow
            SumI(ToBus(k)) = SumI(ToBus(k)) + I_abs_pu
        End If
    Next k
    
    ' 2. Traf
    For k = 1 To nTrafo
        Call CalcTrafoFlow(k, TrFrom(k), TrTo(k), TrR(k), TrX(k), TrRatio(k), TrG(k), TrB(k), Vmag, Vang, _
                           Vi, Vj, I_pu, S_pu) ' Vrti I_prim a S_prim (tok z i do trafa)
        
        ' Primr (i)
        P_flow = -S_pu.Re: Q_flow = -S_pu.Im
        I_abs_pu = CAbs(I_pu)
        If P_flow > 0 Then SumP(TrFrom(k)) = SumP(TrFrom(k)) + P_flow
        If Q_flow > 0 Then SumQ(TrFrom(k)) = SumQ(TrFrom(k)) + Q_flow
        SumI(TrFrom(k)) = SumI(TrFrom(k)) + I_abs_pu
        
        ' Sekundr (j) - musme vypota tok na sekundri
        ' I_sec = (Vj * ys) - (Vi * ys/a) ... z modIO.WriteTransformerFlows
        ' Vypotame znova I_sec
        Dim Zs As Complex, ys_t As Complex, Yseries_a As Complex
        Zs = CCreate(TrR(k), TrX(k)): ys_t = CDiv(CCreate(1, 0), Zs)
        Yseries_a = CCreate(ys_t.Re / TrRatio(k), ys_t.Im / TrRatio(k))
        
        Dim term1 As Complex, term2 As Complex, I_sec As Complex, S_sec As Complex
        term1 = CMul(Vj, ys_t)
        term2 = CMul(Vi, Yseries_a)
        I_sec = CSub(term1, term2) ' Tok z j do trafa?
        ' Vzorec v WriteTransformerFlows bol: I_j = Vj*ys - Vi*(ys/a). Toto je prd teci z uzla j do siete trafa (ak sa nemlim v znamienkach).
        ' Pre istotu: I_sec = I_j. S_sec_inj = Vj * conj(I_sec). Toto je tok Z uzla j DO trafa.
        
        S_sec = CMul(Vj, CConj(I_sec))
        P_flow = -S_sec.Re: Q_flow = -S_sec.Im ' Tok DO uzla j (z trafa) = -(tok z j do trafa)
        ' Poka. S_sec (vypotan) je tok Z uzla J DO trafa.
        ' Take prtok do uzla J je -S_sec.
        ' P_in = - (S_sec.Re).
        
        I_abs_pu = CAbs(I_sec)
        If P_flow > 0 Then SumP(TrTo(k)) = SumP(TrTo(k)) + P_flow
        If Q_flow > 0 Then SumQ(TrTo(k)) = SumQ(TrTo(k)) + Q_flow
        SumI(TrTo(k)) = SumI(TrTo(k)) + I_abs_pu
    Next k
    
    ' 3. Reaktory
    For k = 1 To nReaktory
        ' Analogicky ako vedenie
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
    
    ' 5. Kompenzcia (Shunt)
    For k = 1 To nComp
        If CompStatus(k) = 1 Then
            ' I = V * Y = V * (jB)
            ' S = V * conj(I) = V * conj(V*jB) = |V|^2 * (-jB)
            ' P = 0, Q = -|V|^2 * B
            ' Ak B > 0 (kapacita), Q < 0 (dodva do siete).
            ' Prtok do uzla (zo strany kompenzcie):
            ' S_in = - S_comp = - (0 - j*|V|^2*B) = j*|V|^2*B.
            ' Q_in = |V|^2 * B.
            i = CompBus(k)
            Dim V_sq As Double
            V_sq = Vmag(i) * Vmag(i)
            
            P_flow = 0
            Q_flow = V_sq * CompB(k)
            
            ' Prd I = |V| * |B|
            I_abs_pu = Vmag(i) * Abs(CompB(k))
            
            If P_flow > 0 Then SumP(i) = SumP(i) + P_flow
            If Q_flow > 0 Then SumQ(i) = SumQ(i) + Q_flow
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next k
    
    ' 6. Motory (Shunt)
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            ' Y = G + jB
            ' S_motor = |V|^2 * conj(Y) = |V|^2 * (G - jB)
            ' S_in = -S_motor = |V|^2 * (-G + jB)
            ' P_in = -|V|^2 * G (G je zvyajne kladn = odber, take P_in < 0)
            ' Q_in = |V|^2 * B (B zvyajne zporn pre induktanciu -> Q_in < 0)
            
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
    
    ' 7. Injekcia do uzla (Genertory / Odbery)
    ' Pouvame Pcalc/Qcalc, o je skuton bilancia uzla po vpote (Vroba - Spotreba).
    ' Pre Slack uzol obsahuje Pcalc skuton dodvku.
    ' Pre PQ uzly obsahuje Pcalc = Pspec.
    
    For i = 1 To nBuses
        ' Pcalc > 0 znamen ist dodvka do siete (Genertor) -> Prtok P
        If Pcalc(i) > 0 Then
            SumP(i) = SumP(i) + Pcalc(i)
        End If
        
        ' Qcalc > 0 znamen ist dodvka Q do siete -> Prtok Q
        If Qcalc(i) > 0 Then
            SumQ(i) = SumQ(i) + Qcalc(i)
        End If
        
        ' Prd injekcie (i u genertor alebo odber)
        ' I_inj = |S_calc| / |V|
        ' Tento prd teie medzi uzlom a "okolm" (zemou/zdrojom). Je to prdov zaaenie prpojnice zo strany zdroja/zae.
        If Vmag(i) > 0.0000001 Then
            I_abs_pu = Sqr(Pcalc(i) * Pcalc(i) + Qcalc(i) * Qcalc(i)) / Vmag(i)
            SumI(i) = SumI(i) + I_abs_pu
        End If
    Next i
    
    ' Zpis do listu "uzly"
    Set ws = ThisWorkbook.Worksheets("uzly")
    ' Hlaviky
    ws.Cells(2, 11).Value = "Sum P_in [MW]"    ' K
    ws.Cells(2, 12).Value = "Sum Q_in [Mvar]"  ' L
    ws.Cells(2, 13).Value = "Sum I [A]"        ' M
    
    For i = 1 To nBuses
        ' Prepoet na relne jednotky
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

' Pomocn: Vpoet toku na zaiatku vetvy (I_ij, S_ij)
Private Sub CalcBranchFlow(ByVal k As Long, ByVal i As Long, ByVal J As Long, _
                           ByVal R As Double, ByVal X As Double, _
                           ByRef Vmag() As Double, ByRef Vang() As Double, _
                           ByRef Vi As Complex, ByRef Vj As Complex, _
                           ByRef Z As Complex, ByRef Ys As Complex, _
                           ByRef I_pu As Complex, ByRef S_pu As Complex)
    
    Vi = CFromPolar(Vmag(i), Vang(i) * RAD2DEG)
    Vj = CFromPolar(Vmag(J), Vang(J) * RAD2DEG)
    Z = CCreate(R, X)
    ' I_ij = (Vi - Vj) / Z
    I_pu = CDiv(CSub(Vi, Vj), Z)
    ' S_ij = Vi * conj(I_ij)
    S_pu = CMul(Vi, CConj(I_pu))
End Sub

' Pomocn: Vpoet toku trafa (primr)
Private Sub CalcTrafoFlow(ByVal k As Long, ByVal i As Long, ByVal J As Long, _
                          ByVal R As Double, ByVal X As Double, ByVal Ratio As Double, _
                          ByVal G As Double, ByVal B As Double, _
                          ByRef Vmag() As Double, ByRef Vang() As Double, _
                          ByRef Vi As Complex, ByRef Vj As Complex, _
                          ByRef I_prim As Complex, ByRef S_prim As Complex)
    
    Dim Zs As Complex, Ys As Complex, Ym As Complex
    Dim term1 As Complex, term2 As Complex
    Dim Yseries_a2 As Complex, Yseries_a As Complex
    
    Vi = CFromPolar(Vmag(i), Vang(i) * RAD2DEG)
    Vj = CFromPolar(Vmag(J), Vang(J) * RAD2DEG)
    
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



