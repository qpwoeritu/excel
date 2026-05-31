Attribute VB_Name = "modShortCircuit"
'==========================
' Modul: modShortCircuit
' Skratové výpočty (Ik3''). Rozdelené na dve fázy:
'   1) BuildShortCircuitMatrix – zostavenie admitančnej matice Ysc
'   2) SolveShortCircuit       – inverzia Ysc a výpočet Ik vo všetkých uzloch
' runCALC ich volá oddelene, aby vedel časovať a zobrazovať každú fázu zvlášť.
'==========================
Option Explicit

'--------------------------------------
' Zostavenie skratovej admitančnej matice Ysc.
' Výstup: Ysc(1..nBuses, 1..nBuses)
'--------------------------------------
Public Sub BuildShortCircuitMatrix( _
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
    ByVal nGens As Long, ByRef GenTermBus As Variant, ByRef GenStatus As Variant, ByRef GenRa As Variant, ByRef GenXd As Variant, _
    ByRef BusNames As Variant, ByRef BusTypes As Variant, ByRef BusBaseKV As Variant, _
    ByRef Ik_input As Variant, ByVal SBase_MVA As Double, _
    ByVal scenarioMin As Boolean, ByRef TrKt As Variant, ByRef GenKg As Variant, _
    ByRef IsBusIsolated As Variant, ByRef IsBranchIsolated As Variant, ByRef IsTrafoIsolated As Variant, _
    ByRef IsReaktorIsolated As Variant, ByRef IsDifReaktorIsolated As Variant, ByRef IsSwitchIsolated As Variant, _
    ByRef Ysc() As Complex)

    Dim i As Long, J As Long, k As Long
    Dim Z As Complex, Ys As Complex, t1 As Complex, t2 As Complex, A As Double
    Dim Ik_slack As Double, Z_grid_abs As Double, Un As Double
    Dim kt As Double, kg As Double

    ReDim Ysc(1 To nBuses, 1 To nBuses)

    ' Inicializácia Ysc (izolované uzly = 1.0 na diagonále)
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

    ' Trafá (impedancia korigovaná činiteľom K_T podľa IEC 60909-0; K_T=1 ak chýbajú menovité dáta)
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            kt = CDbl(TrKt(k))
            If kt <= 0# Then kt = 1#
            Z = CCreate(CDbl(TrR(k)) * kt, CDbl(TrX(k)) * kt): Ys = CDiv(CCreate(1, 0), Z)
            i = TrFrom(k): J = TrTo(k): A = TrRatio(k)
            t1 = CCreate(Ys.Re / (A * A), Ys.Im / (A * A))
            Ysc(i, i) = CAdd(Ysc(i, i), t1): Ysc(J, J) = CAdd(Ysc(J, J), Ys)
            t2 = CCreate(Ys.Re / A, Ys.Im / A)
            Ysc(i, J) = CSub(Ysc(i, J), t2): Ysc(J, i) = CSub(Ysc(J, i), t2)
        End If
    Next k

    ' Spínače
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

    ' Motory VN – príspevok do skratu. Model: Z = j*Xk -> Ys = -j*(1/Xk).
    For k = 1 To nMotors
        If MotorStatus(k) = 1 Then
            If Abs(CDbl(MotorXk(k))) > 0.0000001 Then
                Ys = CCreate(0, -1# / CDbl(MotorXk(k)))
                i = MotorBus(k)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys)
            End If
        End If
    Next k

    ' Generátory – subtranzientná impedancia Z = K_G*(R_G + j*X''d) podľa IEC 60909-0
    ' (K_G=1 ak chýbajú menovité dáta pre korekciu).
    For k = 1 To nGens
        If GenStatus(k) = 1 Then
            i = GenTermBus(k)
            If Not IsBusIsolated(i) And Not (GenRa(k) = 0 And GenXd(k) = 0) Then
                kg = CDbl(GenKg(k))
                If kg <= 0# Then kg = 1#
                Z = CCreate(CDbl(GenRa(k)) * kg, CDbl(GenXd(k)) * kg)
                Ys = CDiv(CCreate(1, 0), Z)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys)
            End If
        End If
    Next k

    ' Slack impedancia (zo zadaného Ik3'' v liste uzly), napäťový činiteľ c podľa scenára.
    For i = 1 To nBuses
        If BusTypes(i) = 0 Then ' btSlack
            Ik_slack = Ik_input(i)
            If Ik_slack > 0 Then
                Un = BusBaseKV(i)
                Z_grid_abs = (SCVoltageFactor(Un, scenarioMin) * SBase_MVA) / (Sqr(3) * Un * Ik_slack)
                Ysc(i, i) = CAdd(Ysc(i, i), CDiv(CCreate(1, 0), CCreate(0, Z_grid_abs)))
            End If
            Exit For
        End If
    Next i

    ' Diagnostický výpis skratovej matice
    Call WriteSCMatrix(Ysc, BusNames)
End Sub

'--------------------------------------
' Inverzia Ysc + výpočet Ik vo všetkých uzloch.
' Vstup: Ysc (z BuildShortCircuitMatrix)
' Výstup: Ik_result(1..nBuses) v kA
'--------------------------------------
Public Sub SolveShortCircuit( _
    ByRef Ysc() As Complex, _
    ByVal nBuses As Long, _
    ByRef BusBaseKV As Variant, _
    ByVal SBase_MVA As Double, _
    ByVal scenarioMin As Boolean, _
    ByRef IsBusIsolated As Variant, _
    ByRef Ik_result As Variant, _
    ByRef Rth As Variant, _
    ByRef Xth As Variant)

    Dim i As Long
    Dim Z_inv() As Complex
    Dim R_th As Double, X_th As Double, Z_th As Double, Un As Double, cF As Double

    ReDim Ik_result(1 To nBuses)
    ReDim Rth(1 To nBuses)
    ReDim Xth(1 To nBuses)

    ' Natívna komplexná inverzia n×n (bez 2n×2n reálneho rozšírenia)
    ' – objem aritmetiky klesá zhruba 8-krát oproti pôvodnému prístupu.
    Call ComplexMatrixInverse_Gauss(Ysc, Z_inv)

    ' Výpočet Ik v každom uzle z diagonály Z_th = Ysc^-1.
    ' R_th, X_th sa exportujú pre nadväzný výpočet κ, i_p a I_th (ComputeSCDerived).
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Ik_result(i) = 0
            Rth(i) = 0#
            Xth(i) = 0#
        Else
            R_th = Z_inv(i, i).Re
            X_th = Z_inv(i, i).Im
            Rth(i) = R_th
            Xth(i) = X_th
            Z_th = Sqr(R_th * R_th + X_th * X_th)
            Un = BusBaseKV(i)
            If Z_th > 0.0000001 Then
                cF = SCVoltageFactor(Un, scenarioMin)
                Ik_result(i) = (cF / Z_th) * (SBase_MVA / (Sqr(3) * Un))
            Else
                Ik_result(i) = 0
            End If
        End If
    Next i
End Sub

Public Sub WriteShortCircuitResults(ByRef Ik_result As Variant, ByVal nBuses As Long, _
    ByRef ip_result As Variant, ByRef Ith_result As Variant, _
    ByRef Ib_result As Variant, ByRef Ik_steady As Variant, _
    ByRef kappa_result As Variant, ByRef RXratio_result As Variant)
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("uzly")
    ' Stĺpec J (10) = Ik3'' (vstup pre slack aj výstup); nové veličiny v K..P (11..16).
    ws.Cells(2, 10).Value = "Ik3'' [kA]"
    ws.Cells(2, 11).Value = "ip [kA]"
    ws.Cells(2, 12).Value = "Ith [kA]"
    ws.Cells(2, 13).Value = "Ib [kA]"
    ws.Cells(2, 14).Value = "Ik [kA]"
    ws.Cells(2, 15).Value = "kappa [-]"
    ws.Cells(2, 16).Value = "R/X [-]"
    For i = 1 To nBuses
        ws.Cells(2 + i, 10).Value = Round(Ik_result(i), 2)
        ws.Cells(2 + i, 11).Value = Round(ip_result(i), 2)
        ws.Cells(2 + i, 12).Value = Round(Ith_result(i), 2)
        ws.Cells(2 + i, 13).Value = Round(Ib_result(i), 2)
        ws.Cells(2 + i, 14).Value = Round(Ik_steady(i), 2)
        ws.Cells(2 + i, 15).Value = Round(kappa_result(i), 3)
        ws.Cells(2 + i, 16).Value = Round(RXratio_result(i), 3)
    Next i
End Sub

' Zápis skratovej admitančnej matice pre kontrolu na list SC_matica
Private Sub WriteSCMatrix(ByRef Ysc() As Complex, ByRef BusNames As Variant)
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long, J As Long
    Dim row0 As Long, col0 As Long

    Set ws = GetOrCreateSheet("SC_matica")
    ws.Cells.Clear

    n = UBound(Ysc, 1)

    row0 = 1
    col0 = 1

    Dim arr As Variant
    Dim startRowX As Long

    ' Blok Re(Ysc) – hlavička + matica v jednom Variant poli, jeden Range.Value zápis
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
' Inverzia komplexnej matice Gauss-Jordanovou elimináciou s čiastočným
' pivotovaním. Pracuje natívne nad UDT Complex – nepoužíva 2n×2n reálne
' rozšírenie, čím sa objem aritmetiky zníži zhruba 8-krát.
'
' Pre rýchlosť je komplexná aritmetika v hot-loopoch (normalizácia pivotného
' riadka a eliminácia) inlinovaná – ušetrí sa volanie/kópia UDT cez CMul/CDiv.
'
' Volá PhaseYield (z modProgress) raz za pivotný riadok, aby sa aktualizovala
' časová bunka (J7) a Excel ostal responzívny počas dlhej inverzie.
'--------------------------------------
Private Sub ComplexMatrixInverse_Gauss(ByRef A_in() As Complex, ByRef A_inv() As Complex)
    Dim n As Long, i As Long, J As Long, k As Long
    Dim maxRow As Long, maxMag2 As Double, mag2 As Double
    Dim tempC As Complex
    Dim pivotInvRe As Double, pivotInvIm As Double
    Dim fRe As Double, fIm As Double
    Dim aRe As Double, aIm As Double
    Dim A() As Complex

    n = UBound(A_in, 1)
    ReDim A(1 To n, 1 To 2 * n)

    ' Príprava rozšírenej matice (A | I) – ľavá polovica je vstup, pravá identita
    For i = 1 To n
        For J = 1 To n
            A(i, J) = A_in(i, J)
        Next J
        A(i, n + i).Re = 1#
        ' Imaginárna časť ostáva 0 z inicializácie ReDim
    Next i

    ' Gauss-Jordanova eliminácia
    For i = 1 To n
        ' Heartbeat: max raz za ~200 ms aktualizuje časovú bunku J7 a urobí DoEvents
        Call PhaseYield

        ' Pivotovanie: riadok s najväčším |Z|^2 v stĺpci i.
        ' Stačí |Z|^2 (ušetríme Sqr), na poradie pivotov to nemá vplyv.
        maxRow = i
        maxMag2 = A(i, i).Re * A(i, i).Re + A(i, i).Im * A(i, i).Im
        For k = i + 1 To n
            mag2 = A(k, i).Re * A(k, i).Re + A(k, i).Im * A(k, i).Im
            If mag2 > maxMag2 Then
                maxMag2 = mag2
                maxRow = k
            End If
        Next k

        ' Výmena riadkov. Stĺpce 1..i-1 sú už nulové z predošlých eliminácií,
        ' takže výmenu začíname od stĺpca i.
        If maxRow <> i Then
            For k = i To 2 * n
                tempC = A(i, k)
                A(i, k) = A(maxRow, k)
                A(maxRow, k) = tempC
            Next k
        End If

        ' Test singularity (|Z|^2 < 1e-36 zodpovedá |Z| < 1e-18 ako v pôvodnej verzii)
        If maxMag2 < 1E-36 Then
            Err.Raise vbObjectError + 102, , "Skratová matica je singulárna."
        End If

        ' Predpočítaná inverzia pivotu: 1/p = conj(p) / |p|^2
        pivotInvRe = A(i, i).Re / maxMag2
        pivotInvIm = -A(i, i).Im / maxMag2

        ' Normalizácia pivotného riadka: A(i, :) *= 1/pivot.
        ' Stĺpec i nastavíme na presnú jednotku (predíde sa zaokrúhľovacej chybe).
        A(i, i).Re = 1#
        A(i, i).Im = 0#
        For k = i + 1 To 2 * n
            aRe = A(i, k).Re
            aIm = A(i, k).Im
            A(i, k).Re = aRe * pivotInvRe - aIm * pivotInvIm
            A(i, k).Im = aRe * pivotInvIm + aIm * pivotInvRe
        Next k

        ' Eliminácia ostatných riadkov: A(k, :) -= A(k, i) * A(i, :)
        For k = 1 To n
            If k <> i Then
                fRe = A(k, i).Re
                fIm = A(k, i).Im
                If fRe <> 0# Or fIm <> 0# Then
                    ' Stĺpec i v eliminovanom riadku bude presne 0
                    A(k, i).Re = 0#
                    A(k, i).Im = 0#
                    For J = i + 1 To 2 * n
                        ' (fRe + i*fIm) * (A(i,J).Re + i*A(i,J).Im)
                        '   = (fRe*A.Re - fIm*A.Im) + i*(fRe*A.Im + fIm*A.Re)
                        A(k, J).Re = A(k, J).Re - (fRe * A(i, J).Re - fIm * A(i, J).Im)
                        A(k, J).Im = A(k, J).Im - (fRe * A(i, J).Im + fIm * A(i, J).Re)
                    Next J
                End If
            End If
        Next k
    Next i

    ' Extrakcia inverznej matice z pravej polovice rozšírenej matice
    ReDim A_inv(1 To n, 1 To n)
    For i = 1 To n
        For J = 1 To n
            A_inv(i, J) = A(i, n + J)
        Next J
    Next i
End Sub

'--------------------------------------
' Napäťový činiteľ c podľa IEC 60909-0 (tab. 1).
'   scenario "max" (default): c = 1.1 pre všetky hladiny – zhodné s pôvodným správaním.
'   scenario "min":           c = 1.0 pre U_n > 1 kV, c = 0.95 pre NN.
'--------------------------------------
Public Function SCVoltageFactor(ByVal Un_kV As Double, ByVal scenarioMin As Boolean) As Double
    If scenarioMin Then
        If Un_kV > 1# Then
            SCVoltageFactor = 1#
        Else
            SCVoltageFactor = 0.95
        End If
    Else
        SCVoltageFactor = 1.1
    End If
End Function

'--------------------------------------
' Odvodené skratové veličiny podľa IEC 60909-0 z Ik'' a Théveninovej impedancie:
'   - nárazový činiteľ κ = 1.02 + 0.98·e^(-3·R/X)   (metóda B – R/X v mieste skratu)
'   - nárazový prúd        i_p   = κ·√2·Ik''
'   - tepelný prúd         I_th  = Ik''·√(m + n),  n = 1 (vzdialený skrat),
'        m = 1/(2·f·T_k·ln(κ-1)) · (e^(4·f·T_k·ln(κ-1)) - 1)
'   - vypínací/ustálený    I_b = I_k = Ik''         (predpoklad vzdialeného skratu)
'
' Poznámka: pri skrate v blízkosti generátora/motora útlm (μ, λ, q) pre I_b/I_k
' a teplotná korekcia odporu pre min. prúdy nie sú zahrnuté (vyžadujú ďalšie
' menovité dáta strojov) – I_b/I_k tu zodpovedajú vzdialenému skratu.
'--------------------------------------
Public Sub ComputeSCDerived( _
    ByVal nBuses As Long, _
    ByRef Ik_result As Variant, ByRef Rth As Variant, ByRef Xth As Variant, _
    ByRef IsBusIsolated As Variant, _
    ByVal Tk_s As Double, ByVal f_Hz As Double, _
    ByRef ip_result As Variant, ByRef Ith_result As Variant, _
    ByRef Ib_result As Variant, ByRef Ik_steady As Variant, _
    ByRef kappa_result As Variant, ByRef RXratio_result As Variant)

    Dim i As Long
    Dim rx As Double, kappa As Double, m As Double, n As Double, lnk As Double

    ReDim ip_result(1 To nBuses)
    ReDim Ith_result(1 To nBuses)
    ReDim Ib_result(1 To nBuses)
    ReDim Ik_steady(1 To nBuses)
    ReDim kappa_result(1 To nBuses)
    ReDim RXratio_result(1 To nBuses)

    For i = 1 To nBuses
        If IsBusIsolated(i) Or Ik_result(i) <= 0# Then
            ip_result(i) = 0#: Ith_result(i) = 0#
            Ib_result(i) = 0#: Ik_steady(i) = 0#
            kappa_result(i) = 0#: RXratio_result(i) = 0#
        Else
            ' R/X v mieste skratu z Théveninovej impedancie (guard X_th -> 0)
            If Abs(Xth(i)) > 0.0000001 Then
                rx = Rth(i) / Xth(i)
            Else
                rx = 0#
            End If
            If rx < 0# Then rx = 0#
            RXratio_result(i) = rx

            kappa = 1.02 + 0.98 * Exp(-3# * rx)
            kappa_result(i) = kappa
            ip_result(i) = kappa * Sqr(2#) * Ik_result(i)

            ' Tepelný ekvivalentný prúd: jednosmerná zložka m, striedavá n = 1.
            n = 1#
            m = 0#
            If Tk_s > 0# And f_Hz > 0# And (kappa - 1#) > 0# Then
                lnk = Log(kappa - 1#)              ' VBA Log = prirodzený logaritmus
                If Abs(lnk) < 0.000000001 Then
                    m = 2#                          ' limita pre κ -> 2 (R/X -> 0): m -> 2
                Else
                    m = (1# / (2# * f_Hz * Tk_s * lnk)) * (Exp(4# * f_Hz * Tk_s * lnk) - 1#)
                    If m < 0# Then m = 0#
                End If
            End If
            Ith_result(i) = Ik_result(i) * Sqr(m + n)

            ' Vzdialený skrat: vypínací aj ustálený prúd = začiatočný Ik''.
            Ib_result(i) = Ik_result(i)
            Ik_steady(i) = Ik_result(i)
        End If
    Next i
End Sub
