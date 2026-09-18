Attribute VB_Name = "modShortCircuit"
'==========================
' Modul: modShortCircuit
' Skratové výpočty podľa IEC 60909 (Ik3'', ip, vetvové príspevky).
' Rozdelené na fázy:
'   1) BuildShortCircuitMatrix – zostavenie admitančnej matice Ysc
'      (korekcie K_T pre trafá v max prípade, K_G pre generátory, R/X napájača
'       a motorov; motory sa pri min prípade zanedbávajú)
'   2) SolveShortCircuit       – inverzia Ysc, Ik'' a ip vo všetkých uzloch
'      (c-faktor podľa hladiny a prípadu, kappa metódou B pre zauzlené siete)
'   3) ComputeBranchContributions – prúdy vetiev pri skrate vo zvolenom uzle
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
    ByRef TrR As Variant, ByRef TrX As Variant, ByRef TrRatio As Variant, ByRef TrKT As Variant, _
    ByVal nReaktory As Long, ByRef ReaktorFrom As Variant, ByRef ReaktorTo As Variant, _
    ByRef ReaktorR As Variant, ByRef ReaktorX As Variant, _
    ByVal nDifReaktory As Long, ByRef DifReaktorFrom As Variant, ByRef DifReaktorTo As Variant, _
    ByRef DifReaktorR As Variant, ByRef DifReaktorX As Variant, _
    ByVal nMotors As Long, ByRef MotorBus As Variant, ByRef MotorR As Variant, ByRef MotorXk As Variant, ByRef MotorStatus As Variant, _
    ByVal nGens As Long, ByRef GenTermBus As Variant, ByRef GenStatus As Variant, ByRef GenRa As Variant, ByRef GenXd As Variant, ByRef GenKG As Variant, _
    ByRef BusNames As Variant, ByRef BusTypes As Variant, ByRef BusBaseKV As Variant, _
    ByRef Ik_input As Variant, ByVal SBase_MVA As Double, _
    ByVal caseMax As Boolean, ByVal RXfeeder As Double, _
    ByRef IsBusIsolated As Variant, ByRef IsBranchIsolated As Variant, ByRef IsTrafoIsolated As Variant, _
    ByRef IsReaktorIsolated As Variant, ByRef IsDifReaktorIsolated As Variant, ByRef IsSwitchIsolated As Variant, _
    ByRef Ysc() As Complex)

    Dim i As Long, J As Long, k As Long
    Dim Z As Complex, Ys As Complex, t1 As Complex, t2 As Complex, A As Double
    Dim Ik_slack As Double, Z_grid_abs As Double, Un As Double
    Dim kt As Double, cQ As Double, Xq As Double, Rq As Double
    Dim Zbase As Double, Zm_pu As Double, Rm_pu As Double, Xm_pu As Double
    Dim Ra_g As Double, Xd_g As Double

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

    ' Trafá – v max prípade s korekciou impedancie K_T (IEC 60909-0, čl. 6.3.3);
    ' pri min prípade sa K_T podľa normy neaplikuje (K_T = 1).
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            If caseMax Then kt = CDbl(TrKT(k)) Else kt = 1#
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

    ' Motory VN – príspevok len pri maximálnych prúdoch (IEC 60909: pri výpočte
    ' minimálnych skratových prúdov sa príspevky motorov zanedbávajú).
    ' Model: |Z_M| zo stĺpca P (skratová impedancia), R_M zo stĺpca L [ohm];
    ' ak R_M chýba, použije sa pomer R_M/X_M = 0,1 (VN motory, IEC 60909-0 čl. 6.5).
    If caseMax Then
        For k = 1 To nMotors
            If MotorStatus(k) = 1 Then
                Zm_pu = Abs(CDbl(MotorXk(k)))
                If Zm_pu > 0.0000001 Then
                    i = MotorBus(k)
                    Un = BusBaseKV(i)
                    If SBase_MVA <> 0# And Un <> 0# Then
                        Zbase = (Un * Un) / SBase_MVA
                    Else
                        Zbase = 1#
                    End If
                    Rm_pu = CDbl(MotorR(k)) / Zbase
                    If Rm_pu > 0# And Rm_pu < Zm_pu Then
                        Xm_pu = Sqr(Zm_pu * Zm_pu - Rm_pu * Rm_pu)
                    Else
                        Xm_pu = Zm_pu / Sqr(1.01)   ' R/X = 0,1
                        Rm_pu = 0.1 * Xm_pu
                    End If
                    Ys = CDiv(CCreate(1, 0), CCreate(Rm_pu, Xm_pu))
                    Ysc(i, i) = CAdd(Ysc(i, i), Ys)
                End If
            End If
        Next k
    End If

    ' Generátory – príspevok cez korigovanú subtranzientnú impedanciu
    ' Z = K_G·(R_G + j·X''d)  (IEC 60909-0, čl. 6.6.2).
    ' Ak R_G nie je zadaný, použije sa fiktívny odpor R_Gf = 0,07·X''d
    ' (VN generátory so S_rG < 100 MVA – potrebný pre korektný výpočet ip).
    For k = 1 To nGens
        If GenStatus(k) = 1 Then
            i = GenTermBus(k)
            If Not IsBusIsolated(i) And Not (GenRa(k) = 0 And GenXd(k) = 0) Then
                Ra_g = CDbl(GenRa(k))
                Xd_g = CDbl(GenXd(k))
                If Ra_g = 0# Then Ra_g = 0.07 * Xd_g
                Z = CCreate(CDbl(GenKG(k)) * Ra_g, CDbl(GenKG(k)) * Xd_g)
                Ys = CDiv(CCreate(1, 0), Z)
                Ysc(i, i) = CAdd(Ysc(i, i), Ys)
            End If
        End If
    Next k

    ' Slack impedancia (z dodaného Ik'' napájača):
    '   Z_Q = c(Un, prípad)·U_n/(sqrt(3)·Ik'') prepočítané do p.u.,
    '   rozklad X_Q = Z_Q/sqrt(1+(R/X)²), R_Q = (R/X)·X_Q  (IEC odporúča R/X = 0,1)
    For i = 1 To nBuses
        If BusTypes(i) = 0 Then ' btSlack
            Ik_slack = Ik_input(i)
            If Ik_slack > 0 Then
                Un = BusBaseKV(i)
                cQ = GetVoltageFactorC(Un, caseMax)
                Z_grid_abs = (cQ * SBase_MVA) / (Sqr(3) * Un * Ik_slack)
                Xq = Z_grid_abs / Sqr(1# + RXfeeder * RXfeeder)
                Rq = RXfeeder * Xq
                Ysc(i, i) = CAdd(Ysc(i, i), CDiv(CCreate(1, 0), CCreate(Rq, Xq)))
            End If
            Exit For
        End If
    Next i

    ' Diagnostický výpis skratovej matice
    Call WriteSCMatrix(Ysc, BusNames)
End Sub

'--------------------------------------
' Inverzia Ysc + výpočet Ik'' a ip vo všetkých uzloch.
' Vstup:  Ysc (z BuildShortCircuitMatrix), caseMax (prípad max/min)
' Výstup: Ik_result(1..nBuses) v kA, ip_result(1..nBuses) v kA,
'         Z_inv – celá impedančná matica (pre vetvové príspevky)
'--------------------------------------
Public Sub SolveShortCircuit( _
    ByRef Ysc() As Complex, _
    ByVal nBuses As Long, _
    ByRef BusNames As Variant, _
    ByRef BusBaseKV As Variant, _
    ByVal SBase_MVA As Double, _
    ByVal caseMax As Boolean, _
    ByRef IsBusIsolated As Variant, _
    ByRef Ik_result As Variant, _
    ByRef ip_result As Variant, _
    ByRef Z_inv() As Complex)

    Dim i As Long
    Dim R_th As Double, X_th As Double, Z_th As Double, Un As Double
    Dim c As Double, kappa As Double, kmax As Double

    ReDim Ik_result(1 To nBuses)
    ReDim ip_result(1 To nBuses)

    ' Natívna komplexná inverzia n×n (bez 2n×2n reálneho rozšírenia)
    ' – objem aritmetiky klesá zhruba 8-krát oproti pôvodnému prístupu.
    Call ComplexMatrixInverse_Gauss(Ysc, Z_inv)

    ' Výpočet Ik'' a ip v každom uzle z diagonály Z_th = Ysc^-1
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            Ik_result(i) = 0
            ip_result(i) = 0
        Else
            R_th = Z_inv(i, i).Re
            X_th = Z_inv(i, i).Im
            Z_th = Sqr(R_th * R_th + X_th * X_th)
            Un = BusBaseKV(i)
            c = GetVoltageFactorC(Un, caseMax)
            If Z_th > 0.0000001 Then
                Ik_result(i) = (c / Z_th) * (SBase_MVA / (Sqr(3) * Un))

                ' Nárazový prúd ip = kappa·sqrt(2)·Ik'' – metóda B pre zauzlené
                ' siete (IEC 60909-0, čl. 8): kappa_b z pomeru R/X v mieste skratu,
                ' korekcia 1,15, strop 2,0 (VN) resp. 1,8 (NN), minimum 1,0.
                If X_th > 0.0000001 Then
                    kappa = 1.02 + 0.98 * Exp(-3# * R_th / X_th)
                Else
                    kappa = 2#      ' nefyzikálna Théveninova impedancia - konzervatívne maximum
                    Call AddCalcWarning("Uzol '" & CStr(BusNames(i)) & "': Théveninova reaktancia X_th <= 0, " & _
                        "kappa nastavené konzervatívne na 2,0.")
                End If
                kappa = 1.15 * kappa
                If Un > 1# Then kmax = 2# Else kmax = 1.8
                If kappa > kmax Then kappa = kmax
                If kappa < 1# Then kappa = 1#
                ip_result(i) = kappa * Sqr(2#) * Ik_result(i)
            Else
                Ik_result(i) = 0
                ip_result(i) = 0
            End If
        End If
    Next i
End Sub

' Ik'' do uzly!J, ip do uzly!N (hlavičku N2 si dopĺňa používateľ, kód ju nezapisuje)
Public Sub WriteShortCircuitResults(ByRef Ik_result As Variant, ByRef ip_result As Variant, ByVal nBuses As Long)
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("uzly")
    For i = 1 To nBuses
        ws.Cells(2 + i, 10).Value = Round(Ik_result(i), 2)
        ws.Cells(2 + i, 14).Value = Round(ip_result(i), 2)
    Next i
End Sub

'--------------------------------------
' Vetvové príspevky pri skrate vo zvolenom uzle f (metóda ekvivalentného
' napäťového zdroja, IEC 60909):
'   I_f = c_f / Z_ff          (p.u.)
'   V_i = c_f - Z_if · I_f    (napätia počas skratu; predporuchovo naprázdno)
'   vetvy: I = (V_i - V_j)·y_s (bez B/2), trafo cez ys/a², ys/a,
'   motory I = V_i·y_M (len max prípad), generátory I = V_t·y_G.
' Výsledky sa zapíšu do hárku "skrat_vetvy", zoradené zostupne podľa prúdu.
'--------------------------------------
Public Sub ComputeBranchContributions( _
    ByVal faultBus As Long, ByVal nBuses As Long, ByRef Z_inv() As Complex, _
    ByVal caseMax As Boolean, ByVal SBase_MVA As Double, _
    ByRef BusNames As Variant, ByRef BusBaseKV As Variant, ByRef IsBusIsolated As Variant, _
    ByVal nBranches As Long, ByRef BranchName As Variant, ByRef FromBus As Variant, ByRef ToBus As Variant, _
    ByRef R As Variant, ByRef X As Variant, ByRef BranchStatus As Variant, ByRef IsBranchIsolated As Variant, _
    ByVal nSwitches As Long, ByRef SwitchName As Variant, ByRef SwFrom As Variant, ByRef SwTo As Variant, _
    ByRef SwR As Variant, ByRef SwX As Variant, ByRef SwStatus As Variant, ByRef IsSwitchIsolated As Variant, _
    ByVal nTrafo As Long, ByRef TrName As Variant, ByRef TrFrom As Variant, ByRef TrTo As Variant, _
    ByRef TrR As Variant, ByRef TrX As Variant, ByRef TrRatio As Variant, ByRef TrKT As Variant, ByRef IsTrafoIsolated As Variant, _
    ByVal nReaktory As Long, ByRef ReaktorName As Variant, ByRef ReaktorFrom As Variant, ByRef ReaktorTo As Variant, _
    ByRef ReaktorR As Variant, ByRef ReaktorX As Variant, ByRef IsReaktorIsolated As Variant, _
    ByVal nDifReaktory As Long, ByRef DifReaktorName As Variant, ByRef DifReaktorFrom As Variant, ByRef DifReaktorTo As Variant, _
    ByRef DifReaktorR As Variant, ByRef DifReaktorX As Variant, ByRef IsDifReaktorIsolated As Variant, _
    ByVal nMotors As Long, ByRef MotorName As Variant, ByRef MotorBus As Variant, ByRef MotorR As Variant, _
    ByRef MotorXk As Variant, ByRef MotorStatus As Variant, _
    ByVal nGens As Long, ByRef GenName As Variant, ByRef GenTermBus As Variant, ByRef GenStatus As Variant, _
    ByRef GenRa As Variant, ByRef GenXd As Variant, ByRef GenKG As Variant, _
    ByRef Ik_result As Variant, ByRef ip_result As Variant)

    Dim i As Long, J As Long, k As Long
    Dim cF As Double, Un As Double, Zbase As Double
    Dim Zff As Complex, If_pu As Complex
    Dim Vbus() As Complex
    Dim Z As Complex, Ys As Complex, Ipu As Complex
    Dim A As Double, kt As Double
    Dim Rm_pu As Double, Xm_pu As Double, Zm_pu As Double
    Dim Ra_g As Double, Xd_g As Double
    Dim Iprim As Complex, Isec As Complex

    ' Zberné polia výsledkov (maximálny možný počet riadkov)
    Dim maxN As Long, n As Long
    maxN = nBranches + nSwitches + nTrafo + nReaktory + nDifReaktory + nMotors + nGens
    If maxN < 1 Then maxN = 1
    Dim typArr() As String, nmArr() As String, odArr() As String, doArr() As String
    Dim iOdArr() As Double, iDoArr() As Double
    ReDim typArr(1 To maxN): ReDim nmArr(1 To maxN)
    ReDim odArr(1 To maxN): ReDim doArr(1 To maxN)
    ReDim iOdArr(1 To maxN): ReDim iDoArr(1 To maxN)
    n = 0

    ' Théveninova impedancia v uzle poruchy
    Zff = Z_inv(faultBus, faultBus)
    If Sqr(Zff.Re * Zff.Re + Zff.Im * Zff.Im) <= 0.0000001 Then
        Err.Raise vbObjectError + 31, "ComputeBranchContributions", _
            "Uzol '" & CStr(BusNames(faultBus)) & "': Théveninova impedancia je nulová - vetvové príspevky sa nedajú vypočítať."
    End If

    cF = GetVoltageFactorC(CDbl(BusBaseKV(faultBus)), caseMax)
    If_pu = CDiv(CCreate(cF, 0), Zff)

    ' Napätia uzlov počas skratu: V_i = c_f - Z_if·I_f
    ReDim Vbus(1 To nBuses)
    For i = 1 To nBuses
        Vbus(i) = CSub(CCreate(cF, 0), CMul(Z_inv(i, faultBus), If_pu))
    Next i

    ' Vedenia
    For k = 1 To nBranches
        If BranchStatus(k) > 0 Then
            If Not IsBranchIsolated(k) And Not (R(k) = 0 And X(k) = 0) Then
                i = FromBus(k): J = ToBus(k)
                Z = CCreate(CDbl(R(k)), CDbl(X(k)))
                Ipu = CDiv(CSub(Vbus(i), Vbus(J)), Z)
                n = n + 1
                typArr(n) = "vedenie": nmArr(n) = CStr(BranchName(k))
                odArr(n) = CStr(BusNames(i)): doArr(n) = CStr(BusNames(J))
                iOdArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
                iDoArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(J)), SBase_MVA)
            End If
        End If
    Next k

    ' Trafá (rovnaký model ako v Ysc: K_T len pre max prípad)
    For k = 1 To nTrafo
        If Not IsTrafoIsolated(k) And Not (TrR(k) = 0 And TrX(k) = 0) Then
            If caseMax Then kt = CDbl(TrKT(k)) Else kt = 1#
            If kt <= 0# Then kt = 1#
            i = TrFrom(k): J = TrTo(k): A = TrRatio(k)
            Z = CCreate(CDbl(TrR(k)) * kt, CDbl(TrX(k)) * kt)
            Ys = CDiv(CCreate(1, 0), Z)
            ' I_prim = V_i·ys/a² - V_j·ys/a ;  I_sec = V_j·ys - V_i·ys/a
            Iprim = CSub(CMul(Vbus(i), CCreate(Ys.Re / (A * A), Ys.Im / (A * A))), _
                         CMul(Vbus(J), CCreate(Ys.Re / A, Ys.Im / A)))
            Isec = CSub(CMul(Vbus(J), Ys), _
                        CMul(Vbus(i), CCreate(Ys.Re / A, Ys.Im / A)))
            n = n + 1
            typArr(n) = "trafo": nmArr(n) = CStr(TrName(k))
            odArr(n) = CStr(BusNames(i)): doArr(n) = CStr(BusNames(J))
            iOdArr(n) = CAbs(Iprim) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
            iDoArr(n) = CAbs(Isec) * IbaseKA(CDbl(BusBaseKV(J)), SBase_MVA)
        End If
    Next k

    ' Spínače
    For k = 1 To nSwitches
        If SwStatus(k) > 0 Then
            If Not IsSwitchIsolated(k) And Not (SwR(k) = 0 And SwX(k) = 0) Then
                i = SwFrom(k): J = SwTo(k)
                Z = CCreate(CDbl(SwR(k)), CDbl(SwX(k)))
                Ipu = CDiv(CSub(Vbus(i), Vbus(J)), Z)
                n = n + 1
                typArr(n) = "spinac": nmArr(n) = CStr(SwitchName(k))
                odArr(n) = CStr(BusNames(i)): doArr(n) = CStr(BusNames(J))
                iOdArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
                iDoArr(n) = iOdArr(n)
            End If
        End If
    Next k

    ' Reaktory
    For k = 1 To nReaktory
        If Not IsReaktorIsolated(k) And Not (ReaktorR(k) = 0 And ReaktorX(k) = 0) Then
            i = ReaktorFrom(k): J = ReaktorTo(k)
            Z = CCreate(CDbl(ReaktorR(k)), CDbl(ReaktorX(k)))
            Ipu = CDiv(CSub(Vbus(i), Vbus(J)), Z)
            n = n + 1
            typArr(n) = "reaktor": nmArr(n) = CStr(ReaktorName(k))
            odArr(n) = CStr(BusNames(i)): doArr(n) = CStr(BusNames(J))
            iOdArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
            iDoArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(J)), SBase_MVA)
        End If
    Next k

    ' Dif. reaktory
    For k = 1 To nDifReaktory
        If Not IsDifReaktorIsolated(k) And Not (DifReaktorR(k) = 0 And DifReaktorX(k) = 0) Then
            i = DifReaktorFrom(k): J = DifReaktorTo(k)
            Z = CCreate(CDbl(DifReaktorR(k)), CDbl(DifReaktorX(k)))
            Ipu = CDiv(CSub(Vbus(i), Vbus(J)), Z)
            n = n + 1
            typArr(n) = "dif.reaktor": nmArr(n) = CStr(DifReaktorName(k))
            odArr(n) = CStr(BusNames(i)): doArr(n) = CStr(BusNames(J))
            iOdArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
            iDoArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(J)), SBase_MVA)
        End If
    Next k

    ' Motory VN (príspevok len v max prípade, rovnako ako v Ysc)
    If caseMax Then
        For k = 1 To nMotors
            If MotorStatus(k) = 1 Then
                Zm_pu = Abs(CDbl(MotorXk(k)))
                If Zm_pu > 0.0000001 Then
                    i = MotorBus(k)
                    If Not IsBusIsolated(i) Then
                        Un = CDbl(BusBaseKV(i))
                        If SBase_MVA <> 0# And Un <> 0# Then
                            Zbase = (Un * Un) / SBase_MVA
                        Else
                            Zbase = 1#
                        End If
                        Rm_pu = CDbl(MotorR(k)) / Zbase
                        If Rm_pu > 0# And Rm_pu < Zm_pu Then
                            Xm_pu = Sqr(Zm_pu * Zm_pu - Rm_pu * Rm_pu)
                        Else
                            Xm_pu = Zm_pu / Sqr(1.01)
                            Rm_pu = 0.1 * Xm_pu
                        End If
                        Ipu = CDiv(Vbus(i), CCreate(Rm_pu, Xm_pu))
                        n = n + 1
                        typArr(n) = "motor": nmArr(n) = CStr(MotorName(k))
                        odArr(n) = CStr(BusNames(i)): doArr(n) = "-"
                        iOdArr(n) = CAbs(Ipu) * IbaseKA(Un, SBase_MVA)
                        iDoArr(n) = iOdArr(n)
                    End If
                End If
            End If
        Next k
    End If

    ' Generátory (I = V_t·y_G, y_G = 1/(K_G·(Ra'+jXd'')))
    For k = 1 To nGens
        If GenStatus(k) = 1 Then
            i = GenTermBus(k)
            If Not IsBusIsolated(i) And Not (GenRa(k) = 0 And GenXd(k) = 0) Then
                Ra_g = CDbl(GenRa(k))
                Xd_g = CDbl(GenXd(k))
                If Ra_g = 0# Then Ra_g = 0.07 * Xd_g
                Z = CCreate(CDbl(GenKG(k)) * Ra_g, CDbl(GenKG(k)) * Xd_g)
                Ipu = CDiv(Vbus(i), Z)
                n = n + 1
                typArr(n) = "generator": nmArr(n) = CStr(GenName(k))
                odArr(n) = CStr(BusNames(i)): doArr(n) = "-"
                iOdArr(n) = CAbs(Ipu) * IbaseKA(CDbl(BusBaseKV(i)), SBase_MVA)
                iDoArr(n) = iOdArr(n)
            End If
        End If
    Next k

    ' Zoradenie zostupne podľa väčšieho z prúdov (jednoduchý selection sort)
    Dim idx() As Long, tmp As Long, best As Long
    ReDim idx(1 To maxN)
    For i = 1 To n: idx(i) = i: Next i
    For i = 1 To n - 1
        best = i
        For J = i + 1 To n
            If MaxD(iOdArr(idx(J)), iDoArr(idx(J))) > MaxD(iOdArr(idx(best)), iDoArr(idx(best))) Then best = J
        Next J
        If best <> i Then
            tmp = idx(i): idx(i) = idx(best): idx(best) = tmp
        End If
    Next i

    ' Zápis do hárku "skrat_vetvy"
    Dim ws As Worksheet, rw As Long
    Set ws = GetOrCreateSheet("skrat_vetvy")
    ws.Cells.Clear
    ws.Cells(2, 2).Value = "Skrat v uzle:":  ws.Cells(2, 3).Value = CStr(BusNames(faultBus))
    ws.Cells(3, 2).Value = "Prípad:":        ws.Cells(3, 3).Value = IIf(caseMax, "max", "min")
    ws.Cells(4, 2).Value = "c [-]:":         ws.Cells(4, 3).Value = cF
    ws.Cells(5, 2).Value = "Ik'' [kA]:":     ws.Cells(5, 3).Value = Round(Ik_result(faultBus), 2)
    ws.Cells(6, 2).Value = "ip [kA]:":       ws.Cells(6, 3).Value = Round(ip_result(faultBus), 2)
    ws.Cells(2, 2).Resize(5, 1).Font.Bold = True

    rw = 8
    ws.Cells(rw, 2).Value = "Typ"
    ws.Cells(rw, 3).Value = "Meno"
    ws.Cells(rw, 4).Value = "Uzol od"
    ws.Cells(rw, 5).Value = "Uzol do"
    ws.Cells(rw, 6).Value = "I od [kA]"
    ws.Cells(rw, 7).Value = "I do [kA]"
    ws.Cells(rw, 2).Resize(1, 6).Font.Bold = True

    For i = 1 To n
        rw = rw + 1
        k = idx(i)
        ws.Cells(rw, 2).Value = typArr(k)
        ws.Cells(rw, 3).Value = nmArr(k)
        ws.Cells(rw, 4).Value = odArr(k)
        ws.Cells(rw, 5).Value = doArr(k)
        ws.Cells(rw, 6).Value = Round(iOdArr(k), 2)
        ws.Cells(rw, 7).Value = Round(iDoArr(k), 2)
    Next i
End Sub

' Bázový prúd v kA pre danú hladinu (S_base [MVA], Un [kV])
Private Function IbaseKA(ByVal Un_kV As Double, ByVal SBase_MVA As Double) As Double
    If Un_kV <> 0# Then
        IbaseKA = SBase_MVA / (Sqr(3) * Un_kV)
    Else
        IbaseKA = 0#
    End If
End Function

Private Function MaxD(ByVal A As Double, ByVal B As Double) As Double
    If A > B Then MaxD = A Else MaxD = B
End Function

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
