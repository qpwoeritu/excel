Option Explicit

' Načítanie dát uzlov z listu "uzly"
' Skutočné hodnoty -> prepočet do pomerných (p.u.)
'
' SBase_MVA - bázový výkon zo data!K11
' VLevels   - pole napäťových hladín zo data!K3:K8 (kV); každý uzol dostane
'             najbližšiu hladinu ako menovité napätie a Vmag = V_kV / U_base
'             ako prvý odhad (Slack/PV/PQ rovnako).
Public Sub LoadBusData( _
    ByRef nBuses As Long, _
    ByRef BusNames() As String, _
    ByRef BusTypes() As BusType, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef Pspec() As Double, _
    ByRef Qspec() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef VLevels() As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim t As String
    Dim V_kV As Double, P_MW As Double, Q_mvar As Double
    Dim busBase As Double
    Dim data As Variant

    Set ws = ThisWorkbook.Worksheets("uzly")

    ' posledný riadok podľa stĺpca B (Názov uzla)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        Err.Raise vbObjectError + 1, , "V liste 'uzly' nie sú žiadne uzly (očakávam dáta od riadku 3)."
    End If

    ' prvé dáta sú v riadku 3 => počet uzlov
    nBuses = lastRow - 2

    ReDim BusNames(1 To nBuses)
    ReDim BusTypes(1 To nBuses)
    ReDim Vmag(1 To nBuses)
    ReDim Vang(1 To nBuses)
    ReDim Pspec(1 To nBuses)
    ReDim Qspec(1 To nBuses)
    ReDim BusBaseKV(1 To nBuses)

    ' bulk read B..G (1 COM call) - stĺpce v poli: 1=B(meno), 2=C(typ), 3=D(V_kV), 5=F(P), 6=G(Q)
    data = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 7)).Value

    For i = 1 To nBuses
        BusNames(i) = CStr(data(i, 1))                      ' B: Názov uzla

        t = CStr(data(i, 2))                                ' C: Typ
        Select Case UCase$(Trim$(t))
            Case "SLACK"
                BusTypes(i) = btSlack
            Case "PQ"
                BusTypes(i) = btPQ
            Case "PV"
                BusTypes(i) = btPV
            Case Else
                BusTypes(i) = btPQ
        End Select

        ' načítanie skutočných hodnôt
        V_kV = ParseDouble(data(i, 3))                      ' D: |V| [kV]
        P_MW = ParseDouble(data(i, 5))                      ' F: P [MW]
        Q_mvar = ParseDouble(data(i, 6))                    ' G: Q [Mvar]

        ' Určenie bázy pre uzol - najbližšia hladina z data!K3:K8 (tvrdá kontrola pri >25 %)
        busBase = GetBaseVoltageForBus(V_kV, VLevels, BusNames(i))
        BusBaseKV(i) = busBase

        ' prvý odhad napätia v p.u. (rovnako pre Slack, PV aj PQ)
        Vmag(i) = V_kV / busBase

        ' Pôvodne načítanie uhla zo stĺpca E:
        ' Vang(i) = ParseDouble(ws.Cells(2 + i, 5).Value) * DEG2RAD
        ' Zmena (27.12.2025): Uhol pre prvý krok je vždy 0
        Vang(i) = 0#

        If SBase_MVA <> 0# Then
            Pspec(i) = P_MW / SBase_MVA
            Qspec(i) = Q_mvar / SBase_MVA
        Else
            Pspec(i) = 0#
            Qspec(i) = 0#
        End If
    Next i

    ' Postavíme dict raz, používa sa vo všetkých Load*Data subsekvenciách (O(1) lookup)
    Set busDict = BuildBusIndexDict(BusNames)
End Sub

' Načítanie dát transformátorov z listu "transformatory"
' C: Uzol od, D: Uzol do
' Detekcia stĺpcov pre Rk, Xk, G0, B0, Prevod
' Ak je prítomný stĺpec Zk (napr. na pozícii 18), posun +1.
Public Sub LoadTransformerData( _
    ByRef nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef TrR() As Double, _
    ByRef TrX() As Double, _
    ByRef TrG() As Double, _
    ByRef TrB() As Double, _
    ByRef TrRatio() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase_prim As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim G_siemens As Double, B_siemens As Double
    Dim Ubase1 As Double
    Dim colOffset As Long
    Dim data As Variant
    Dim lastCol As Long

    Set ws = ThisWorkbook.Worksheets("transformatory")

    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 3 Then
        nTrafo = 0
        Exit Sub
    End If

    nTrafo = lastRow - 2

    ReDim TrFrom(1 To nTrafo)
    ReDim TrTo(1 To nTrafo)
    ReDim TrR(1 To nTrafo)
    ReDim TrX(1 To nTrafo)
    ReDim TrG(1 To nTrafo)
    ReDim TrB(1 To nTrafo)
    ReDim TrRatio(1 To nTrafo)

    ' Detekcia posunu stĺpcov:
    If InStr(1, LCase(CStr(ws.Cells(2, 18).Value)), "zk") > 0 Then
        colOffset = 1
    Else
        colOffset = 0
    End If

    ' bulk read C..(22+colOffset) - stĺpce v poli: 1=C(from), 2=D(to), 16+co=R, 17+co=X, 18+co=G, 19+co=B, 20+co=ratio
    lastCol = 22 + colOffset
    data = ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, lastCol)).Value

    For i = 1 To nTrafo
        fromName = CStr(data(i, 1))                         ' C: Uzol od
        toName = CStr(data(i, 2))                           ' D: Uzol do

        idxFrom = GetBusIndexD(fromName, busDict)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 4, , "Uzol '" & fromName & "' v liste 'transformatory', riadok " & (i + 2) & " neexistuje."
        End If
        TrFrom(i) = idxFrom

        idxTo = GetBusIndexD(toName, busDict)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 5, , "Uzol '" & toName & "' v liste 'transformatory', riadok " & (i + 2) & " neexistuje."
        End If
        TrTo(i) = idxTo

        ' Báza impedancie na strane primáru (uzol od)
        Ubase1 = BusBaseKV(idxFrom)

        If SBase_MVA <> 0# Then
            Zbase_prim = (Ubase1 * Ubase1) / SBase_MVA
        Else
            Zbase_prim = 1#
        End If

        ' R, X v Ohmoch -> p.u.   (col 18+co = pole-col 16+co, col 19+co = 17+co)
        R_ohm = ParseDouble(data(i, 16 + colOffset))
        X_ohm = ParseDouble(data(i, 17 + colOffset))

        If Zbase_prim <> 0# Then
            TrR(i) = R_ohm / Zbase_prim
            TrX(i) = X_ohm / Zbase_prim
        Else
            TrR(i) = 0#
            TrX(i) = 0#
        End If

        ' G, B v Siemensoch -> p.u.
        G_siemens = ParseDouble(data(i, 18 + colOffset))
        B_siemens = ParseDouble(data(i, 19 + colOffset))

        TrG(i) = G_siemens * Zbase_prim
        TrB(i) = B_siemens * Zbase_prim

        ' Prevod a
        TrRatio(i) = ParseDouble(data(i, 20 + colOffset))
        If TrRatio(i) <= 0# Then TrRatio(i) = 1#
    Next i
End Sub

' Načítanie dát reaktorov z listu "reaktory"
' B: Označenie, C: Uzol od, D: Uzol do, H: R[ohm], I: X[ohm]
Public Sub LoadReactorData( _
    ByRef nReaktory As Long, _
    ByRef ReaktorName() As String, _
    ByRef ReaktorFrom() As Long, _
    ByRef ReaktorTo() As Long, _
    ByRef ReaktorR() As Double, _
    ByRef ReaktorX() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase As Double
    Dim data As Variant

    Set ws = GetOrCreateSheet("reaktory")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        nReaktory = 0
        Exit Sub
    End If

    nReaktory = lastRow - 3

    ReDim ReaktorName(1 To nReaktory)
    ReDim ReaktorFrom(1 To nReaktory)
    ReDim ReaktorTo(1 To nReaktory)
    ReDim ReaktorR(1 To nReaktory)
    ReDim ReaktorX(1 To nReaktory)

    ' bulk read B..I (1=B name, 2=C from, 3=D to, 7=H R, 8=I X)
    data = ws.Range(ws.Cells(4, 2), ws.Cells(lastRow, 9)).Value

    For i = 1 To nReaktory
        ReaktorName(i) = CStr(data(i, 1))                   ' B
        fromName = CStr(data(i, 2))                         ' C
        toName = CStr(data(i, 3))                           ' D

        idxFrom = GetBusIndexD(fromName, busDict)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 6, , "Uzol '" & fromName & "' v liste 'reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        ReaktorFrom(i) = idxFrom

        idxTo = GetBusIndexD(toName, busDict)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 7, , "Uzol '" & toName & "' v liste 'reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        ReaktorTo(i) = idxTo

        ' Impedančná báza podľa uzla "od"
        Ubase = BusBaseKV(idxFrom)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If

        R_ohm = ParseDouble(data(i, 7))                     ' H (8) -> col 7 v poli
        X_ohm = ParseDouble(data(i, 8))                     ' I (9) -> col 8

        If Zbase <> 0# Then
            ReaktorR(i) = R_ohm / Zbase
            ReaktorX(i) = X_ohm / Zbase
        Else
            ReaktorR(i) = 0#
            ReaktorX(i) = 0#
        End If
    Next i
End Sub

' Načítanie dát dif. reaktorov z listu "dif_reaktory"
' C: Uzol od, D: Uzol do, I: R[ohm], J: X[ohm]
' Hlavička: riadok 3, dáta od 4
Public Sub LoadDifReactorData( _
    ByRef nDifReaktory As Long, _
    ByRef DifReaktorName() As String, _
    ByRef DifReaktorFrom() As Long, _
    ByRef DifReaktorTo() As Long, _
    ByRef DifReaktorR() As Double, _
    ByRef DifReaktorX() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idxFrom As Long, idxTo As Long
    Dim Zbase As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase As Double
    Dim data As Variant

    Set ws = GetOrCreateSheet("dif_reaktory")

    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 4 Then
        nDifReaktory = 0
        Exit Sub
    End If

    nDifReaktory = lastRow - 3

    ReDim DifReaktorName(1 To nDifReaktory)
    ReDim DifReaktorFrom(1 To nDifReaktory)
    ReDim DifReaktorTo(1 To nDifReaktory)
    ReDim DifReaktorR(1 To nDifReaktory)
    ReDim DifReaktorX(1 To nDifReaktory)

    ' bulk read C..J (1=C from, 2=D to, 7=I R, 8=J X)
    data = ws.Range(ws.Cells(4, 3), ws.Cells(lastRow, 10)).Value

    For i = 1 To nDifReaktory
        ' Name nie je v špecifikácii, ale pole existuje. Dáme tam "DR" + index alebo prázdny string?
        DifReaktorName(i) = "DR" & i
        fromName = CStr(data(i, 1))                         ' C
        toName = CStr(data(i, 2))                           ' D

        idxFrom = GetBusIndexD(fromName, busDict)
        If idxFrom = 0 Then
            Err.Raise vbObjectError + 8, , "Uzol '" & fromName & "' v liste 'dif_reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        DifReaktorFrom(i) = idxFrom

        idxTo = GetBusIndexD(toName, busDict)
        If idxTo = 0 Then
            Err.Raise vbObjectError + 9, , "Uzol '" & toName & "' v liste 'dif_reaktory', riadok " & (i + 3) & " neexistuje."
        End If
        DifReaktorTo(i) = idxTo

        ' Impedančná báza podľa uzla "od"
        Ubase = BusBaseKV(idxFrom)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If

        R_ohm = ParseDouble(data(i, 7))                     ' I (9) -> col 7 v poli
        X_ohm = ParseDouble(data(i, 8))                     ' J (10) -> col 8

        If Zbase <> 0# Then
            DifReaktorR(i) = R_ohm / Zbase
            DifReaktorX(i) = X_ohm / Zbase
        Else
            DifReaktorR(i) = 0#
            DifReaktorX(i) = 0#
        End If
    Next i
End Sub

' Načítanie kompenzácie z listu "kompenzácia"
' B: Názov, C: Uzol, N: Status (1/0), P: X_L[ohm], Q: X_C[ohm]
' Hlavička: riadok 3, dáta od 4
Public Sub LoadCompData( _
    ByRef nComp As Long, _
    ByRef CompName() As String, _
    ByRef CompBus() As Long, _
    ByRef CompB() As Double, _
    ByRef CompStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim busName As String, idxBus As Long
    Dim XL_ohm As Double, XC_ohm As Double
    Dim Zbase As Double, Ubase As Double
    Dim BL_pu As Double, BC_pu As Double
    Dim statusVal As Variant
    Dim data As Variant

    Set ws = GetOrCreateSheet("kompenzácia")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        nComp = 0
        Exit Sub
    End If

    nComp = lastRow - 3

    ReDim CompName(1 To nComp)
    ReDim CompBus(1 To nComp)
    ReDim CompB(1 To nComp)
    ReDim CompStatus(1 To nComp)

    ' bulk read B..Q (1=B name, 2=C bus, 13=N status, 15=P XL, 16=Q XC)
    data = ws.Range(ws.Cells(4, 2), ws.Cells(lastRow, 17)).Value

    For i = 1 To nComp
        CompName(i) = CStr(data(i, 1))                      ' B
        busName = CStr(data(i, 2))                          ' C

        idxBus = GetBusIndexD(busName, busDict)
        If idxBus = 0 Then
            Err.Raise vbObjectError + 10, , "Uzol '" & busName & "' v liste 'kompenzácia', riadok " & (i + 3) & " neexistuje."
        End If
        CompBus(i) = idxBus

        ' Status N (14) -> col 13 v poli
        statusVal = data(i, 13)
        If IsNumeric(statusVal) Then
            CompStatus(i) = CInt(statusVal)
        Else
            CompStatus(i) = 0
        End If

        ' Ak nie je aktívna (Status=0), susceptancia je 0
        If CompStatus(i) = 1 Then
            Ubase = BusBaseKV(idxBus)
            If SBase_MVA <> 0# Then
                Zbase = (Ubase * Ubase) / SBase_MVA
            Else
                Zbase = 1#
            End If

            XL_ohm = ParseDouble(data(i, 15))               ' P (16) -> col 15
            XC_ohm = ParseDouble(data(i, 16))               ' Q (17) -> col 16
            
            ' Sériová kombinácia / Výsledná reaktancia:
            ' Užívateľ požaduje 1/(XC - XL)
            ' Ak XL=0 a XC>0 -> 1/XC (Kapacitný, B > 0)
            ' Ak XL>0 a XC=0 -> 1/-XL (Induktívny, B < 0)
            ' Ak sú zadané obe, rozhoduje rozdiel.
            
            Dim X_net_ohm As Double
            X_net_ohm = XC_ohm - XL_ohm
            
            If Zbase <> 0# And Abs(X_net_ohm) > 0.000001 Then
                ' B [S] = 1 / X_net [Ohm]
                ' B [p.u.] = B [S] * Zbase
                CompB(i) = (1# / X_net_ohm) * Zbase
            Else
                CompB(i) = 0#
            End If
        Else
            CompB(i) = 0#
        End If
    Next i
End Sub

' Načítanie dát generátorov z listu "generatory"
' B: Názov, C: PQ_Uzol, D: PV_Uzol, E: Status, L: Ra, M: P_gen, O: Xs, P: Xd''
Public Sub LoadGeneratorData( _
    ByRef nGens As Long, _
    ByRef GenName() As String, _
    ByRef GenBusPQ() As Long, _
    ByRef GenBusPV() As Long, _
    ByRef GenStatus() As Integer, _
    ByRef GenRa() As Double, _
    ByRef GenP() As Double, _
    ByRef GenXs() As Double, _
    ByRef GenXd() As Double, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim pqName As String, pvName As String
    Dim idxPQ As Long, idxPV As Long
    Dim Ubase As Double, Zbase As Double
    Dim stVal As Variant
    Dim data As Variant

    Set ws = GetOrCreateSheet("generatory")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nGens = 0
        Exit Sub
    End If

    nGens = lastRow - 2

    ReDim GenName(1 To nGens)
    ReDim GenBusPQ(1 To nGens)
    ReDim GenBusPV(1 To nGens)
    ReDim GenStatus(1 To nGens)
    ReDim GenRa(1 To nGens)
    ReDim GenP(1 To nGens)
    ReDim GenXs(1 To nGens)
    ReDim GenXd(1 To nGens)

    ' bulk read B..P (1=B name, 2=C pq, 3=D pv, 4=E status, 11=L Ra, 12=M P, 14=O Xs, 15=P Xd'')
    data = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 16)).Value

    For i = 1 To nGens
        GenName(i) = CStr(data(i, 1))                       ' B
        pqName = CStr(data(i, 2))                           ' C
        pvName = CStr(data(i, 3))                           ' D

        ' Status E (5) -> col 4
        stVal = data(i, 4)
        If IsNumeric(stVal) Then
            GenStatus(i) = CInt(stVal)
        Else
            GenStatus(i) = 0
        End If

        idxPQ = GetBusIndexD(pqName, busDict)
        If idxPQ = 0 Then
            Err.Raise vbObjectError + 12, , "PQ Uzol '" & pqName & "' v liste 'generatory' neexistuje."
        End If
        GenBusPQ(i) = idxPQ

        idxPV = GetBusIndexD(pvName, busDict)
        If idxPV = 0 Then
            Err.Raise vbObjectError + 13, , "PV Uzol '" & pvName & "' v liste 'generatory' neexistuje."
        End If
        GenBusPV(i) = idxPV

        ' P_gen [MW] -> M (13) -> col 12
        If SBase_MVA <> 0 Then
            GenP(i) = ParseDouble(data(i, 12)) / SBase_MVA
        Else
            GenP(i) = 0
        End If

        ' Ra [ohm] -> L (12) -> col 11, Xs [ohm] -> O (15) -> col 14, Xd'' [ohm] -> P (16) -> col 15
        Ubase = BusBaseKV(idxPQ)
        If SBase_MVA <> 0 Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1
        End If

        If Zbase <> 0 Then
            GenRa(i) = ParseDouble(data(i, 11)) / Zbase
            GenXs(i) = ParseDouble(data(i, 14)) / Zbase
            GenXd(i) = ParseDouble(data(i, 15)) / Zbase
        End If
    Next i
End Sub

' Zápis výsledkov generátorov (Load Flow) - Q do stĺpca N
Public Sub WriteGeneratorResults( _
    ByVal nGens As Long, _
    ByRef GenBusPQ() As Long, _
    ByRef GenBusPV() As Long, _
    ByRef GenXs() As Double, _
    ByRef GenRa() As Double, _
    ByRef GenStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim idxPQ As Long, idxPV As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zs As Complex, I_pu As Complex, S_pu As Complex
    Dim Q_mvar As Double
    
    Set ws = GetOrCreateSheet("generatory")
    ' Hlavička
    ws.Cells(2, 14).Value = "Q_gen [Mvar]" ' N (14)
    
    For i = 1 To nGens
        If GenStatus(i) = 1 Then
            idxPQ = GenBusPQ(i)
            idxPV = GenBusPV(i)
            
            ' Tok z PV do PQ? Nie, Q_gen je injekcia do PQ uzla (alebo výroba generátora).
            ' Q_gen sa meria na svorkách (PQ uzol).
            ' Prúd I = (V_pv - V_pq) / (Ra + jXs)
            ' S_pq = V_pq * conj(I) ... to je príkon do PQ uzla z vetvy.
            ' Ak je S_pq kladné, tak to "tečie do uzla".
            
            Vi = CFromPolar(Vmag(idxPV), Vang(idxPV) * RAD2DEG)
            Vj = CFromPolar(Vmag(idxPQ), Vang(idxPQ) * RAD2DEG)
            Zs = CCreate(GenRa(i), GenXs(i))
            
            ' I = (Vi - Vj) / Zs
            I_pu = CDiv(CSub(Vi, Vj), Zs)
            
            ' Výkon dodávaný do siete (na svorkách j)
            ' S_gen = Vj * conj(I)
            ' (Pozor na znamienko: I tečie z PV do PQ. Takže I vstupuje do PQ. S = U*I*)
            S_pu = CMul(Vj, CConj(I_pu))
            
            Q_mvar = S_pu.Im * SBase_MVA
            
            ws.Cells(i + 2, 14).Value = Round(Q_mvar, 2)
        Else
            ws.Cells(i + 2, 14).Value = 0
        End If
    Next i
End Sub

' Načítanie motorov z listu "motoryVN"
' B: Názov, C: Uzol, L: R[ohm], P: Xk[ohm], Q: G[S], R: B[S], S: Status
' Prepočet Xk na p.u. (Zbase), G a B na p.u.
Public Sub LoadMotorData( _
    ByRef nMotors As Long, _
    ByRef MotorName() As String, _
    ByRef MotorBus() As Long, _
    ByRef MotorR() As Double, _
    ByRef MotorXk() As Double, _
    ByRef MotorG() As Double, _
    ByRef MotorB() As Double, _
    ByRef MotorStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim busName As String, idxBus As Long
    Dim Ubase As Double, Zbase As Double
    Dim R_ohm As Double, Xk_ohm As Double
    Dim G_siemens As Double, B_siemens As Double
    Dim statusVal As Variant
    Dim data As Variant

    Set ws = GetOrCreateSheet("motoryVN")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nMotors = 0
        Exit Sub
    End If

    nMotors = lastRow - 2

    ReDim MotorName(1 To nMotors)
    ReDim MotorBus(1 To nMotors)
    ReDim MotorR(1 To nMotors)
    ReDim MotorXk(1 To nMotors)
    ReDim MotorG(1 To nMotors)
    ReDim MotorB(1 To nMotors)
    ReDim MotorStatus(1 To nMotors)

    ' bulk read B..S (1=B name, 2=C bus, 11=L R, 15=P Xk, 16=Q G, 17=R B, 18=S status)
    data = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 19)).Value

    For i = 1 To nMotors
        MotorName(i) = CStr(data(i, 1))                     ' B
        busName = CStr(data(i, 2))                          ' C

        idxBus = GetBusIndexD(busName, busDict)
        If idxBus = 0 Then
            Err.Raise vbObjectError + 11, , "Uzol '" & busName & "' v liste 'motoryVN', riadok " & (i + 2) & " neexistuje."
        End If
        MotorBus(i) = idxBus

        ' Status S (19) -> col 18
        statusVal = data(i, 18)
        If IsNumeric(statusVal) Then
            MotorStatus(i) = CInt(statusVal)
        Else
            MotorStatus(i) = 0
        End If

        ' R [ohm] L (12) -> col 11
        MotorR(i) = ParseDouble(data(i, 11))

        ' Báza pre Xk, G, B
        Ubase = BusBaseKV(idxBus)
        If SBase_MVA <> 0# Then
            Zbase = (Ubase * Ubase) / SBase_MVA
        Else
            Zbase = 1#
        End If

        If MotorStatus(i) = 1 Then
            ' Xk [ohm] P (16) -> col 15
            Xk_ohm = ParseDouble(data(i, 15))
            If Zbase <> 0# Then
                MotorXk(i) = Xk_ohm / Zbase
            Else
                MotorXk(i) = 0#
            End If

            ' G [S] Q (17) -> col 16
            G_siemens = ParseDouble(data(i, 16))
            MotorG(i) = G_siemens * Zbase

            ' B [S] R (18) -> col 17
            B_siemens = ParseDouble(data(i, 17))
            MotorB(i) = B_siemens * Zbase
        Else
            MotorXk(i) = 0#
            MotorG(i) = 0#
            MotorB(i) = 0#
        End If
    Next i
End Sub

' Zápis výsledkov motorov (Load Flow) - I [A] do AD, Ploss [kW] do AE
Public Sub WriteMotorResults( _
    ByVal nMotors As Long, _
    ByRef MotorBus() As Long, _
    ByRef MotorR() As Double, _
    ByRef MotorG() As Double, _
    ByRef MotorB() As Double, _
    ByRef MotorStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim idxBus As Long
    Dim V_pu As Double
    Dim Y_mod_pu As Double
    Dim I_pu As Double, I_abs As Double, Ibase As Double
    Dim Ploss_kW As Double
    
    Set ws = GetOrCreateSheet("motoryVN")
    
    ' Hlavička
    ws.Cells(2, 30).Value = "I [A]"     ' AD (30)
    ws.Cells(2, 31).Value = "Ploss [kW]" ' AE (31)
    
    For i = 1 To nMotors
        If MotorStatus(i) = 1 Then
            idxBus = MotorBus(i)
            V_pu = Vmag(idxBus)
            
            ' I_pu = V_pu * |Y_motor_pu|
            ' Y_motor = G + jB
            Y_mod_pu = Sqr(MotorG(i) * MotorG(i) + MotorB(i) * MotorB(i))
            I_pu = V_pu * Y_mod_pu
            
            ' Ibase [kA] = Sbase / (sqrt(3)*Ubase) -> *1000 = [A]
            If BusBaseKV(idxBus) <> 0 Then
                Ibase = (SBase_MVA * 1000#) / (Sqr(3) * BusBaseKV(idxBus))
            Else
                Ibase = 0
            End If
            
            I_abs = I_pu * Ibase
            
            ' Ploss = 3 * R * I^2  (I v A, R v Ohm -> W -> /1000 -> kW)
            Ploss_kW = (3# * MotorR(i) * I_abs * I_abs) / 1000#
            
            ws.Cells(i + 2, 30).Value = Round(I_abs, 2)
            ws.Cells(i + 2, 31).Value = Round(Ploss_kW, 2)
        Else
            ws.Cells(i + 2, 30).Value = 0
            ws.Cells(i + 2, 31).Value = 0
        End If
    Next i
End Sub

' Zápis výsledkov kompenzácie (Load Flow) - Napätie do stĺpca O
Public Sub WriteCompResults( _
    ByVal nComp As Long, _
    ByRef CompBus() As Long, _
    ByRef Vmag() As Double, _
    ByRef BusBaseKV() As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim U_kV As Double
    Dim idxBus As Long
    
    Set ws = GetOrCreateSheet("kompenzácia")
    
    ' Hlavička
    ws.Cells(3, 15).Value = "U [kV]" ' O (15)
    
    For i = 1 To nComp
        idxBus = CompBus(i)
        U_kV = Vmag(idxBus) * BusBaseKV(idxBus)
        ws.Cells(i + 3, 15).Value = Round(U_kV, 2)
    Next i
End Sub

' Zápis výsledkov reaktorov (Load Flow)
' Z: I[A], AA: dU[%], AB: P[MW], AC: Q[MVAr], AD: Ploss[kW]
Public Sub WriteReactorResults( _
    ByVal nReaktory As Long, _
    ByRef ReaktorFrom() As Long, _
    ByRef ReaktorTo() As Long, _
    ByRef ReaktorR() As Double, _
    ByRef ReaktorX() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_kW As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("reaktory")
    
    ' Hlavičky (riadok 3)
    ws.Cells(3, 26).Value = "I [A]"
    ws.Cells(3, 27).Value = "dU [%]"
    ws.Cells(3, 28).Value = "P [MW]"
    ws.Cells(3, 29).Value = "Q [MVAr]"
    ws.Cells(3, 30).Value = "Ploss [kW]"
    
    For k = 1 To nReaktory
        iBus = ReaktorFrom(k)
        jBus = ReaktorTo(k)
        
        Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
        Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
        
        Zpu = CCreate(ReaktorR(k), ReaktorX(k))
        
        If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
            Iabs_pu = 0#
            dU_percent = 0#
            Pij_MW = 0#
            Qij_MVAr = 0#
            Ploss_kW = 0#
        Else
            ' I = (Vi - Vj) / Z
            Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
            Iabs_pu = CAbs(Iij_pu)
            
            ' dU = (|Vi| - |Vj|) * 100
            dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
            
            ' S = Vi * conj(I)
            Sij_pu = CMul(Vi, CConj(Iij_pu))
            Pij_MW = Sij_pu.Re * SBase_MVA
            Qij_MVAr = Sij_pu.Im * SBase_MVA
            
            ' Ploss = |I|^2 * R * Sbase * 1000
            Ploss_kW = Iabs_pu * Iabs_pu * ReaktorR(k) * SBase_MVA * 1000#
        End If
        
        ' I [A]
        Ubase = BusBaseKV(iBus)
        If SBase_MVA <> 0# And Ubase <> 0# Then
            Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase)
        Else
            Ibase_A = 1#
        End If
        Iabs_A = Iabs_pu * Ibase_A
        
        ws.Cells(k + 3, 26).Value = Round(Iabs_A, 2)
        ws.Cells(k + 3, 27).Value = Round(dU_percent, 2)
        ws.Cells(k + 3, 28).Value = Round(Pij_MW, 2)
        ws.Cells(k + 3, 29).Value = Round(Qij_MVAr, 2)
        ws.Cells(k + 3, 30).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Zápis výsledkov dif. reaktorov (Load Flow)
' X: I[A], Y: dU[%], Z: P[MW], AA: Q[MVAr], AB: Ploss[kW]
Public Sub WriteDifReactorResults( _
    ByVal nDifReaktory As Long, _
    ByRef DifReaktorFrom() As Long, _
    ByRef DifReaktorTo() As Long, _
    ByRef DifReaktorR() As Double, _
    ByRef DifReaktorX() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_kW As Double
    Dim Ubase As Double
    
    Set ws = GetOrCreateSheet("dif_reaktory")
    
    ' Hlavičky (riadok 3) od stĺpca X (24)
    ws.Cells(3, 24).Value = "I [A]"
    ws.Cells(3, 25).Value = "dU [%]"
    ws.Cells(3, 26).Value = "P [MW]"
    ws.Cells(3, 27).Value = "Q [MVAr]"
    ws.Cells(3, 28).Value = "Ploss [kW]"
    
    For k = 1 To nDifReaktory
        iBus = DifReaktorFrom(k)
        jBus = DifReaktorTo(k)
        
        Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
        Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
        
        Zpu = CCreate(DifReaktorR(k), DifReaktorX(k))
        
        If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
            Iabs_pu = 0#
            dU_percent = 0#
            Pij_MW = 0#
            Qij_MVAr = 0#
            Ploss_kW = 0#
        Else
            Iij_pu = CDiv(CSub(Vi, Vj), Zpu)
            Iabs_pu = CAbs(Iij_pu)
            
            dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
            
            Sij_pu = CMul(Vi, CConj(Iij_pu))
            Pij_MW = Sij_pu.Re * SBase_MVA
            Qij_MVAr = Sij_pu.Im * SBase_MVA
            
            Ploss_kW = Iabs_pu * Iabs_pu * DifReaktorR(k) * SBase_MVA * 1000#
        End If
        
        ' I [A]
        Ubase = BusBaseKV(iBus)
        If SBase_MVA <> 0# And Ubase <> 0# Then
            Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase)
        Else
            Ibase_A = 1#
        End If
        Iabs_A = Iabs_pu * Ibase_A
        
        ws.Cells(k + 3, 24).Value = Round(Iabs_A, 2)
        ws.Cells(k + 3, 25).Value = Round(dU_percent, 2)
        ws.Cells(k + 3, 26).Value = Round(Pij_MW, 2)
        ws.Cells(k + 3, 27).Value = Round(Qij_MVAr, 2)
        ws.Cells(k + 3, 28).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Zápis admitančnej matice a pomocných blokov G a B
Public Sub WriteYMatrix( _
    ByRef Y() As Complex, _
    ByRef G() As Double, _
    ByRef B() As Double, _
    ByRef BusNames() As String)

    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long, J As Long
    Dim row0 As Long, col0 As Long
    Dim arr As Variant
    Dim startRowB As Long

    Set ws = GetOrCreateSheet("Y_matica")
    ws.Cells.Clear

    n = UBound(BusNames)

    row0 = 1
    col0 = 1

    ' Blok G = Re(Y) - hlavička + matica v jednom Variant poli, jeden Range.Value zápis
    ws.Cells(row0, col0).Value = "G = Re(Y)"
    ReDim arr(1 To n + 1, 1 To n + 1)
    arr(1, 1) = ""
    For J = 1 To n
        arr(1, J + 1) = BusNames(J)
    Next J
    For i = 1 To n
        arr(i + 1, 1) = BusNames(i)
        For J = 1 To n
            arr(i + 1, J + 1) = G(i, J)
        Next J
    Next i
    ws.Range(ws.Cells(row0 + 1, col0), ws.Cells(row0 + 1 + n, col0 + n)).Value = arr

    ' Blok B = Im(Y)
    startRowB = row0 + n + 3
    ws.Cells(startRowB, col0).Value = "B = Im(Y)"
    ReDim arr(1 To n + 1, 1 To n + 1)
    arr(1, 1) = ""
    For J = 1 To n
        arr(1, J + 1) = BusNames(J)
    Next J
    For i = 1 To n
        arr(i + 1, 1) = BusNames(i)
        For J = 1 To n
            arr(i + 1, J + 1) = B(i, J)
        Next J
    Next i
    ws.Range(ws.Cells(startRowB + 1, col0), ws.Cells(startRowB + 1 + n, col0 + n)).Value = arr
End Sub

' Vymazanie / príprava výsledkových listov "napatia" a "epsilon"
Public Sub ClearResultsSheets()
    Dim wsV As Worksheet, wsE As Worksheet
    
    Set wsV = GetOrCreateSheet("napatia")
    Set wsE = GetOrCreateSheet("epsilon")
    
    wsV.Cells.Clear
    wsE.Cells.Clear
    
    ' hlavičky
    wsV.Cells(1, 1).Value = "Iterácia"
    wsV.Cells(1, 2).Value = "Uzol"
    wsV.Cells(1, 3).Value = "|V| [p.u.]"
    wsV.Cells(1, 4).Value = "? [deg]"
    
    wsE.Cells(1, 1).Value = "Iterácia"
    wsE.Cells(1, 2).Value = "max|?P|"
    wsE.Cells(1, 3).Value = "max|?Q|"
    wsE.Cells(1, 4).Value = "epsilon"
End Sub

' Flush buffrov napätí / epsilon po skončení NR slučky (jeden Range.Value zápis na list)
' VBuf má 4 stĺpce: iter, BusName, |V|, theta_deg; VRow = počet naplnených riadkov.
Public Sub FlushVoltageLog(ByRef VBuf As Variant, ByVal VRow As Long)
    If VRow < 1 Then Exit Sub

    Dim ws As Worksheet
    Dim outArr As Variant
    Dim r As Long, c As Long

    Set ws = GetOrCreateSheet("napatia")

    ReDim outArr(1 To VRow, 1 To 4)
    For r = 1 To VRow
        For c = 1 To 4
            outArr(r, c) = VBuf(r, c)
        Next c
    Next r

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + VRow, 4)).Value = outArr
End Sub

' EBuf má 4 stĺpce: iter, maxDP, maxDQ, eps; ERow = počet naplnených riadkov.
Public Sub FlushEpsilonLog(ByRef EBuf As Variant, ByVal ERow As Long)
    If ERow < 1 Then Exit Sub

    Dim ws As Worksheet
    Dim outArr As Variant
    Dim r As Long, c As Long

    Set ws = GetOrCreateSheet("epsilon")

    ReDim outArr(1 To ERow, 1 To 4)
    For r = 1 To ERow
        For c = 1 To 4
            outArr(r, c) = EBuf(r, c)
        Next c
    Next r

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + ERow, 4)).Value = outArr
End Sub

' Zápis súhrnných výsledkov do listu "index"
Public Sub WriteSummaryToIndex(ByVal totalTime As Double, _
                               ByVal iterCount As Long, _
                               ByVal epsFinal As Double, _
                               ByVal converged As Boolean)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("index")
    
    ws.Range("B6").Value = totalTime
    ws.Range("B7").Value = iterCount
    ws.Range("B8").Value = epsFinal
    ws.Range("B9").Value = IIf(converged, "Konvergovalo", "Nekonvergovalo")
End Sub

' Zápis výsledných napätí z NR výpočtu na list "uzly"
' H (stĺpec 8): |V| výp. [kV]
' I (stĺpec 9): ? výp. [deg]
'
' Vmag(i) je v p.u., prepočet na kV cez BusBaseKV(i)
Public Sub WriteFinalVoltagesToUzly( _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim nBuses As Long
    
    Set ws = ThisWorkbook.Worksheets("uzly")
    
    nBuses = UBound(Vmag)
    
    ' hlavičky výsledkov (riadok 2)
    ws.Cells(2, 8).Value = "|V| výp. [kV]"
    ws.Cells(2, 9).Value = "? výp. [deg]"
    
    ' dáta od riadku 3
    For i = 1 To nBuses
        ws.Cells(2 + i, 8).Value = Round(Vmag(i) * BusBaseKV(i), 2)
        ws.Cells(2 + i, 9).Value = Round(Vang(i) * RAD2DEG, 2)
    Next i
End Sub

' Výpočet prúdov a tokov výkonu vo vedeniach po NR výpočte
' Vnútorný výpočet: v p.u.
' Výstupy na list "vedenia":
'   F: |I_ij| [A]
'   G: ?U [%]        (V_from - V_to v p.u. * 100)
'   H: P_ij [MW]     (činný výkon z "uzol-od" do "uzol-do")
'   I: Q_ij [MVAr]   (jalový výkon z "uzol-od" do "uzol-do")
'   J: P_str [kW]    (činné straty na vedení)
'
' R(), X()  – v p.u. (po prepočte v LoadBranchData)
' Vmag(), Vang() – napätia uzlov v p.u., Vang v radianoch
' SBase_MVA – báza výkonu
' BusBaseKV() - bázy napätí uzlov
Public Sub WriteBranchCurrents( _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BranchStatus() As Integer, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef BusBaseKV() As Double, _
    ByRef Bshunt() As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim iBus As Long, jBus As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zpu As Complex, Iij_pu As Complex
    Dim Iabs_pu As Double, Iabs_A As Double
    Dim Ibase_A As Double
    Dim dU_percent As Double
    Dim Sij_pu As Complex
    Dim Pij_pu As Double, Qij_pu As Double
    Dim Pij_MW As Double, Qij_MVAr As Double
    Dim Ploss_pu As Double, Ploss_kW As Double
    Dim Ubase_line As Double
    
    Set ws = ThisWorkbook.Worksheets("vedenia")
    
    ' hlavičky stĺpcov - riadok 2, stĺpce Q(17)..U(21)
    ws.Cells(2, 17).Value = "|I_ij| [A]"
    ws.Cells(2, 18).Value = "?U [%]"
    ws.Cells(2, 19).Value = "P_ij [MW]"
    ws.Cells(2, 20).Value = "Q_ij [MVAr]"
    ws.Cells(2, 21).Value = "P_str [kW]"
    
    ' vetvy sú v riadkoch 3..(nBranches+2)
    For k = 1 To nBranches
        ' Ak je vedenie vypnuté (Status = 0), zapíšeme nuly
        If BranchStatus(k) = 0 Then
            ws.Cells(2 + k, 17).Value = 0
            ws.Cells(2 + k, 18).Value = 0
            ws.Cells(2 + k, 19).Value = 0
            ws.Cells(2 + k, 20).Value = 0
            ws.Cells(2 + k, 21).Value = 0
        Else
            iBus = FromBus(k)
            jBus = ToBus(k)
            
            ' komplexné napätia uzlov v p.u. (polárny -> karteziánsky)
            ' Vmag je v p.u., Vang v radianoch -> CFromPolar očakáva uhol v stupňoch
            Vi = CFromPolar(Vmag(iBus), Vang(iBus) * RAD2DEG)
            Vj = CFromPolar(Vmag(jBus), Vang(jBus) * RAD2DEG)
            
            ' impedancia vetvy v p.u.
            Zpu = CCreate(R(k), X(k))
            
            If Abs(Zpu.Re) < 0.000000001 And Abs(Zpu.Im) < 0.000000001 Then
                ' nulová impedancia – prúd a výkony nedefinované; nastavíme 0
                Iabs_pu = 0#
                dU_percent = 0#
                Pij_pu = 0#
                Qij_pu = 0#
                Ploss_pu = 0#
            Else
                ' prúd vetvou v p.u.: I_ij = (Vi - Vj) / Z_ij + Vi * j*B/2 (PI článok)
                Dim I_series As Complex, I_shunt As Complex
                I_series = CDiv(CSub(Vi, Vj), Zpu)
                I_shunt = CMul(Vi, CCreate(0, Bshunt(k) / 2#))
                Iij_pu = CAdd(I_series, I_shunt)
                Iabs_pu = CAbs(Iij_pu)
                
                ' úbytok napätia v % (z pohľadu "uzol-od" -> "uzol-do"):
                ' ?U [%] = (V_from_pu - V_to_pu) * 100
                dU_percent = (Vmag(iBus) - Vmag(jBus)) * 100#
                
                ' tok výkonu z uzla-od do uzla-do:
                ' S_ij_pu = V_i * conj(I_ij)
                Sij_pu = CMul(Vi, CConj(Iij_pu))
                Pij_pu = Sij_pu.Re
                Qij_pu = Sij_pu.Im
                
                ' činné straty v p.u.:
                ' P_str_pu = |I|^2 * R_pu
                Ploss_pu = Iabs_pu * Iabs_pu * R(k)
            End If
            
            ' Určenie bázy prúdu pre vedenie (podľa "from" uzla)
            Ubase_line = BusBaseKV(iBus)
            If SBase_MVA <> 0# And Ubase_line <> 0# Then
                Ibase_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase_line)
            Else
                Ibase_A = 1#
            End If
            
            ' prúd v ampéroch
            Iabs_A = Iabs_pu * Ibase_A
            
            ' prepočet výkonov do MW / MVAr
            Pij_MW = Pij_pu * SBase_MVA
            Qij_MVAr = Qij_pu * SBase_MVA
            
            ' činné straty do kW
            Ploss_kW = Ploss_pu * SBase_MVA * 1000#
            
            ' zápis do riadku vetvy (riadky 3..) - zaokrúhlenie na 2 des. miesta
            ws.Cells(2 + k, 17).Value = Round(Iabs_A, 2)
            ws.Cells(2 + k, 18).Value = Round(dU_percent, 2)
            ws.Cells(2 + k, 19).Value = Round(Pij_MW, 2)
            ws.Cells(2 + k, 20).Value = Round(Qij_MVAr, 2)
            ws.Cells(2 + k, 21).Value = Round(Ploss_kW, 2)
        End If
    Next k
End Sub

' Načítanie dát vedení z listu "vedenia"
' Skutočné R, X [ohm] -> prepočet do p.u. na danú základňu
'
' Formát "vedenia":
'   riadok 2: hlavička
'   riadky 3.. : dáta
'   A: (voľné)
'   B: Názov vedenia
'   C: uzol-od
'   D: uzol-do
'   E: Status (0/1/2)
'   N: R [ohm] (14)
'   O: X [ohm] (15)
'
' SBase_MVA
Public Sub LoadBranchData( _
    ByRef nBranches As Long, _
    ByRef BranchName() As String, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef R() As Double, _
    ByRef X() As Double, _
    ByRef BranchStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef Bshunt() As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idx As Long
    Dim Zbase_ohm As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase1 As Double
    Dim stVal As Variant
    Dim data As Variant
    Dim B_S As Double

    Set ws = ThisWorkbook.Worksheets("vedenia")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nBranches = 0
        Exit Sub
    End If

    nBranches = lastRow - 2

    ReDim BranchName(1 To nBranches)
    ReDim FromBus(1 To nBranches)
    ReDim ToBus(1 To nBranches)
    ReDim R(1 To nBranches)
    ReDim X(1 To nBranches)
    ReDim Bshunt(1 To nBranches)
    ReDim BranchStatus(1 To nBranches)

    ' bulk read B..P (1=B name, 2=C from, 3=D to, 4=E status, 13=N R, 14=O X, 15=P B)
    data = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 16)).Value

    For i = 1 To nBranches
        BranchName(i) = CStr(data(i, 1))                    ' B: Názov
        fromName = CStr(data(i, 2))                         ' C: Uzol od
        toName = CStr(data(i, 3))                           ' D: Uzol do

        ' Status E (5) -> col 4
        stVal = data(i, 4)
        If IsEmpty(stVal) Or Trim(CStr(stVal)) = "" Then
            BranchStatus(i) = 1 ' Default ON
        ElseIf IsNumeric(stVal) Then
            BranchStatus(i) = CInt(stVal)
        Else
            BranchStatus(i) = 0
        End If

        idx = GetBusIndexD(fromName, busDict)
        If idx = 0 Then
            Err.Raise vbObjectError + 2, , "Uzol '" & fromName & "' v liste 'vedenia', riadok " & (i + 2) & " neexistuje v liste 'uzly'."
        End If
        FromBus(i) = idx

        idx = GetBusIndexD(toName, busDict)
        If idx = 0 Then
            Err.Raise vbObjectError + 3, , "Uzol '" & toName & "' v liste 'vedenia', riadok " & (i + 2) & " neexistuje v liste 'uzly'."
        End If
        ToBus(i) = idx

        ' Určenie Zbase pre vedenie (podľa "from" uzla; idx je teraz toName, ale Ubase pôvodne
        ' brané z posledne nastaveného idx -> bola to logika pôvodného kódu)
        Ubase1 = BusBaseKV(idx)

        If SBase_MVA <> 0# Then
            Zbase_ohm = (Ubase1 * Ubase1) / SBase_MVA
        Else
            Zbase_ohm = 1#
        End If

        ' načítanie skutočnej impedancie - R z N(14)->col 13, X z O(15)->14, B z P(16)->15
        R_ohm = ParseDouble(data(i, 13))
        X_ohm = ParseDouble(data(i, 14))
        B_S = ParseDouble(data(i, 15))

        ' prepočet do p.u.
        If Zbase_ohm <> 0# Then
            R(i) = R_ohm / Zbase_ohm
            X(i) = X_ohm / Zbase_ohm
            Bshunt(i) = B_S * Zbase_ohm
        Else
            R(i) = 0#
            X(i) = 0#
            Bshunt(i) = 0#
        End If
    Next i
End Sub

' Výpočet prúdov a tokov výkonu v transformátoroch po NR výpočte
' Analogicky k WriteBranchCurrents, ale zohľadňuje prevod
' Výstupy na list "transformatory" od stĺpca X:
'   X: |I_prim| [A]
'   Y: |I_sec| [A]
'   Z: P_prim [MW] (do trafa z primáru)
'   AA: Q_prim [MVAr]
'   AB: P_sec [MW] (z trafa do sekundáru)
'   AC: Q_sec [MVAr]
'   AD: P_str [kW]
Public Sub WriteTransformerFlows( _
    ByVal nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef TrR() As Double, _
    ByRef TrX() As Double, _
    ByRef TrG() As Double, _
    ByRef TrB() As Double, _
    ByRef TrRatio() As Double, _
    ByRef Vmag() As Double, _
    ByRef Vang() As Double, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double)

    Dim ws As Worksheet
    Dim k As Long
    Dim i As Long, J As Long
    Dim Vi As Complex, Vj As Complex
    Dim Zs As Complex, Ys As Complex
    Dim Ym As Complex
    Dim A As Double
    Dim I_prim As Complex, I_sec As Complex ' Prúdy v p.u.
    Dim S_prim As Complex, S_sec As Complex ' Výkony v p.u.
    Dim Ibase_prim_A As Double, Ibase_sec_A As Double
    Dim Ubase1 As Double, Ubase2 As Double
    
    Dim Iprim_A As Double, Isec_A As Double
    Dim Pprim_MW As Double, Qprim_MVAr As Double
    Dim Psec_MW As Double, Qsec_MVAr As Double
    Dim Ploss_kW As Double
    
    Set ws = ThisWorkbook.Worksheets("transformatory")
    
    ' Hlavičky
    ws.Cells(2, 24).Value = "|I_prim| [A]"
    ws.Cells(2, 25).Value = "|I_sec| [A]"
    ws.Cells(2, 26).Value = "P_prim [MW]"
    ws.Cells(2, 27).Value = "Q_prim [MVAr]"
    ws.Cells(2, 28).Value = "P_sec [MW]"
    ws.Cells(2, 29).Value = "Q_sec [MVAr]"
    ws.Cells(2, 30).Value = "P_str [kW]"
    
    For k = 1 To nTrafo
        i = TrFrom(k) ' Primár (s odbočkou)
        J = TrTo(k)   ' Sekundár
        A = TrRatio(k)
        
        Vi = CFromPolar(Vmag(i), Vang(i) * RAD2DEG)
        Vj = CFromPolar(Vmag(J), Vang(J) * RAD2DEG)
        
        Zs = CCreate(TrR(k), TrX(k))
        Ys = CDiv(CCreate(1#, 0#), Zs)
        Ym = CCreate(TrG(k), TrB(k)) ' Priečna admitancia na primári
        
        ' Prúd do primáru (uzol i):
        ' I_i = (Vi/a^2 + Vi*Ym - Vj/a) * ys ... nie presne takto z modelu
        ' Z modelu PI:
        ' I_i = Vi * (ys/a^2 + Ym) - Vj * (ys/a)
        ' I_j = Vj * (ys) - Vi * (ys/a)
        
        ' Pomocné pre primár:
        ' term1 = Vi * (ys/a^2 + Ym)
        Dim term1 As Complex, term2 As Complex
        Dim Yseries_a2 As Complex
        
        Yseries_a2 = CCreate(Ys.Re / (A * A), Ys.Im / (A * A))
        term1 = CMul(Vi, CAdd(Yseries_a2, Ym))
        
        ' term2 = Vj * (ys/a)
        Dim Yseries_a As Complex
        Yseries_a = CCreate(Ys.Re / A, Ys.Im / A)
        term2 = CMul(Vj, Yseries_a)
        
        I_prim = CSub(term1, term2)
        
        ' Pomocné pre sekundár (tok DO siete z uzla j):
        ' I_j = Vj * ys - Vi * (ys/a)
        term1 = CMul(Vj, Ys)
        term2 = CMul(Vi, Yseries_a)
        I_sec = CSub(term1, term2)
        
        ' Výkony
        S_prim = CMul(Vi, CConj(I_prim))
        S_sec = CMul(Vj, CConj(I_sec))
        
        ' Bázy prúdu pre primár a sekundár
        Ubase1 = BusBaseKV(i)
        Ubase2 = BusBaseKV(J)
        
        If SBase_MVA <> 0# And Ubase1 <> 0# Then
            Ibase_prim_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase1)
        Else
            Ibase_prim_A = 1#
        End If
        
        If SBase_MVA <> 0# And Ubase2 <> 0# Then
            Ibase_sec_A = (SBase_MVA * 1000#) / (Sqr(3#) * Ubase2)
        Else
            Ibase_sec_A = 1#
        End If
        
        ' Absolútne hodnoty prúdov
        Iprim_A = CAbs(I_prim) * Ibase_prim_A
        Isec_A = CAbs(I_sec) * Ibase_sec_A
        
        ' Výkony v MW/MVAr
        Pprim_MW = S_prim.Re * SBase_MVA
        Qprim_MVAr = S_prim.Im * SBase_MVA
        
        ' Pozor: P_sec definujeme ako tok Z trafa DO sekundáru?
        ' Štandardne v tabuľkách tokov sa píše tok z i->j.
        ' Tu máme P_prim (vstup do trafa) a P_sec (výstup z trafa = -S_sec, ak S_sec je injekcia do siete)
        ' Alebo jednoducho zapíšeme injekciu do uzla j (S_sec).
        ' Dohodnime sa: P_sec bude tok smerom OD uzla j DO trafa (čiže S_sec tak ako vyšlo),
        ' alebo tok smerom Z trafa DO uzla j?
        ' Obvykle sa udáva tok na začiatku a na konci vedenia.
        ' Takže P_sec = P_ji = S_sec.Re * SBase_MVA.
        Psec_MW = S_sec.Re * SBase_MVA
        Qsec_MVAr = S_sec.Im * SBase_MVA
        
        ' Straty = S_prim + S_sec
        Ploss_kW = (S_prim.Re + S_sec.Re) * SBase_MVA * 1000#
        
        ' Zápis
        ws.Cells(2 + k, 24).Value = Round(Iprim_A, 2)
        ws.Cells(2 + k, 25).Value = Round(Isec_A, 2)
        ws.Cells(2 + k, 26).Value = Round(Pprim_MW, 2)
        ws.Cells(2 + k, 27).Value = Round(Qprim_MVAr, 2)
        ws.Cells(2 + k, 28).Value = Round(Psec_MW, 2)
        ws.Cells(2 + k, 29).Value = Round(Qsec_MVAr, 2)
        ws.Cells(2 + k, 30).Value = Round(Ploss_kW, 2)
    Next k
End Sub

' Zápis reportu o izolovaných uzloch a vedeniach
Public Sub WriteIsolationReport( _
    ByVal nBuses As Long, _
    ByRef BusNames() As String, _
    ByRef IsBusIsolated() As Boolean, _
    ByVal nBranches As Long, _
    ByRef FromBus() As Long, _
    ByRef ToBus() As Long, _
    ByRef IsBranchIsolated() As Boolean, _
    ByVal nTrafo As Long, _
    ByRef TrFrom() As Long, _
    ByRef TrTo() As Long, _
    ByRef IsTrafoIsolated() As Boolean, _
    ByVal nComp As Long, _
    ByRef CompBus() As Long, _
    ByRef IsCompIsolated() As Boolean)

    Dim ws As Worksheet
    Dim R As Long
    Dim i As Long
    
    Set ws = GetOrCreateSheet("report")
    ws.Cells.Clear
    
    R = 2
    ws.Cells(R, 2).Value = "Izolované uzly"
    ws.Cells(R, 2).Font.Bold = True
    R = R + 1
    
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            ws.Cells(R, 2).Value = BusNames(i)
            R = R + 1
        End If
    Next i
    
    R = R + 1 ' Jeden voľný riadok
    ws.Cells(R, 2).Value = "Izolované vedenia"
    ws.Cells(R, 2).Font.Bold = True
    R = R + 1
    
    For i = 1 To nBranches
        If IsBranchIsolated(i) Then
            ' Vypíšeme "From -> To"
            ws.Cells(R, 2).Value = BusNames(FromBus(i)) & " -> " & BusNames(ToBus(i))
            R = R + 1
        End If
    Next i
    
    ' Môžeme pridať aj izolované transformátory
    For i = 1 To nTrafo
        If IsTrafoIsolated(i) Then
            ws.Cells(R, 2).Value = "Trafo: " & BusNames(TrFrom(i)) & " -> " & BusNames(TrTo(i))
            R = R + 1
        End If
    Next i
    
    ' Izolované kompenzácie
    If nComp > 0 Then
        R = R + 1
        ws.Cells(R, 2).Value = "Izolované kompenzácie"
        ws.Cells(R, 2).Font.Bold = True
        R = R + 1
        
        For i = 1 To nComp
            If IsCompIsolated(i) Then
                ws.Cells(R, 2).Value = "Kompenzácia na uzle: " & BusNames(CompBus(i))
                R = R + 1
            End If
        Next i
    End If
    
    ' Motory VN nie sú explicitne izolované objekty (sú to spotrebiče),
    ' ale môžu byť na izolovaných uzloch.
    
    ' Zápis poznámky do listu "uzly" stĺpec H (8)
    Dim wsUzly As Worksheet
    Set wsUzly = ThisWorkbook.Worksheets("uzly")
    For i = 1 To nBuses
        If IsBusIsolated(i) Then
            ' Ak je izolovaný, do stĺpca H napíšeme "izolovane"
            ' Pozor: WriteFinalVoltagesToUzly prepisuje stĺpec H.
            ' Musíme zabezpečiť poradie volania, alebo to zapísať po WriteFinalVoltages.
            wsUzly.Cells(2 + i, 8).Value = "izolovane"
            wsUzly.Cells(2 + i, 9).Value = "-"
        End If
    Next i
End Sub



' Načítanie dát o spínačoch z hárka "spinace"
Public Sub LoadSwitchData( _
    ByRef nSwitches As Long, _
    ByRef SwitchName() As String, _
    ByRef SwFrom() As Long, _
    ByRef SwTo() As Long, _
    ByRef SwR() As Double, _
    ByRef SwX() As Double, _
    ByRef SwStatus() As Integer, _
    ByRef BusNames() As String, _
    ByRef BusBaseKV() As Double, _
    ByVal SBase_MVA As Double, _
    ByRef busDict As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fromName As String, toName As String
    Dim idx As Long
    Dim Zbase_ohm As Double
    Dim R_ohm As Double, X_ohm As Double
    Dim Ubase1 As Double
    Dim stVal As Variant
    Dim data As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("spinace")
    If ws Is Nothing Then
        nSwitches = 0
        Exit Sub
    End If
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        nSwitches = 0
        Exit Sub
    End If

    nSwitches = lastRow - 2

    ReDim SwitchName(1 To nSwitches)
    ReDim SwFrom(1 To nSwitches)
    ReDim SwTo(1 To nSwitches)
    ReDim SwR(1 To nSwitches)
    ReDim SwX(1 To nSwitches)
    ReDim SwStatus(1 To nSwitches)

    ' bulk read B..H (1=B tag, 2=C from, 3=D to, 4=E status, 6=G R, 7=H X)
    data = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 8)).Value

    For i = 1 To nSwitches
        SwitchName(i) = CStr(data(i, 1))                    ' B: Tag
        fromName = CStr(data(i, 2))                         ' C: Uzol od
        toName = CStr(data(i, 3))                           ' D: Uzol do

        ' E: Status (0/1) -> col 4
        stVal = data(i, 4)
        If IsEmpty(stVal) Or Trim(CStr(stVal)) = "" Then
            SwStatus(i) = 1
        Else
            SwStatus(i) = CInt(stVal)
        End If

        idx = GetBusIndexD(fromName, busDict)
        If idx = 0 Then Err.Raise vbObjectError + 4, , "Uzol '" & fromName & "' v liste 'spinace' neexistuje."
        SwFrom(i) = idx

        idx = GetBusIndexD(toName, busDict)
        If idx = 0 Then Err.Raise vbObjectError + 5, , "Uzol '" & toName & "' v liste 'spinace' neexistuje."
        SwTo(i) = idx

        Ubase1 = BusBaseKV(SwFrom(i))
        If SBase_MVA <> 0# Then
            Zbase_ohm = (Ubase1 * Ubase1) / SBase_MVA
        Else
            Zbase_ohm = 1#
        End If

        ' G: R [ohm] -> col 6, H: X [ohm] -> col 7
        R_ohm = ParseDouble(data(i, 6))
        X_ohm = ParseDouble(data(i, 7))

        If Zbase_ohm <> 0# Then
            SwR(i) = R_ohm / Zbase_ohm
            SwX(i) = X_ohm / Zbase_ohm
        Else
            SwR(i) = 0#
            SwX(i) = 0#
        End If
    Next i
End Sub

' Zápis výsledných prúdov spínačov do hárka "spinace"
Public Sub WriteSwitchResults(ByVal nSwitches As Long, ByRef SwCurrent() As Double)
    Dim ws As Worksheet
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("spinace")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Hlavička N(14)
    ws.Cells(2, 14).Value = "I [A]"
    
    For i = 1 To nSwitches
        ws.Cells(i + 2, 14).Value = Round(SwCurrent(i), 2)
    Next i
End Sub


