Attribute VB_Name = "modReduce"
'==========================
' Modul: modReduce
' Redukcia uzlov pre spínače (bus fusion podľa pandapower) + KCL výpočet
' prúdov fúzovaných spínačov.
'   - zopnutý spínač s |Z| <= prah (data!K16) zlúči svoje uzly do supernodu
'     (union-find; ekvivalent pandapower kritéria z_ohm <= 0)
'   - reprezentant supernodu: slack > PV > uzol s injekciou > najnižší index
'     (poradie priorít podľa pandapower build_bus.py)
'   - výpočty (NR, skraty) bežia na supernodoch, výsledky sa expandujú späť
'     na pôvodné uzly cez mapu BusToNode
'   - prúdy fúzovaných spínačov: KCL rozklad vo vnútri supernodu -
'     strom rezovými súčtami presne, slučky lokálnym riešením podľa
'     impedancií spínačov (s varovaním do hárku report)
'==========================
Option Explicit

' ---------- zdieľaný stav pre KCL prúdy fúzovaných spínačov ----------
Private m_kActive As Boolean
Private m_kNBuses As Long                       ' počet pôvodných uzlov
Private m_kBusToNode() As Long                  ' pôvodný uzol -> supernode
Private m_kBaseKV() As Double                   ' bázy pôvodných uzlov [kV]
Private m_kSBase As Double
Private m_kVm() As Double, m_kVa() As Double    ' napätia supernodov (index = výpočtový uzol, vrátane fantómov)
Private m_kInjRe() As Double, m_kInjIm() As Double ' injekcie do pôvodných uzlov [p.u.]
Private m_kNSw As Long
Private m_kSwFrom() As Long, m_kSwTo() As Long  ' pôvodné konce spínačov
Private m_kSwR() As Double, m_kSwX() As Double  ' p.u.
Private m_kSwFused() As Boolean
Private m_kSwName() As String

'--------------------------------------
' Union-find s kompresiou cesty
'--------------------------------------
Private Function UfFind(ByRef parent() As Long, ByVal x As Long) As Long
    Dim r As Long, p As Long
    r = x
    Do While parent(r) <> r
        r = parent(r)
    Loop
    Do While parent(x) <> r
        p = parent(x): parent(x) = r: x = p
    Loop
    UfFind = r
End Function

' Priorita reprezentanta supernodu (podľa pandapower: slack > PV > aktívny uzol)
Private Function RepScore(ByVal busType As BusType, ByVal P As Double, ByVal Q As Double) As Long
    If busType = btSlack Then
        RepScore = 3
    ElseIf busType = btPV Then
        RepScore = 2
    ElseIf Abs(P) + Abs(Q) > 0# Then
        RepScore = 1
    Else
        RepScore = 0
    End If
End Function

'--------------------------------------
' Zostavenie redukcie: BusToNode mapa + node-level polia.
' Fúzuje sa zopnutý spínač (SwStatus > 0) s |Z| <= fuseThrOhm [ohm].
'--------------------------------------
Public Sub BuildNodeReduction( _
    ByVal nBuses As Long, ByRef BusNames() As String, ByRef BusTypes() As BusType, _
    ByRef BusBaseKV() As Double, ByRef Vmag() As Double, ByRef Vang() As Double, _
    ByRef Pspec() As Double, ByRef Qspec() As Double, ByRef IsBusIsolated() As Boolean, _
    ByVal nSwitches As Long, ByRef SwitchName() As String, ByRef SwFrom() As Long, ByRef SwTo() As Long, _
    ByRef SwR() As Double, ByRef SwX() As Double, ByRef SwStatus() As Integer, _
    ByVal fuseThrOhm As Double, ByVal SBase_MVA As Double, _
    ByRef BusToNode() As Long, ByRef nNodes As Long, _
    ByRef NodeNames() As String, ByRef NodeTypes() As BusType, ByRef NodeBaseKV() As Double, _
    ByRef NodeVmag() As Double, ByRef NodeVang() As Double, _
    ByRef NodePspec() As Double, ByRef NodeQspec() As Double, _
    ByRef IsNodeIsolated() As Boolean, ByRef SwFused() As Boolean)

    Dim parent() As Long
    Dim rootOf() As Long, bestOf() As Long, nodeOfRoot() As Long, repOfNode() As Long
    Dim i As Long, k As Long, r As Long, n As Long, rep As Long
    Dim Zb As Double, Zohm As Double
    Dim nFusedCnt As Long

    ReDim parent(1 To nBuses)
    For i = 1 To nBuses: parent(i) = i: Next i

    If nSwitches > 0 Then
        ReDim SwFused(1 To nSwitches)
    Else
        ReDim SwFused(0 To 0)
    End If

    ' 1) union-find nad fúzovanými spínačmi
    nFusedCnt = 0
    For k = 1 To nSwitches
        SwFused(k) = False
        If SwStatus(k) > 0 Then
            If SBase_MVA > 0# Then
                Zb = (BusBaseKV(SwFrom(k)) * BusBaseKV(SwFrom(k))) / SBase_MVA
            Else
                Zb = 1#
            End If
            Zohm = Sqr(SwR(k) * SwR(k) + SwX(k) * SwX(k)) * Zb
            If Zohm <= fuseThrOhm Then
                SwFused(k) = True
                nFusedCnt = nFusedCnt + 1
                parent(UfFind(parent, SwFrom(k))) = UfFind(parent, SwTo(k))
            End If
        End If
    Next k

    ' 2) korene, výber reprezentanta (najvyššie skóre, pri zhode najnižší index)
    ReDim rootOf(1 To nBuses)
    ReDim bestOf(1 To nBuses)
    For i = 1 To nBuses
        rootOf(i) = UfFind(parent, i)
    Next i
    For i = 1 To nBuses
        r = rootOf(i)
        If bestOf(r) = 0 Then
            bestOf(r) = i
        ElseIf RepScore(BusTypes(i), Pspec(i), Qspec(i)) > _
               RepScore(BusTypes(bestOf(r)), Pspec(bestOf(r)), Qspec(bestOf(r))) Then
            bestOf(r) = i
        End If
    Next i

    ' 3) číslovanie supernodov v poradí prvého výskytu člena
    ReDim nodeOfRoot(1 To nBuses)
    ReDim BusToNode(1 To nBuses)
    nNodes = 0
    For i = 1 To nBuses
        r = rootOf(i)
        If nodeOfRoot(r) = 0 Then
            nNodes = nNodes + 1
            nodeOfRoot(r) = nNodes
        End If
        BusToNode(i) = nodeOfRoot(r)
    Next i
    ReDim repOfNode(1 To nNodes)
    For i = 1 To nBuses
        If repOfNode(BusToNode(i)) = 0 Then repOfNode(BusToNode(i)) = bestOf(rootOf(i))
    Next i

    ' 4) kontroly konzistencie supernodov
    Dim cntSlack() As Long
    ReDim cntSlack(1 To nNodes)
    For i = 1 To nBuses
        n = BusToNode(i): rep = repOfNode(n)
        If Abs(BusBaseKV(i) - BusBaseKV(rep)) > 0.000001 Then
            Err.Raise vbObjectError + 40, "BuildNodeReduction", _
                "Fúzia spínačov spája uzly rôznych napäťových hladín: '" & BusNames(i) & _
                "' (" & BusBaseKV(i) & " kV) a '" & BusNames(rep) & "' (" & BusBaseKV(rep) & " kV)." & _
                " Skontrolujte hárok 'spinace' alebo zvýšte prah v data!K16."
        End If
        If BusTypes(i) = btSlack Then cntSlack(n) = cntSlack(n) + 1
        If (BusTypes(i) = btSlack Or BusTypes(i) = btPV) And _
           (BusTypes(rep) = btSlack Or BusTypes(rep) = btPV) And i <> rep Then
            If Abs(Vmag(i) - Vmag(rep)) > 0.000001 Then
                Err.Raise vbObjectError + 41, "BuildNodeReduction", _
                    "Supernode obsahuje slack/PV uzly s rôznym zadaným napätím: '" & _
                    BusNames(i) & "' a '" & BusNames(rep) & "'."
            End If
        End If
        If IsBusIsolated(i) <> IsBusIsolated(rep) Then
            Call AddCalcWarning("Redukcia: uzly '" & BusNames(i) & "' a '" & BusNames(rep) & _
                "' v jednom supernode majú rôznu izolovanosť - skontrolujte topológiu.")
        End If
    Next i
    For n = 1 To nNodes
        If cntSlack(n) > 1 Then
            Err.Raise vbObjectError + 42, "BuildNodeReduction", _
                "Supernode '" & BusNames(repOfNode(n)) & "' obsahuje viac ako jeden slack uzol."
        End If
    Next n

    ' 5) node-level polia (agregácia injekcií, hodnoty reprezentanta)
    ReDim NodeNames(1 To nNodes): ReDim NodeTypes(1 To nNodes)
    ReDim NodeBaseKV(1 To nNodes): ReDim NodeVmag(1 To nNodes): ReDim NodeVang(1 To nNodes)
    ReDim NodePspec(1 To nNodes): ReDim NodeQspec(1 To nNodes)
    ReDim IsNodeIsolated(1 To nNodes)
    For n = 1 To nNodes
        rep = repOfNode(n)
        NodeNames(n) = BusNames(rep)
        NodeTypes(n) = BusTypes(rep)
        NodeBaseKV(n) = BusBaseKV(rep)
        NodeVmag(n) = Vmag(rep)
        NodeVang(n) = Vang(rep)
        IsNodeIsolated(n) = IsBusIsolated(rep)
    Next n
    For i = 1 To nBuses
        n = BusToNode(i)
        NodePspec(n) = NodePspec(n) + Pspec(i)
        NodeQspec(n) = NodeQspec(n) + Qspec(i)
    Next i
End Sub

'======================================================================
' KCL prúdy fúzovaných spínačov (builder: Begin -> Add* -> SolveAndWrite)
'======================================================================

Public Sub SwitchKclBegin( _
    ByVal nBusesOrig As Long, ByRef BusToNode() As Long, ByRef BusBaseKV_O() As Double, _
    ByVal SBase_MVA As Double, ByRef VmagN() As Double, ByRef VangN() As Double, _
    ByVal nSwitches As Long, ByRef SwitchName() As String, _
    ByRef SwFromO() As Long, ByRef SwToO() As Long, _
    ByRef SwRpu() As Double, ByRef SwXpu() As Double, ByRef SwFused() As Boolean)

    Dim i As Long
    m_kNBuses = nBusesOrig
    m_kBusToNode = BusToNode
    m_kBaseKV = BusBaseKV_O
    m_kSBase = SBase_MVA
    m_kVm = VmagN
    m_kVa = VangN
    m_kNSw = nSwitches
    m_kSwFrom = SwFromO
    m_kSwTo = SwToO
    m_kSwR = SwRpu
    m_kSwX = SwXpu
    m_kSwFused = SwFused
    m_kSwName = SwitchName
    ReDim m_kInjRe(1 To nBusesOrig)
    ReDim m_kInjIm(1 To nBusesOrig)
    m_kActive = True
End Sub

Private Sub KclCheckActive(ByVal caller As String)
    If Not m_kActive Then
        Err.Raise vbObjectError + 43, caller, _
            "Interná chyba: SwitchKclBegin nebol zavolaný pred " & caller & "."
    End If
End Sub

' Napätie pôvodného uzla (cez jeho supernode)
Private Function KVBus(ByVal b As Long) As Complex
    KVBus = CFromPolarRad(m_kVm(m_kBusToNode(b)), m_kVa(m_kBusToNode(b)))
End Function

' Napätie priamo podľa výpočtového indexu (fantómové uzly EMF generátorov)
Private Function KVIdx(ByVal n As Long) As Complex
    KVIdx = CFromPolarRad(m_kVm(n), m_kVa(n))
End Function

Private Sub KclAddInj(ByVal b As Long, ByRef I As Complex)
    m_kInjRe(b) = m_kInjRe(b) + I.Re
    m_kInjIm(b) = m_kInjIm(b) + I.Im
End Sub

' Sériové prvky (vedenia s B/2, nefúzované spínače, reaktory, dif. reaktory).
' Injekcia do uzla = -(prúd tečúci z uzla do prvku). Filter zhodný s BuildYBus/
' BranchContribSeries: rešpektuje status (ak useStatus) aj izolovanosť prvku.
Public Sub SwitchKclAddSeries( _
    ByVal nElems As Long, ByRef FromO() As Long, ByRef ToO() As Long, _
    ByRef Rpu() As Double, ByRef Xpu() As Double, _
    ByRef StatusArr As Variant, ByVal useStatus As Boolean, ByRef ElemIso As Variant, _
    ByRef Bsh As Variant, ByVal hasShunt As Boolean)

    Dim k As Long, f As Long, t As Long
    Dim Vf As Complex, Vt As Complex, Z As Complex, Ys As Complex
    Dim Ife As Complex, Ite As Complex, tmp As Complex
    Dim ok As Boolean

    Call KclCheckActive("SwitchKclAddSeries")

    For k = 1 To nElems
        ok = True
        If useStatus Then
            If StatusArr(k) <= 0 Then ok = False
        End If
        If ok Then
            If Not ElemIso(k) And Not (Rpu(k) = 0 And Xpu(k) = 0) Then
                f = FromO(k): t = ToO(k)
                Vf = KVBus(f): Vt = KVBus(t)
                Z = CCreate(Rpu(k), Xpu(k))
                Ys = CDiv(CCreate(1, 0), Z)
                Ife = CMul(CSub(Vf, Vt), Ys)
                Ite = CMul(CSub(Vt, Vf), Ys)
                If hasShunt Then
                    If Bsh(k) <> 0 Then
                        tmp = CMul(Vf, CCreate(0, CDbl(Bsh(k)) / 2#)): Ife = CAdd(Ife, tmp)
                        tmp = CMul(Vt, CCreate(0, CDbl(Bsh(k)) / 2#)): Ite = CAdd(Ite, tmp)
                    End If
                End If
                tmp = CCreate(-Ife.Re, -Ife.Im): Call KclAddInj(f, tmp)
                tmp = CCreate(-Ite.Re, -Ite.Im): Call KclAddInj(t, tmp)
            End If
        End If
    Next k
End Sub

' Trafá (load-flow model: Ym celé na primári)
Public Sub SwitchKclAddTrafo( _
    ByVal nTrafo As Long, ByRef FromO() As Long, ByRef ToO() As Long, _
    ByRef Rpu() As Double, ByRef Xpu() As Double, _
    ByRef Gpu() As Double, ByRef Bpu() As Double, ByRef Ratio() As Double, _
    ByRef ElemIso As Variant)

    Dim k As Long, f As Long, t As Long
    Dim A As Double
    Dim Vf As Complex, Vt As Complex, Z As Complex, Ys As Complex, Ym As Complex
    Dim Ife As Complex, Ite As Complex, tmp As Complex

    Call KclCheckActive("SwitchKclAddTrafo")

    For k = 1 To nTrafo
        If Not ElemIso(k) And Not (Rpu(k) = 0 And Xpu(k) = 0) Then
            f = FromO(k): t = ToO(k): A = Ratio(k)
            Vf = KVBus(f): Vt = KVBus(t)
            Z = CCreate(Rpu(k), Xpu(k))
            Ys = CDiv(CCreate(1, 0), Z)
            Ym = CCreate(Gpu(k), Bpu(k))
            ' I_f = Vf·(ys/a² + Ym) - Vt·ys/a ;  I_t = Vt·ys - Vf·ys/a
            Ife = CSub(CMul(Vf, CAdd(CCreate(Ys.Re / (A * A), Ys.Im / (A * A)), Ym)), _
                       CMul(Vt, CCreate(Ys.Re / A, Ys.Im / A)))
            Ite = CSub(CMul(Vt, Ys), CMul(Vf, CCreate(Ys.Re / A, Ys.Im / A)))
            tmp = CCreate(-Ife.Re, -Ife.Im): Call KclAddInj(f, tmp)
            tmp = CCreate(-Ite.Re, -Ite.Im): Call KclAddInj(t, tmp)
        End If
    Next k
End Sub

' Priečne prvky (kompenzácie: G = 0, motory: G + jB)
Public Sub SwitchKclAddShunt( _
    ByVal nElems As Long, ByRef BusO() As Long, _
    ByRef Gpu() As Double, ByRef Bpu() As Double, ByRef StatusArr() As Integer)

    Dim k As Long, b As Long
    Dim V As Complex, I As Complex, tmp As Complex

    Call KclCheckActive("SwitchKclAddShunt")

    For k = 1 To nElems
        If StatusArr(k) = 1 Then
            b = BusO(k)
            V = KVBus(b)
            I = CMul(V, CCreate(Gpu(k), Bpu(k)))
            tmp = CCreate(-I.Re, -I.Im): Call KclAddInj(b, tmp)
        End If
    Next k
End Sub

' Zadané injekcie P/Q pôvodných uzlov (odbery záporné): I = conj(S/V)
Public Sub SwitchKclAddLoads(ByRef PspecO() As Double, ByRef QspecO() As Double)
    Dim b As Long
    Dim V As Complex, I As Complex

    Call KclCheckActive("SwitchKclAddLoads")

    For b = 1 To m_kNBuses
        If PspecO(b) <> 0# Or QspecO(b) <> 0# Then
            V = KVBus(b)
            If CAbs(V) > 0.000001 Then
                I = CConj(CDiv(CCreate(PspecO(b), QspecO(b)), V))
                Call KclAddInj(b, I)
            End If
        End If
    Next b
End Sub

' Generátory: PQ režim ako injekcia S, EMF režim prúdom fantóm -> svorka
Public Sub SwitchKclAddGens( _
    ByVal nGens As Long, ByRef TermO() As Long, ByRef GenMode() As Integer, _
    ByRef GenStatus() As Integer, ByRef GenRa() As Double, ByRef GenXs() As Double, _
    ByRef GenP() As Double, ByRef GenQref() As Double, ByRef GenPhantomIdx() As Long)

    Dim k As Long, b As Long
    Dim V As Complex, E As Complex, Z As Complex, I As Complex

    Call KclCheckActive("SwitchKclAddGens")

    For k = 1 To nGens
        If GenStatus(k) = 1 Then
            b = TermO(k)
            V = KVBus(b)
            If GenMode(k) = 1 Then
                If GenPhantomIdx(k) > 0 And Not (GenRa(k) = 0 And GenXs(k) = 0) Then
                    E = KVIdx(GenPhantomIdx(k))
                    Z = CCreate(GenRa(k), GenXs(k))
                    I = CDiv(CSub(E, V), Z)
                    Call KclAddInj(b, I)
                End If
            Else
                If CAbs(V) > 0.000001 Then
                    I = CConj(CDiv(CCreate(GenP(k), GenQref(k)), V))
                    Call KclAddInj(b, I)
                End If
            End If
        End If
    Next k
End Sub

'--------------------------------------
' Rozklad injekcií na prúdy fúzovaných spínačov a zápis do spinace!N.
' Strom: rezové súčty (odlupovanie listov). Slučky: lokálne riešenie
' potenciálov podľa impedancií spínačov + varovanie.
'--------------------------------------
Public Sub SwitchKclSolveAndWrite()
    Dim ws As Worksheet
    Dim maxNode As Long, b As Long, k As Long, n As Long, i As Long, J As Long
    Dim ecnt() As Long

    Call KclCheckActive("SwitchKclSolveAndWrite")

    maxNode = 0
    For b = 1 To m_kNBuses
        If m_kBusToNode(b) > maxNode Then maxNode = m_kBusToNode(b)
    Next b
    If maxNode < 1 Then GoTo Cleanup

    ReDim ecnt(1 To maxNode)
    For k = 1 To m_kNSw
        If m_kSwFused(k) Then ecnt(m_kBusToNode(m_kSwFrom(k))) = ecnt(m_kBusToNode(m_kSwFrom(k))) + 1
    Next k

    Set ws = ThisWorkbook.Worksheets("spinace")

    ' lokálny index pôvodného uzla v aktuálnej skupine (0 = mimo skupiny)
    Dim loc() As Long
    ReDim loc(1 To m_kNBuses)

    For n = 1 To maxNode
        If ecnt(n) > 0 Then
            ' --- zostav skupinu: členy (vrcholy) a fúzované hrany ---
            Dim mV As Long, mE As Long
            Dim vb() As Long, eIdx() As Long
            mV = 0
            For b = 1 To m_kNBuses
                If m_kBusToNode(b) = n Then mV = mV + 1
            Next b
            ReDim vb(1 To mV)
            i = 0
            For b = 1 To m_kNBuses
                If m_kBusToNode(b) = n Then
                    i = i + 1: vb(i) = b: loc(b) = i
                End If
            Next b
            ReDim eIdx(1 To ecnt(n))
            mE = 0
            For k = 1 To m_kNSw
                If m_kSwFused(k) Then
                    If m_kBusToNode(m_kSwFrom(k)) = n Then
                        mE = mE + 1: eIdx(mE) = k
                    End If
                End If
            Next k

            ' --- KCL kontrola: súčet injekcií supernodu má byť ~0 ---
            Dim sRe As Double, sIm As Double, mx As Double
            sRe = 0#: sIm = 0#: mx = 0#
            For i = 1 To mV
                sRe = sRe + m_kInjRe(vb(i)): sIm = sIm + m_kInjIm(vb(i))
                If Abs(m_kInjRe(vb(i))) + Abs(m_kInjIm(vb(i))) > mx Then mx = Abs(m_kInjRe(vb(i))) + Abs(m_kInjIm(vb(i)))
            Next i
            If Sqr(sRe * sRe + sIm * sIm) > 0.000001 + 0.001 * mx Then
                Call AddCalcWarning("KCL kontrola supernodu s uzlom '" & CStr(vb(1)) & _
                    "': suma injekcii = " & Format(Sqr(sRe * sRe + sIm * sIm), "0.000000") & _
                    " p.u. - prúdy spínačov môžu byť nepresné.")
            End If

            ' --- výpočet prúdov hrán ---
            Dim Icur() As Complex
            ReDim Icur(1 To mE)
            If mE = mV - 1 Then
                Call KclSolveTree(mV, mE, vb, eIdx, loc, Icur)
            Else
                Call KclSolveLoop(mV, mE, vb, eIdx, loc, Icur)
                Dim nm As String
                nm = ""
                For J = 1 To mE
                    If J > 1 Then nm = nm & ", "
                    nm = nm & m_kSwName(eIdx(J))
                Next J
                Call AddCalcWarning("Supernode obsahuje slučku spínačov - prúdy rozdelené podľa impedancií (spínače: " & nm & ").")
            End If

            ' --- zápis [A] do spinace!N ---
            Dim Ubase As Double, Ibase_A As Double
            For J = 1 To mE
                k = eIdx(J)
                Ubase = m_kBaseKV(m_kSwFrom(k))
                If Ubase <> 0# Then
                    Ibase_A = (m_kSBase * 1000#) / (Sqr(3) * Ubase)
                Else
                    Ibase_A = 0#
                End If
                ws.Cells(k + 2, 14).Value = Round(CAbs(Icur(J)) * Ibase_A, 2)
            Next J

            ' vyčisti lokálne indexy skupiny
            For i = 1 To mV: loc(vb(i)) = 0: Next i
        End If
    Next n

Cleanup:
    m_kActive = False
End Sub

' Strom: odlupovanie listov - prúd hrany listu = injekcia nahromadená v liste.
Private Sub KclSolveTree(ByVal mV As Long, ByVal mE As Long, _
    ByRef vb() As Long, ByRef eIdx() As Long, ByRef loc() As Long, ByRef Icur() As Complex)

    Dim deg() As Long, accRe() As Double, accIm() As Double, done() As Boolean
    Dim i As Long, J As Long, k As Long, lf As Long, oth As Long
    Dim remaining As Long, found As Boolean

    ReDim deg(1 To mV): ReDim accRe(1 To mV): ReDim accIm(1 To mV): ReDim done(1 To mE)
    For i = 1 To mV
        accRe(i) = m_kInjRe(vb(i)): accIm(i) = m_kInjIm(vb(i))
    Next i
    For J = 1 To mE
        k = eIdx(J)
        deg(loc(m_kSwFrom(k))) = deg(loc(m_kSwFrom(k))) + 1
        deg(loc(m_kSwTo(k))) = deg(loc(m_kSwTo(k))) + 1
    Next J

    remaining = mE
    Do While remaining > 0
        found = False
        For i = 1 To mV
            If deg(i) = 1 Then
                ' nájdi jedinú nevyriešenú hranu listu i
                For J = 1 To mE
                    If Not done(J) Then
                        k = eIdx(J)
                        lf = 0
                        If loc(m_kSwFrom(k)) = i Then lf = i: oth = loc(m_kSwTo(k))
                        If loc(m_kSwTo(k)) = i Then lf = i: oth = loc(m_kSwFrom(k))
                        If lf > 0 Then
                            ' prúd hrany = injekcia odtekajúca z listu (znamienko podľa orientácie netreba - píše sa |I|)
                            Icur(J) = CCreate(accRe(i), accIm(i))
                            accRe(oth) = accRe(oth) + accRe(i)
                            accIm(oth) = accIm(oth) + accIm(i)
                            done(J) = True
                            deg(i) = 0
                            deg(oth) = deg(oth) - 1
                            remaining = remaining - 1
                            found = True
                            Exit For
                        End If
                    End If
                Next J
                If found Then Exit For
            End If
        Next i
        If Not found Then Exit Do   ' nemalo by nastať pri strome
    Loop
End Sub

' Slučky: potenciály vrcholov z admitancií spínačov (referencia = vrchol 1),
' Y·u = inj, prúd hrany = (u_f - u_t)·y.
Private Sub KclSolveLoop(ByVal mV As Long, ByVal mE As Long, _
    ByRef vb() As Long, ByRef eIdx() As Long, ByRef loc() As Long, ByRef Icur() As Complex)

    Dim Y() As Complex, rhs() As Complex, u() As Complex
    Dim i As Long, J As Long, k As Long, f As Long, t As Long
    Dim ye As Complex, Z As Complex

    If mV < 2 Then Exit Sub
    ReDim Y(1 To mV - 1, 1 To mV - 1)
    ReDim rhs(1 To mV - 1)
    For i = 1 To mV - 1
        rhs(i) = CCreate(m_kInjRe(vb(i + 1)), m_kInjIm(vb(i + 1)))
        For J = 1 To mV - 1: Y(i, J) = CCreate(0, 0): Next J
    Next i

    For J = 1 To mE
        k = eIdx(J)
        f = loc(m_kSwFrom(k)): t = loc(m_kSwTo(k))
        Z = CCreate(m_kSwR(k), m_kSwX(k))
        If Z.Re = 0 And Z.Im = 0 Then Z = CCreate(0.000000001, 0)
        ye = CDiv(CCreate(1, 0), Z)
        If f > 1 Then Y(f - 1, f - 1) = CAdd(Y(f - 1, f - 1), ye)
        If t > 1 Then Y(t - 1, t - 1) = CAdd(Y(t - 1, t - 1), ye)
        If f > 1 And t > 1 Then
            Y(f - 1, t - 1) = CSub(Y(f - 1, t - 1), ye)
            Y(t - 1, f - 1) = CSub(Y(t - 1, f - 1), ye)
        End If
    Next J

    Dim x() As Complex
    Call KclGaussSolve(Y, rhs, x, mV - 1)

    ReDim u(1 To mV)
    u(1) = CCreate(0, 0)
    For i = 2 To mV: u(i) = x(i - 1): Next i

    For J = 1 To mE
        k = eIdx(J)
        f = loc(m_kSwFrom(k)): t = loc(m_kSwTo(k))
        Z = CCreate(m_kSwR(k), m_kSwX(k))
        If Z.Re = 0 And Z.Im = 0 Then Z = CCreate(0.000000001, 0)
        Icur(J) = CDiv(CSub(u(f), u(t)), Z)
    Next J
End Sub

' Malé komplexné Gaussovo riešenie A·x = b s čiastočným pivotovaním (|.|²)
Private Sub KclGaussSolve(ByRef A() As Complex, ByRef b() As Complex, ByRef x() As Complex, ByVal n As Long)
    Dim i As Long, J As Long, k As Long, piv As Long
    Dim mag As Double, best As Double
    Dim fac As Complex, tmp As Complex

    If n < 1 Then Exit Sub
    ReDim x(1 To n)

    For k = 1 To n
        best = -1#: piv = k
        For i = k To n
            mag = A(i, k).Re * A(i, k).Re + A(i, k).Im * A(i, k).Im
            If mag > best Then best = mag: piv = i
        Next i
        If best < 1E-30 Then
            Err.Raise vbObjectError + 44, "KclGaussSolve", _
                "Lokálna sústava spínačov je singulárna - prúdy slučky sa nedajú určiť."
        End If
        If piv <> k Then
            For J = k To n
                tmp = A(k, J): A(k, J) = A(piv, J): A(piv, J) = tmp
            Next J
            tmp = b(k): b(k) = b(piv): b(piv) = tmp
        End If
        For i = k + 1 To n
            fac = CDiv(A(i, k), A(k, k))
            For J = k To n
                A(i, J) = CSub(A(i, J), CMul(fac, A(k, J)))
            Next J
            b(i) = CSub(b(i), CMul(fac, b(k)))
        Next i
    Next k

    For i = n To 1 Step -1
        tmp = b(i)
        For J = i + 1 To n
            tmp = CSub(tmp, CMul(A(i, J), x(J)))
        Next J
        x(i) = CDiv(tmp, A(i, i))
    Next i
End Sub
