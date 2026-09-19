Option Explicit

'==========================
' Modul: modMain
' Jediný verejný vstupný bod pre tlačidlo: runCALC.
'   - Číta typ výpočtu z index!G5 (1 = load-flow, 2 = skraty)
'   - Riadi všetky fázy výpočtu cez stavový panel index!I3:J8
'   - Aktualizuje stav (text + farba) a čas trvania v reálnom čase
'==========================

' Zámok proti dvojkliku – počas behu runCALC nepustí druhú inštanciu.
Private m_calcBusy As Boolean

'--------------------------------------
' Hlavná procedúra – tlačidlo nech volá toto.
'
' Fázy a ich bunky v paneli (index!I3:J8):
'   I3/J3  Načítanie dát + topológia
'   I4/J4  Tvorba matice (Y pre LF, Ysc pre skraty)
'   I5/J5  Výpočet load-flow (J5 priebežný čas)
'   I6     Aktuálne číslo iterácie load-flow
'   I7/J7  Výpočet skratov (J7 priebežný čas)
'   I8/J8  Zápis do SLD
'
' Bunky neaktívne pre zvolený mód sú zošedené:
'   LF mód  -> I7:J7 sivé
'   SC mód  -> I5:J6 sivé
'--------------------------------------
Public Sub runCALC()
    ' Zámok proti dvojkliku – ak už beží, druhá inštancia sa potichu nepustí.
    If m_calcBusy Then Exit Sub
    m_calcBusy = True

    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevStatusBar As Boolean
    Dim settingsSaved As Boolean
    settingsSaved = False

    Dim wsIdx As Worksheet
    Dim phaseCell As Range
    Dim t0 As Double
    Dim modeNum As Long

    ' Lokálne premenné s dátami siete (sú v scope všetkých fáz)
    Dim SBase_MVA As Double
    Dim VLevels() As Double
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusBaseKV() As Double
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim BranchName() As String
    Dim R() As Double, X() As Double
    Dim Bshunt() As Double
    Dim BranchStatus() As Integer

    Dim nTrafo As Long
    Dim TrName() As String
    Dim TrFrom() As Long, TrTo() As Long
    Dim TrR() As Double, TrX() As Double
    Dim TrG() As Double, TrB() As Double
    Dim TrRatio() As Double
    Dim TrKT() As Double

    Dim nReaktory As Long
    Dim ReaktorName() As String
    Dim ReaktorFrom() As Long, ReaktorTo() As Long
    Dim ReaktorR() As Double, ReaktorX() As Double

    Dim nDifReaktory As Long
    Dim DifReaktorName() As String
    Dim DifReaktorFrom() As Long, DifReaktorTo() As Long
    Dim DifReaktorR() As Double, DifReaktorX() As Double

    Dim nSwitches As Long
    Dim SwitchName() As String
    Dim SwFrom() As Long, SwTo() As Long
    Dim SwR() As Double, SwX() As Double
    Dim SwStatus() As Integer

    Dim nComp As Long
    Dim CompName() As String
    Dim CompBus() As Long
    Dim CompB() As Double
    Dim CompStatus() As Integer

    Dim nMotors As Long
    Dim MotorName() As String
    Dim MotorBus() As Long
    Dim MotorR() As Double
    Dim MotorXk() As Double
    Dim MotorG() As Double
    Dim MotorB() As Double
    Dim MotorStatus() As Integer

    ' Generátory (reálne dáta z listu "generatory")
    Dim nGens As Long
    Dim GenName() As String
    Dim GenTermBus() As Long
    Dim GenMode() As Integer
    Dim GenStatus() As Integer
    Dim GenRa() As Double, GenXs() As Double, GenXd() As Double
    Dim GenP() As Double, GenQref() As Double, GenVref() As Double
    Dim GenEmag() As Double, GenPint() As Double
    Dim GenKG() As Double

    ' Rozšírené polia pre NR (reálne uzly + fantómové PV uzly EMF generátorov)
    Dim nBusNR As Long
    Dim BusNamesNR() As String
    Dim BusTypesNR() As BusType
    Dim BusBaseKVNR() As Double
    Dim VmagNR() As Double, VangNR() As Double
    Dim PspecNR() As Double, QspecNR() As Double
    Dim IsBusIsolatedNR() As Boolean
    Dim GenPhantomIdx() As Long
    Dim nGenBr As Long
    Dim GenBrFrom() As Long, GenBrTo() As Long
    Dim GenBrR() As Double, GenBrX() As Double

    Dim IsBusIsolated() As Boolean
    Dim IsBranchIsolated() As Boolean
    Dim IsTrafoIsolated() As Boolean
    Dim IsReaktorIsolated() As Boolean
    Dim IsDifReaktorIsolated() As Boolean
    Dim IsSwitchIsolated() As Boolean
    Dim IsCompIsolated() As Boolean
    Dim IsMotorIsolated() As Boolean
    Dim isolatedCount As Long

    Dim busDict As Object

    Dim Y() As Complex
    Dim G() As Double, B() As Double

    Dim Ysc() As Complex
    Dim Ik_input() As Double, Ik_inputN() As Double
    Dim Ik_result As Variant, ip_result As Variant
    Dim Ik_resultN As Variant, ip_resultN As Variant
    Dim Z_inv() As Complex
    Dim caseMax As Boolean, faultBusName As String
    Dim IkFeederMax As Double, IkFeederMin As Double, RXfeeder As Double
    Dim faultBusIdx As Long, slackIdx As Long, faultNode As Long
    Dim ws As Worksheet
    Dim i As Long

    ' Redukcia uzlov pre spínače (bus fusion podľa pandapower) - nastavenia a mapovanie.
    ' Pri reduceMode=False je BusToNode identita a nNodes=nBuses (žiadna zmena správania).
    Dim reduceMode As Boolean, fuseThrOhm As Double
    Dim BusToNode() As Long, nNodes As Long
    Dim NodeNames() As String, NodeTypes() As BusType, NodeBaseKV() As Double
    Dim NodeVmag() As Double, NodeVang() As Double
    Dim NodePspec() As Double, NodeQspec() As Double
    Dim IsNodeIsolated() As Boolean
    Dim SwFused() As Boolean

    ' Mapované (výpočtové) konce prvkov: napr. FromBusC(k) = BusToNode(FromBus(k)).
    ' Kŕmia sa nimi BuildYBus/RunNRPhase/BuildShortCircuitMatrix - jadrá výpočtov
    ' samotné sa nemenia, len dostávajú uzly na úrovni supernodov.
    Dim FromBusC() As Long, ToBusC() As Long
    Dim TrFromC() As Long, TrToC() As Long
    Dim ReaktorFromC() As Long, ReaktorToC() As Long
    Dim DifReaktorFromC() As Long, DifReaktorToC() As Long
    Dim SwFromC() As Long, SwToC() As Long
    Dim CompBusC() As Long, MotorBusC() As Long
    Dim GenTermBusC() As Long
    Dim CompZeroG() As Double

    ' Uloženie pôvodných nastavení Excelu (obnovíme v Cleanup aj ErrHandler)
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevStatusBar = Application.DisplayStatusBar
    settingsSaved = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    ' Stlačenie ESC počas behu vyhodí chybu 18 – chytíme ju v ErrHandler
    ' a označíme aktuálnu fázu ako "chyba" (namiesto tichého zastavenia VBA).
    Application.EnableCancelKey = xlErrorHandler

    Set wsIdx = ThisWorkbook.Worksheets("index")

    ' Čítanie typu výpočtu z G5
    Dim modeRaw As Variant
    modeRaw = wsIdx.Range("G5").Value

    ' Prázdne G5 – nebol vybratý typ výpočtu
    If IsEmpty(modeRaw) Or (VarType(modeRaw) = vbString And Len(Trim(CStr(modeRaw))) = 0) Then
        Call RestoreExcelSettings(settingsSaved, prevCalc, prevScreen, prevEvents, prevStatusBar)
        m_calcBusy = False
        MsgBox "Nebol vybratý typ výpočtu.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(modeRaw) Then
        Call RestoreExcelSettings(settingsSaved, prevCalc, prevScreen, prevEvents, prevStatusBar)
        m_calcBusy = False
        MsgBox "Neplatný typ výpočtu (povolené 1 = load-flow, 2 = skraty).", vbExclamation
        Exit Sub
    End If

    modeNum = CLng(modeRaw)
    If modeNum <> 1 And modeNum <> 2 Then
        Call RestoreExcelSettings(settingsSaved, prevCalc, prevScreen, prevEvents, prevStatusBar)
        m_calcBusy = False
        MsgBox "Neplatný typ výpočtu (povolené 1 = load-flow, 2 = skraty).", vbExclamation
        Exit Sub
    End If

    ' Vyčistenie stavového panelu I3:J8 (obsah + podfarbenie + farba písma)
    Call ClearPhasePanel

    ' Počiatočné stavy fáz – "nezačaté" pre relevantné fázy
    Call SetPhase(wsIdx.Range("I3"), psNotStarted)
    Call SetPhase(wsIdx.Range("I4"), psNotStarted)
    Call SetPhase(wsIdx.Range("I8"), psNotStarted)
    If modeNum = 1 Then
        Call SetPhase(wsIdx.Range("I5"), psNotStarted)
        Call DisablePhaseCells(wsIdx.Range("I7:J7"))
    Else
        Call SetPhase(wsIdx.Range("I7"), psNotStarted)
        Call DisablePhaseCells(wsIdx.Range("I5:J6"))
    End If

    ' Vykreslenie počiatočného stavu používateľovi
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    '====================================================================
    ' FÁZA 1: Načítanie dát + kontrola topológie (I3/J3)
    '====================================================================
    Set phaseCell = wsIdx.Range("I3")
    Call SetPhase(phaseCell, psRunning)
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    t0 = Timer

    Call ResetCalcWarnings
    Call GetBaseValues(SBase_MVA, VLevels)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, VLevels, busDict)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, Bshunt, busDict)
    Call LoadTransformerData(nTrafo, TrName, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, TrKT, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadGeneratorData(nGens, GenName, GenTermBus, GenMode, GenStatus, _
                           GenRa, GenXs, GenXd, GenP, GenQref, GenVref, GenEmag, GenPint, GenKG, _
                           BusNames, BusBaseKV, SBase_MVA, busDict)

    Call FindIsolatedParts(nBuses, nBranches, FromBus, ToBus, BranchStatus, _
                           nTrafo, TrFrom, TrTo, _
                           nReaktory, ReaktorFrom, ReaktorTo, _
                           nDifReaktory, DifReaktorFrom, DifReaktorTo, _
                           nSwitches, SwFrom, SwTo, SwStatus, _
                           nComp, CompBus, _
                           nMotors, MotorBus, _
                           BusTypes, _
                           IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, IsCompIsolated, IsMotorIsolated, isolatedCount)

    Call WriteIsolationReport(nBuses, BusNames, IsBusIsolated, _
                              nBranches, FromBus, ToBus, IsBranchIsolated, _
                              nTrafo, TrFrom, TrTo, IsTrafoIsolated, _
                              nComp, CompBus, IsCompIsolated)
    Call FlushCalcWarnings

    Call WritePhaseTime(wsIdx.Range("J3"), Timer - t0)
    Call SetPhase(phaseCell, psDone)
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False

    '====================================================================
    ' FÁZA 2: Tvorba matice (I4/J4)
    '====================================================================
    Set phaseCell = wsIdx.Range("I4")
    Call SetPhase(phaseCell, psRunning)
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    t0 = Timer

    If modeNum = 1 Then
        ' Pre load-flow: vynulujeme izolované uzly v Pspec/Qspec/Vmag (aby mismatch vektor nebol skreslený)
        For i = 1 To nBuses
            If IsBusIsolated(i) Then
                Pspec(i) = 0#
                Qspec(i) = 0#
                Vmag(i) = 0#
            End If
        Next i
    End If

    ' Redukcia uzlov pre spínače (bus fusion podľa pandapower) - spoločná pre load-flow
    ' aj skraty. Pri reduceMode=False sa nastaví identita (BusToNode(i)=i, nNodes=nBuses)
    ' a správanie je zhodné s pôvodnou impedančnou vetvou spínačov.
    Call LoadSwitchReductionSettings(reduceMode, fuseThrOhm)
    If reduceMode Then
        Call BuildNodeReduction(nBuses, BusNames, BusTypes, BusBaseKV, Vmag, Vang, Pspec, Qspec, IsBusIsolated, _
                                nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, _
                                fuseThrOhm, SBase_MVA, _
                                BusToNode, nNodes, NodeNames, NodeTypes, NodeBaseKV, NodeVmag, NodeVang, _
                                NodePspec, NodeQspec, IsNodeIsolated, SwFused)
    Else
        nNodes = nBuses
        ReDim BusToNode(1 To nBuses)
        For i = 1 To nBuses: BusToNode(i) = i: Next i
        NodeNames = BusNames
        NodeTypes = BusTypes
        NodeBaseKV = BusBaseKV
        NodeVmag = Vmag
        NodeVang = Vang
        NodePspec = Pspec
        NodeQspec = Qspec
        IsNodeIsolated = IsBusIsolated
        If nSwitches > 0 Then
            ReDim SwFused(1 To nSwitches)
        Else
            ReDim SwFused(0 To 0)
        End If
    End If

    ' Mapovanie koncov prvkov na výpočtové uzly (pri reduceMode=False je to identita)
    ReDim FromBusC(1 To nBranches): ReDim ToBusC(1 To nBranches)
    For i = 1 To nBranches
        FromBusC(i) = BusToNode(FromBus(i)): ToBusC(i) = BusToNode(ToBus(i))
    Next i
    ReDim TrFromC(1 To nTrafo): ReDim TrToC(1 To nTrafo)
    For i = 1 To nTrafo
        TrFromC(i) = BusToNode(TrFrom(i)): TrToC(i) = BusToNode(TrTo(i))
    Next i
    ReDim ReaktorFromC(1 To nReaktory): ReDim ReaktorToC(1 To nReaktory)
    For i = 1 To nReaktory
        ReaktorFromC(i) = BusToNode(ReaktorFrom(i)): ReaktorToC(i) = BusToNode(ReaktorTo(i))
    Next i
    ReDim DifReaktorFromC(1 To nDifReaktory): ReDim DifReaktorToC(1 To nDifReaktory)
    For i = 1 To nDifReaktory
        DifReaktorFromC(i) = BusToNode(DifReaktorFrom(i)): DifReaktorToC(i) = BusToNode(DifReaktorTo(i))
    Next i
    ReDim SwFromC(1 To nSwitches): ReDim SwToC(1 To nSwitches)
    For i = 1 To nSwitches
        SwFromC(i) = BusToNode(SwFrom(i)): SwToC(i) = BusToNode(SwTo(i))
    Next i
    ReDim CompBusC(1 To nComp)
    For i = 1 To nComp: CompBusC(i) = BusToNode(CompBus(i)): Next i
    ReDim MotorBusC(1 To nMotors)
    For i = 1 To nMotors: MotorBusC(i) = BusToNode(MotorBus(i)): Next i
    ReDim GenTermBusC(1 To nGens)
    For i = 1 To nGens: GenTermBusC(i) = BusToNode(GenTermBus(i)): Next i

    If modeNum = 1 Then
        ' Rozšírenie modelu o generátory: fantómové PV uzly pre EMF, injekcia pre PQ.
        ' Vstupuje na úrovni výpočtových uzlov (supernodov), nie pôvodných zberníc.
        Call ApplyGeneratorModel(nNodes, NodeNames, NodeTypes, NodeBaseKV, _
                                 NodeVmag, NodeVang, NodePspec, NodeQspec, IsNodeIsolated, _
                                 nGens, GenName, GenTermBusC, GenMode, GenStatus, GenRa, GenXs, _
                                 GenP, GenQref, GenEmag, GenPint, _
                                 nBusNR, BusNamesNR, BusTypesNR, BusBaseKVNR, _
                                 VmagNR, VangNR, PspecNR, QspecNR, IsBusIsolatedNR, _
                                 GenPhantomIdx, nGenBr, GenBrFrom, GenBrTo, GenBrR, GenBrX)

        Call BuildYBus(nBusNR, nBranches, FromBusC, ToBusC, R, X, BranchStatus, Bshunt, _
                       nSwitches, SwFromC, SwToC, SwR, SwX, SwStatus, _
                       nTrafo, TrFromC, TrToC, TrR, TrX, TrG, TrB, TrRatio, _
                       nReaktory, ReaktorFromC, ReaktorToC, ReaktorR, ReaktorX, _
                       nDifReaktory, DifReaktorFromC, DifReaktorToC, DifReaktorR, DifReaktorX, _
                       nComp, CompBusC, CompB, CompStatus, _
                       nMotors, MotorBusC, MotorG, MotorB, MotorStatus, _
                       nGenBr, GenBrFrom, GenBrTo, GenBrR, GenBrX, _
                       BusNamesNR, IsBusIsolatedNR, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
                       Y, G, B)
    Else
        ' Pre skraty: nastavenia IEC 60909 (prípad max/min, uzol poruchy, napájač)
        Call LoadShortCircuitSettings(caseMax, faultBusName, IkFeederMax, IkFeederMin, RXfeeder)

        ' Načítaj Ik_input zo stĺpca J listu uzly (spätná kompatibilita pre slack)
        ReDim Ik_input(1 To nBuses)
        Set ws = ThisWorkbook.Worksheets("uzly")
        If nBuses = 1 Then
            Ik_input(1) = ParseDouble(ws.Cells(3, 10).Value)
        Else
            Dim ikArr As Variant
            ikArr = ws.Range(ws.Cells(3, 10), ws.Cells(2 + nBuses, 10)).Value
            For i = 1 To nBuses
                Ik_input(i) = ParseDouble(ikArr(i, 1))
            Next i
        End If

        ' Ik'' napájača z data!K13/K14 podľa prípadu (má prednosť pred uzly!J,
        ' ktorý sa každým behom prepisuje výsledkami)
        slackIdx = 0
        For i = 1 To nBuses
            If BusTypes(i) = btSlack Then slackIdx = i: Exit For
        Next i
        If slackIdx > 0 Then
            If caseMax Then
                If IkFeederMax > 0# Then Ik_input(slackIdx) = IkFeederMax
            Else
                If IkFeederMin > 0# Then
                    Ik_input(slackIdx) = IkFeederMin
                ElseIf IkFeederMax > 0# Then
                    Ik_input(slackIdx) = IkFeederMax
                    Call AddCalcWarning("Prípad min: chýba Ik''min napájača v data!K14 - použitá hodnota Ik''max z data!K13.")
                Else
                    Call AddCalcWarning("Prípad min: chýba Ik''min napájača v data!K14 - použitá hodnota z uzly!J (riadok slacku).")
                End If
            End If
        End If

        ' Ik_input premietnuté na výpočtový uzol slacku (jediný relevantný záznam)
        ReDim Ik_inputN(1 To nNodes)
        If slackIdx > 0 Then Ik_inputN(BusToNode(slackIdx)) = Ik_input(slackIdx)

        Call BuildShortCircuitMatrix(nNodes, nBranches, FromBusC, ToBusC, R, X, BranchStatus, _
                                     nSwitches, SwFromC, SwToC, SwR, SwX, SwStatus, _
                                     nTrafo, TrFromC, TrToC, TrR, TrX, TrRatio, TrKT, _
                                     nReaktory, ReaktorFromC, ReaktorToC, ReaktorR, ReaktorX, _
                                     nDifReaktory, DifReaktorFromC, DifReaktorToC, DifReaktorR, DifReaktorX, _
                                     nMotors, MotorBusC, MotorR, MotorXk, MotorStatus, _
                                     nGens, GenTermBusC, GenStatus, GenRa, GenXd, GenKG, _
                                     NodeNames, NodeTypes, NodeBaseKV, Ik_inputN, SBase_MVA, _
                                     caseMax, RXfeeder, _
                                     IsNodeIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
                                     Ysc)
    End If

    Call WritePhaseTime(wsIdx.Range("J4"), Timer - t0)
    Call SetPhase(phaseCell, psDone)
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False

    '====================================================================
    ' FÁZA 3: Výpočet (LF: I5/J5/I6  |  SC: I7/J7)
    '====================================================================
    If modeNum = 1 Then
        Set phaseCell = wsIdx.Range("I5")
        Call SetPhase(phaseCell, psRunning)
        Call WritePhaseIter(wsIdx.Range("I6"), 0)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False

        ' BeginPhaseTimer nastaví bunku J5 na formát "0.0" a hodnotu 0;
        ' následné PhaseYield z NR Gauss solvera ju potom priebežne aktualizujú.
        ' (PhaseYield si sám krátko zapne ScreenUpdating kvôli prekresleniu.)
        Call BeginPhaseTimer(wsIdx.Range("J5"))

        Call RunNRPhase(SBase_MVA, nBusNR, nBuses, BusNamesNR, BusTypesNR, BusBaseKVNR, _
                        VmagNR, VangNR, PspecNR, QspecNR, G, B, _
                        nBranches, FromBusC, ToBusC, R, X, BranchStatus, Bshunt, _
                        nSwitches, SwFromC, SwToC, SwR, SwX, SwStatus, _
                        nTrafo, TrFromC, TrToC, TrR, TrX, TrG, TrB, TrRatio, _
                        nReaktory, ReaktorFromC, ReaktorToC, ReaktorR, ReaktorX, _
                        nDifReaktory, DifReaktorFromC, DifReaktorToC, DifReaktorR, DifReaktorX, _
                        nComp, CompBusC, CompB, CompStatus, _
                        nMotors, MotorBusC, MotorR, MotorG, MotorB, MotorStatus, _
                        IsBusIsolatedNR, BusToNode, _
                        wsIdx.Range("I6"))

        ' Výsledky generátorov (δ, Q_gen, I, Ploss) do listu "generatory"
        Call WriteGeneratorResults(nGens, GenName, GenTermBusC, GenMode, GenStatus, GenRa, GenXs, _
                                   GenP, GenQref, GenPhantomIdx, VmagNR, VangNR, NodeBaseKV, SBase_MVA)

        ' Prúdy fúzovaných spínačov (KCL rozklad vo vnútri supernodov) - len v režime
        ' redukcia. Nefúzované spínače majú prúd už zapísaný z RunNRPhase (WriteSwitchResults).
        If reduceMode And nSwitches > 0 Then
            Call SwitchKclBegin(nBuses, BusToNode, BusBaseKV, SBase_MVA, VmagNR, VangNR, _
                                nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwFused)
            Call SwitchKclAddSeries(nBranches, FromBus, ToBus, R, X, BranchStatus, True, IsBranchIsolated, Bshunt, True)
            Call SwitchKclAddSeries(nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, True, IsSwitchIsolated, SwStatus, False)
            Call SwitchKclAddSeries(nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, ReaktorR, False, IsReaktorIsolated, ReaktorR, False)
            Call SwitchKclAddSeries(nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, DifReaktorR, False, IsDifReaktorIsolated, DifReaktorR, False)
            Call SwitchKclAddTrafo(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, IsTrafoIsolated)
            ReDim CompZeroG(1 To nComp)
            Call SwitchKclAddShunt(nComp, CompBus, CompZeroG, CompB, CompStatus)
            Call SwitchKclAddShunt(nMotors, MotorBus, MotorG, MotorB, MotorStatus)
            Call SwitchKclAddLoads(Pspec, Qspec)
            Call SwitchKclAddGens(nGens, GenTermBus, GenMode, GenStatus, GenRa, GenXs, GenP, GenQref, GenPhantomIdx)
            Call SwitchKclSolveAndWrite()
        End If

        Call WritePhaseTime(wsIdx.Range("J5"), PhaseElapsed)
        Call EndPhaseTimer
        Call SetPhase(phaseCell, psDone)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    Else
        Set phaseCell = wsIdx.Range("I7")
        Call SetPhase(phaseCell, psRunning)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False

        Call BeginPhaseTimer(wsIdx.Range("J7"))

        Call SolveShortCircuit(Ysc, nNodes, NodeNames, NodeBaseKV, SBase_MVA, caseMax, IsNodeIsolated, Ik_resultN, ip_resultN, Z_inv)

        ' Expanzia výsledkov z výpočtových uzlov (supernodov) na pôvodné uzly
        ReDim Ik_result(1 To nBuses): ReDim ip_result(1 To nBuses)
        For i = 1 To nBuses
            Ik_result(i) = Ik_resultN(BusToNode(i))
            ip_result(i) = ip_resultN(BusToNode(i))
        Next i
        Call WriteShortCircuitResults(Ik_result, ip_result, nBuses)

        ' Vetvové príspevky pre zvolený uzol poruchy (index!G7, voliteľné).
        ' Kvôli limitu VBA (max. 60 parametrov na procedúru) je výpočet rozdelený
        ' na sekvenciu Begin -> prvky -> Finish so zdieľaným stavom v modShortCircuit.
        If Len(faultBusName) > 0 Then
            faultBusIdx = GetBusIndexD(faultBusName, busDict)
            If faultBusIdx = 0 Then
                Err.Raise vbObjectError + 32, , "Uzol poruchy '" & faultBusName & "' z index!G7 neexistuje v liste 'uzly'."
            End If
            If IsBusIsolated(faultBusIdx) Then
                Err.Raise vbObjectError + 33, , "Uzol poruchy '" & faultBusName & "' z index!G7 je izolovaný od slacku."
            End If
            faultNode = BusToNode(faultBusIdx)
            Call BranchContribBegin(faultNode, faultBusName, nNodes, Z_inv, caseMax, SBase_MVA, NodeBaseKV, _
                                    nBranches + nSwitches + nTrafo + nReaktory + nDifReaktory + nMotors + nGens)
            Call BranchContribSeries("vedenie", nBranches, BranchName, FromBusC, ToBusC, R, X, _
                                     BranchStatus, True, IsBranchIsolated, NodeNames, NodeBaseKV)
            Call BranchContribTrafo(nTrafo, TrName, TrFromC, TrToC, TrR, TrX, TrRatio, TrKT, _
                                    IsTrafoIsolated, NodeNames, NodeBaseKV)
            Call BranchContribSeries("spinac", nSwitches, SwitchName, SwFromC, SwToC, SwR, SwX, _
                                     SwStatus, True, IsSwitchIsolated, NodeNames, NodeBaseKV)
            Call BranchContribSeries("reaktor", nReaktory, ReaktorName, ReaktorFromC, ReaktorToC, ReaktorR, ReaktorX, _
                                     IsReaktorIsolated, False, IsReaktorIsolated, NodeNames, NodeBaseKV)
            Call BranchContribSeries("dif.reaktor", nDifReaktory, DifReaktorName, DifReaktorFromC, DifReaktorToC, DifReaktorR, DifReaktorX, _
                                     IsDifReaktorIsolated, False, IsDifReaktorIsolated, NodeNames, NodeBaseKV)
            Call BranchContribMotors(nMotors, MotorName, MotorBusC, MotorR, MotorXk, MotorStatus, _
                                     IsNodeIsolated, NodeNames, NodeBaseKV)
            Call BranchContribGens(nGens, GenName, GenTermBusC, GenStatus, GenRa, GenXd, GenKG, _
                                   IsNodeIsolated, NodeNames, NodeBaseKV)
            Call BranchContribFinish(faultBusName, CDbl(Ik_resultN(faultNode)), CDbl(ip_resultN(faultNode)))
        End If

        ' Varovania z fázy 2/3 (napájač, kappa) - report je už zapísaný z fázy 1
        Call FlushCalcWarnings

        Call WritePhaseTime(wsIdx.Range("J7"), PhaseElapsed)
        Call EndPhaseTimer
        Call SetPhase(phaseCell, psDone)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    End If

    '====================================================================
    ' FÁZA 4: Zápis do SLD (I8/J8)
    '====================================================================
    Set phaseCell = wsIdx.Range("I8")
    Call SetPhase(phaseCell, psRunning)
    Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    t0 = Timer

    Call UpdateSLD

    Call WritePhaseTime(wsIdx.Range("J8"), Timer - t0)
    Call SetPhase(phaseCell, psDone)

    ' Úspešné dokončenie – obnoviť Excel a oznámiť
    Call RestoreExcelSettings(settingsSaved, prevCalc, prevScreen, prevEvents, prevStatusBar)
    m_calcBusy = False
    MsgBox "Výpočet dokončený.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    ' V ErrHandleri nesmieme dopustiť ďalšiu chybu kým ukončíme upratovanie
    On Error Resume Next
    Call EndPhaseTimer
    If Not phaseCell Is Nothing Then
        Call SetPhase(phaseCell, psError)
    End If
    On Error GoTo 0

    Call RestoreExcelSettings(settingsSaved, prevCalc, prevScreen, prevEvents, prevStatusBar)
    m_calcBusy = False

    If errNum = 18 Then
        MsgBox "Výpočet bol zrušený používateľom (ESC).", vbExclamation
    Else
        MsgBox "Chyba pri výpočte: " & errDesc, vbCritical
    End If
End Sub

'--------------------------------------
' Obnoví pôvodné nastavenia Excelu pred opustením runCALC.
' Volá sa z hlavnej vetvy aj z ErrHandlera.
'--------------------------------------
Private Sub RestoreExcelSettings(ByVal saved As Boolean, _
                                 ByVal prevCalc As XlCalculation, _
                                 ByVal prevScreen As Boolean, _
                                 ByVal prevEvents As Boolean, _
                                 ByVal prevStatusBar As Boolean)
    If Not saved Then Exit Sub
    Application.EnableCancelKey = xlInterrupt
    Application.DisplayStatusBar = prevStatusBar
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
End Sub
