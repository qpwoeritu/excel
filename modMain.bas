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
    Dim TrFrom() As Long, TrTo() As Long
    Dim TrR() As Double, TrX() As Double
    Dim TrG() As Double, TrB() As Double
    Dim TrRatio() As Double

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
    Dim Ik_input() As Double, Ik_result As Variant
    Dim TrKt() As Double, GenKg() As Double
    Dim Rth As Variant, Xth As Variant
    Dim ip_result As Variant, Ith_result As Variant
    Dim Ib_result As Variant, Ik_steady As Variant
    Dim kappa_result As Variant, RXratio_result As Variant
    Dim scenarioMin As Boolean, Tk_s As Double, f_Hz As Double, scenStr As String
    Dim ws As Worksheet
    Dim i As Long

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

    Call GetBaseValues(SBase_MVA, VLevels)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, VLevels, busDict)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, Bshunt, busDict)
    Call LoadTransformerData(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, TrKt, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA, busDict)
    Call LoadGeneratorData(nGens, GenName, GenTermBus, GenMode, GenStatus, _
                           GenRa, GenXs, GenXd, GenP, GenQref, GenVref, GenEmag, GenPint, GenKg, _
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

        ' Rozšírenie modelu o generátory: fantómové PV uzly pre EMF, injekcia pre PQ
        Call ApplyGeneratorModel(nBuses, BusNames, BusTypes, BusBaseKV, _
                                 Vmag, Vang, Pspec, Qspec, IsBusIsolated, _
                                 nGens, GenName, GenTermBus, GenMode, GenStatus, GenRa, GenXs, _
                                 GenP, GenQref, GenEmag, GenPint, _
                                 nBusNR, BusNamesNR, BusTypesNR, BusBaseKVNR, _
                                 VmagNR, VangNR, PspecNR, QspecNR, IsBusIsolatedNR, _
                                 GenPhantomIdx, nGenBr, GenBrFrom, GenBrTo, GenBrR, GenBrX)

        Call BuildYBus(nBusNR, nBranches, FromBus, ToBus, R, X, BranchStatus, Bshunt, _
                       nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                       nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, _
                       nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                       nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                       nComp, CompBus, CompB, CompStatus, _
                       nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                       nGenBr, GenBrFrom, GenBrTo, GenBrR, GenBrX, _
                       BusNamesNR, IsBusIsolatedNR, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
                       Y, G, B)
    Else
        ' Pre skraty: načítaj Ik_input zo stĺpca J listu uzly (pre slack)
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

        ' Voliteľné parametre skratu z listu data (defaulty zachovajú pôvodné správanie):
        '   K12 = doba trvania skratu T_k [s] (default 1.0)
        '   K13 = scenár "min"/"max" pre napäťový činiteľ c (default "max")
        '   K14 = frekvencia f [Hz] (default 50)
        With ThisWorkbook.Worksheets("data")
            Tk_s = ParseDouble(.Range("K12").Value)
            If Tk_s <= 0# Then Tk_s = 1#
            scenStr = LCase$(Trim$(CStr(.Range("K13").Value)))
            scenarioMin = (scenStr = "min")
            f_Hz = ParseDouble(.Range("K14").Value)
            If f_Hz <= 0# Then f_Hz = 50#
        End With

        Call BuildShortCircuitMatrix(nBuses, nBranches, FromBus, ToBus, R, X, BranchStatus, _
                                     nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                                     nTrafo, TrFrom, TrTo, TrR, TrX, TrRatio, _
                                     nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                                     nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                                     nMotors, MotorBus, MotorXk, MotorStatus, _
                                     nGens, GenTermBus, GenStatus, GenRa, GenXd, _
                                     BusNames, BusTypes, BusBaseKV, Ik_input, SBase_MVA, _
                                     scenarioMin, TrKt, GenKg, _
                                     IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
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
                        nBranches, FromBus, ToBus, R, X, BranchStatus, Bshunt, _
                        nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                        nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, _
                        nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                        nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                        nComp, CompBus, CompB, CompStatus, _
                        nMotors, MotorBus, MotorR, MotorG, MotorB, MotorStatus, _
                        IsBusIsolatedNR, _
                        wsIdx.Range("I6"))

        ' Výsledky generátorov (δ, Q_gen, I, Ploss) do listu "generatory"
        Call WriteGeneratorResults(nGens, GenName, GenTermBus, GenMode, GenStatus, GenRa, GenXs, _
                                   GenP, GenQref, GenPhantomIdx, VmagNR, VangNR, BusBaseKV, SBase_MVA)

        Call WritePhaseTime(wsIdx.Range("J5"), PhaseElapsed)
        Call EndPhaseTimer
        Call SetPhase(phaseCell, psDone)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False
    Else
        Set phaseCell = wsIdx.Range("I7")
        Call SetPhase(phaseCell, psRunning)
        Application.ScreenUpdating = True: DoEvents: Application.ScreenUpdating = False

        Call BeginPhaseTimer(wsIdx.Range("J7"))

        Call SolveShortCircuit(Ysc, nBuses, BusBaseKV, SBase_MVA, scenarioMin, IsBusIsolated, Ik_result, Rth, Xth)
        Call ComputeSCDerived(nBuses, Ik_result, Rth, Xth, IsBusIsolated, Tk_s, f_Hz, _
                              ip_result, Ith_result, Ib_result, Ik_steady, kappa_result, RXratio_result)
        Call WriteShortCircuitResults(Ik_result, nBuses, ip_result, Ith_result, Ib_result, Ik_steady, kappa_result, RXratio_result)

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
