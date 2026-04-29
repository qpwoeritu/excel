Attribute VB_Name = "modMain"
'==========================
' Modul: modMain
' Posledn prava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

' Tlaidlo na tvorbu Y-matice
Public Sub CmdBuildYMatrix()
    On Error GoTo ErrHandler
    
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusBaseKV() As Double
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim BranchName() As String
    Dim R() As Double, X() As Double
    Dim BranchBshunt() As Double
    Dim BranchStatus() As Integer
    
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
    
    Dim Y() As Complex
    Dim G() As Double, B() As Double
    Dim SBase_MVA As Double
    Dim BaseVoltages() As Double
    
    ' bzy
    Call GetBaseValues(SBase_MVA, BaseVoltages)
    
    ' Topolgia - Izolovan asti
    Dim IsBusIsolated() As Boolean
    Dim IsBranchIsolated() As Boolean
    Dim IsTrafoIsolated() As Boolean
    Dim IsReaktorIsolated() As Boolean
    Dim IsDifReaktorIsolated() As Boolean
    Dim IsSwitchIsolated() As Boolean
    Dim IsCompIsolated() As Boolean
    Dim IsMotorIsolated() As Boolean
    Dim isolatedCount As Long

    ' uzly a vetvy v p.u.
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, BaseVoltages)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, BranchBshunt)
    Call LoadTransformerData(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, BusNames, BusBaseKV, SBase_MVA)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA)
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA)
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA)
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA)
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA)
    
    ' Identifikcia izolovanch ast (aj pre samostatn stavbu matice)
    ' Posielame BranchStatus
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
    
    Call BuildYBus(nBuses, nBranches, FromBus, ToBus, R, X, BranchStatus, BranchBshunt, _
                   nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                   nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, _
                   nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                   nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                   nComp, CompBus, CompB, CompStatus, _
                   nMotors, MotorBus, MotorG, MotorB, MotorStatus, _
                   BusNames, IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, _
                   Y, G, B)
    
    MsgBox "Admitann matica bola vytvoren.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Chyba pri tvorbe Y-matice: " & Err.Description, vbCritical
End Sub


' Tlaidlo na spustenie NR load-flow
Public Sub CmdRunNR()
    Call NewtonRaphsonLoadFlow
End Sub

' Tlaidlo na vpoet skratovch prdov
Public Sub CmdCalculateShortCircuit()
    On Error GoTo ErrHandler
    
    Dim nBuses As Long, nBranches As Long
    Dim BusNames() As String
    Dim BusBaseKV() As Double
    Dim BusTypes() As BusType
    Dim Vmag() As Double, Vang() As Double
    Dim Pspec() As Double, Qspec() As Double
    Dim FromBus() As Long, ToBus() As Long
    Dim BranchName() As String
    Dim R() As Double, X() As Double
    Dim BranchBshunt() As Double
    Dim BranchStatus() As Integer
    
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
    
    ' Kompenzcia (len pre natanie do topolgie, vpoet ju ignoruje)
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
    
    Dim SBase_MVA As Double
    Dim BaseVoltages() As Double
    
    Dim Ik_input() As Double
    Dim Ik_result() As Double
    Dim i As Long, ws As Worksheet
    
    ' Topolgia - Izolovan asti
    Dim IsBusIsolated() As Boolean
    Dim IsBranchIsolated() As Boolean
    Dim IsTrafoIsolated() As Boolean
    Dim IsReaktorIsolated() As Boolean
    Dim IsDifReaktorIsolated() As Boolean
    Dim IsCompIsolated() As Boolean
    Dim IsSwitchIsolated() As Boolean

    Dim IsMotorIsolated() As Boolean
    Dim IsGenIsolated() As Boolean
    Dim isolatedCount As Long
    
    ' Natanie dt
    Call GetBaseValues(SBase_MVA, BaseVoltages)
    Call LoadBusData(nBuses, BusNames, BusTypes, Vmag, Vang, Pspec, Qspec, BusBaseKV, SBase_MVA, BaseVoltages)
    Call LoadBranchData(nBranches, BranchName, FromBus, ToBus, R, X, BranchStatus, BusNames, BusBaseKV, SBase_MVA, BranchBshunt)
    Call LoadTransformerData(nTrafo, TrFrom, TrTo, TrR, TrX, TrG, TrB, TrRatio, BusNames, BusBaseKV, SBase_MVA)
    Call LoadReactorData(nReaktory, ReaktorName, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, BusNames, BusBaseKV, SBase_MVA)
    Call LoadDifReactorData(nDifReaktory, DifReaktorName, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, BusNames, BusBaseKV, SBase_MVA)
    Call LoadSwitchData(nSwitches, SwitchName, SwFrom, SwTo, SwR, SwX, SwStatus, BusNames, BusBaseKV, SBase_MVA)
    Call LoadCompData(nComp, CompName, CompBus, CompB, CompStatus, BusNames, BusBaseKV, SBase_MVA)
    Call LoadMotorData(nMotors, MotorName, MotorBus, MotorR, MotorXk, MotorG, MotorB, MotorStatus, BusNames, BusBaseKV, SBase_MVA)
    
    ' Identifikcia izolovanch (aby vpoet nezlyhal na singulrnej matici)
    Call FindIsolatedParts(nBuses, nBranches, FromBus, ToBus, BranchStatus, _
                           nTrafo, TrFrom, TrTo, _
                           nReaktory, ReaktorFrom, ReaktorTo, _
                           nDifReaktory, DifReaktorFrom, DifReaktorTo, _
                           nSwitches, SwFrom, SwTo, SwStatus, _
                           nComp, CompBus, _
                           nMotors, MotorBus, _
                           BusTypes, _
                           IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated, IsCompIsolated, IsMotorIsolated, isolatedCount)
    
    ' Natanie vstupnch skratov zo stpca J (pre Slack)
    ReDim Ik_input(1 To nBuses)
    Set ws = ThisWorkbook.Worksheets("uzly")
    For i = 1 To nBuses
        Ik_input(i) = ParseDouble(ws.Cells(2 + i, 10).Value)
    Next i
    
    ' Vpoet (Kompenzcia tu nie je, lebo sa ignoruje pri skratoch)
    ' Motory s zahrnut
    Call CalculateShortCircuit(nBuses, nBranches, FromBus, ToBus, R, X, BranchStatus, _
                               nSwitches, SwFrom, SwTo, SwR, SwX, SwStatus, _
                               nTrafo, TrFrom, TrTo, TrR, TrX, TrRatio, _
                               nReaktory, ReaktorFrom, ReaktorTo, ReaktorR, ReaktorX, _
                               nDifReaktory, DifReaktorFrom, DifReaktorTo, DifReaktorR, DifReaktorX, _
                               nMotors, MotorBus, MotorXk, MotorStatus, _
                               BusNames, BusTypes, BusBaseKV, Ik_input, Ik_result, SBase_MVA, _
                               IsBusIsolated, IsBranchIsolated, IsTrafoIsolated, IsReaktorIsolated, IsDifReaktorIsolated, IsSwitchIsolated)
                               
    ' Zpis vsledkov
    Call WriteShortCircuitResults(Ik_result, nBuses)
    
    ' Aktualizcia SLD
    Call UpdateSLD
    
    MsgBox "Vpoet skratov ukonen.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Chyba pri vpote skratov: " & Err.Description, vbCritical
End Sub

' Makro pre VBS  kompletn beh: Y-matica + NR
Public Sub RunFullLoadFlow()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Cleanup
    
    Call CmdBuildYMatrix
    Call CmdRunNR
    
    ' Aktualizcia SLD (ak NR zbehne OK, CmdRunNR vol UpdateSLD? Nie, pridme to do CmdRunNR alebo sem)
    ' CmdRunNR je Sub, ktor vol NewtonRaphsonLoadFlow.
    ' NewtonRaphsonLoadFlow je v modNR.
    ' Pridme volanie UpdateSLD na koniec NewtonRaphsonLoadFlow v modNR.
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub




