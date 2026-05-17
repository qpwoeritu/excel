Attribute VB_Name = "modProgress"
'==========================
' Modul: modProgress
' Stavový panel index!I3:J8 + heartbeat časovača pre dlhobežiace slučky.
'==========================
Option Explicit

' Stavy fáz (text + farba pozadia)
Public Enum PhaseState
    psNotStarted = 0    ' "nezačaté"  – oranžová
    psRunning = 1       ' "prebieha"  – modrá
    psDone = 2          ' "dokončené" – zelená
    psError = 3         ' "chyba"     – červená
    psDisabled = 4      ' "—"         – sivá (fáza sa pri danom móde nepoužíva)
End Enum

' Texty stavov
Public Const PHASE_TEXT_NOT_STARTED As String = "nezačaté"
Public Const PHASE_TEXT_RUNNING As String = "prebieha"
Public Const PHASE_TEXT_DONE As String = "dokončené"
Public Const PHASE_TEXT_ERROR As String = "chyba"
Public Const PHASE_TEXT_DISABLED As String = "—"

' Module-level stav heartbeatu (zdielaný medzi NR a skratovou inverziou)
Private m_yieldTimerCell As Range
Private m_yieldStartTime As Double
Private m_yieldLastTime As Double

'--------------------------------------
' Nastaví stav fázy v zadanej bunke (text + farba pozadia + farba písma).
'--------------------------------------
Public Sub SetPhase(ByVal cell As Range, ByVal state As PhaseState)
    Dim txt As String
    Dim bgColor As Long
    Dim fontColor As Long

    Select Case state
        Case psNotStarted
            txt = PHASE_TEXT_NOT_STARTED
            bgColor = RGB(255, 192, 0)      ' oranžová
            fontColor = vbBlack
        Case psRunning
            txt = PHASE_TEXT_RUNNING
            bgColor = RGB(0, 112, 192)      ' modrá
            fontColor = vbWhite
        Case psDone
            txt = PHASE_TEXT_DONE
            bgColor = RGB(0, 176, 80)       ' zelená
            fontColor = vbWhite
        Case psError
            txt = PHASE_TEXT_ERROR
            bgColor = RGB(255, 0, 0)        ' červená
            fontColor = vbWhite
        Case psDisabled
            txt = PHASE_TEXT_DISABLED
            bgColor = RGB(191, 191, 191)    ' sivá
            fontColor = vbBlack
        Case Else
            Exit Sub
    End Select

    cell.Value = txt
    cell.Interior.Color = bgColor
    cell.Font.Color = fontColor
End Sub

'--------------------------------------
' Zápis trvania fázy v sekundách s jedným desatinným miestom.
'--------------------------------------
Public Sub WritePhaseTime(ByVal cell As Range, ByVal seconds As Double)
    cell.NumberFormat = "0.0"
    cell.Value = seconds
End Sub

'--------------------------------------
' Zápis aktuálneho čísla iterácie (celé číslo).
'--------------------------------------
Public Sub WritePhaseIter(ByVal cell As Range, ByVal iter As Long)
    cell.NumberFormat = "0"
    cell.Value = iter
End Sub

'--------------------------------------
' Vyčistí celý stavový panel I3:J8 – obsah, podfarbenie aj farbu písma.
'--------------------------------------
Public Sub ClearPhasePanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("index")
    With ws.Range("I3:J8")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
End Sub

'--------------------------------------
' Označí zadané bunky ako "neaktívne" pri zvolenom móde výpočtu
' (sivé pozadie, znak "—" v textových bunkách, prázdne časové bunky).
' Vstup: súvislý rozsah (napr. "I5:J6" alebo "I7:J7").
'--------------------------------------
Public Sub DisablePhaseCells(ByVal cells As Range)
    Dim c As Range
    cells.ClearContents
    cells.Interior.Color = RGB(191, 191, 191)
    cells.Font.Color = vbBlack
    ' Do textových (I) buniek vložíme "—", číselné (J) ostávajú prázdne.
    For Each c In cells.Cells
        If c.Column = 9 Then ' stĺpec I
            c.Value = PHASE_TEXT_DISABLED
        End If
    Next c
End Sub

'--------------------------------------
' Štart heartbeat časovača pre dlhobežiacu fázu.
' Bunka sa nastaví na formát "0.0" a hodnotu 0.
'--------------------------------------
Public Sub BeginPhaseTimer(ByVal timerCell As Range)
    timerCell.NumberFormat = "0.0"
    timerCell.Value = 0#
    Set m_yieldTimerCell = timerCell
    m_yieldStartTime = Timer
    m_yieldLastTime = Timer
End Sub

'--------------------------------------
' Ukončenie heartbeat časovača – uvoľnenie referencie na bunku.
'--------------------------------------
Public Sub EndPhaseTimer()
    Set m_yieldTimerCell = Nothing
End Sub

'--------------------------------------
' Aktuálny uplynulý čas fázy v sekundách (od BeginPhaseTimer).
'--------------------------------------
Public Function PhaseElapsed() As Double
    PhaseElapsed = Timer - m_yieldStartTime
End Function

'--------------------------------------
' Heartbeat – volá sa z dlhobežiacich slučiek (NR Gauss, skratová inverzia).
' Najviac raz za ~200 ms zapíše uplynulý čas do bunky a urobí DoEvents
' (aby UI neostalo "nereaguje" a aby sa propagoval prípadný stlačený ESC).
'
' Volajúci má vo všeobecnosti vypnuté ScreenUpdating (rýchle zápisy do listov);
' aby sa časová bunka aj tak prekreslila, krátko ho zapneme a po DoEvents zase vypneme.
'--------------------------------------
Public Sub PhaseYield()
    If m_yieldTimerCell Is Nothing Then Exit Sub
    Dim t As Double
    t = Timer
    If t - m_yieldLastTime > 0.2 Then
        On Error Resume Next
        m_yieldTimerCell.Value = t - m_yieldStartTime
        Dim prevScr As Boolean
        prevScr = Application.ScreenUpdating
        If Not prevScr Then Application.ScreenUpdating = True
        DoEvents
        If Not prevScr Then Application.ScreenUpdating = False
        On Error GoTo 0
        m_yieldLastTime = t
    End If
End Sub
