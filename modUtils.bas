Attribute VB_Name = "modUtils"
'==========================
' Modul: modUtils
' Posledná úprava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

' Univerzálne spracovanie čísel
Public Function ParseDouble(ByVal v As Variant) As Double
    Dim s As String, decSep As String, otherSep As String
    On Error GoTo FailSafe
    
    If IsEmpty(v) Or IsNull(v) Then ParseDouble = 0#: Exit Function
    If IsNumeric(v) Then ParseDouble = CDbl(v): Exit Function
    
    s = Trim$(CStr(v))
    If s = "" Then ParseDouble = 0#: Exit Function
    
    decSep = Application.International(xlDecimalSeparator)
    otherSep = IIf(decSep = ".", ",", ".")
    s = Replace$(s, " ", "")
    s = Replace$(s, otherSep, decSep)
    ParseDouble = CDbl(s)
    Exit Function
FailSafe:
    ParseDouble = 0#
End Function

' Vyhľadanie indexu uzla
Public Function GetBusIndex(ByVal busName As String, ByRef BusNames() As String) As Long
    Dim i As Long
    For i = LBound(BusNames) To UBound(BusNames)
        If StrComp(Trim$(BusNames(i)), Trim$(busName), vbTextCompare) = 0 Then
            GetBusIndex = i
            Exit Function
        End If
    Next i
    GetBusIndex = 0
End Function

' Nájde prvý voľný riadok
Public Function FirstFreeRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    With ws
        FirstFreeRow = IIf(.Cells(.Rows.Count, col).End(xlUp).Row < 2, 2, .Cells(.Rows.Count, col).End(xlUp).Row + 1)
    End With
End Function

' Získa alebo vytvorí list
Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.name = sheetName
    End If
End Function

' Načítanie bázových hodnôt zo sheetu "data"
' K11: S_base [MVA]
' K3:K8: bázové napäťové hladiny [kV] (prázdne bunky sa ignorujú)
' Poznámka: očakáva sa aspoň jedna platná hodnota > 0
Public Sub GetBaseValues(ByRef SBase_MVA As Double, ByRef BaseVoltages() As Double)
    Dim ws As Worksheet
    Dim i As Long, n As Long
    Dim val As Double

    Set ws = ThisWorkbook.Worksheets("data")
    SBase_MVA = ParseDouble(ws.Range("K11").Value)
    If SBase_MVA <= 0# Then
        Err.Raise vbObjectError + 1001, "GetBaseValues", "Neplatna hodnota S_base v data!K11. Hodnota musi byt vacsia ako 0."
    End If

    ReDim BaseVoltages(1 To 1)
    n = 0
    For i = 3 To 8
        val = ParseDouble(ws.Cells(i, 11).Value) ' K = 11
        If val > 0# Then
            n = n + 1
            ReDim Preserve BaseVoltages(1 To n)
            BaseVoltages(n) = val
        End If
    Next i

    If n = 0 Then
        Err.Raise vbObjectError + 1002, "GetBaseValues", "V rozsahu data!K3:K8 nie je ziadna platna bazova napatova hladina."
    End If
End Sub

Public Function GetBaseVoltageForBus(ByVal V_start_kV As Double, ByRef BaseVoltages() As Double) As Double
    Dim i As Long
    Dim matchCount As Long
    Dim matchedBase As Double
    Dim uBase As Double

    If V_start_kV <= 0# Then
        Err.Raise vbObjectError + 1003, "GetBaseVoltageForBus", "Neplatne napatie uzla (" & CStr(V_start_kV) & " kV). Hodnota musi byt vacsia ako 0."
    End If

    matchCount = 0
    matchedBase = 0#
    For i = LBound(BaseVoltages) To UBound(BaseVoltages)
        uBase = BaseVoltages(i)
        If uBase > 0# Then
            If V_start_kV >= 0.9 * uBase And V_start_kV <= 1.1 * uBase Then
                matchCount = matchCount + 1
                matchedBase = uBase
            End If
        End If
    Next i

    If matchCount = 1 Then
        GetBaseVoltageForBus = matchedBase
    ElseIf matchCount = 0 Then
        Err.Raise vbObjectError + 1004, "GetBaseVoltageForBus", _
                  "Napatie uzla " & Format(V_start_kV, "0.###") & " kV nepatri do ziadnej bazovej hladiny (povoleny interval je 0.9 az 1.1 nasobok)."
    Else
        Err.Raise vbObjectError + 1005, "GetBaseVoltageForBus", _
                  "Napatie uzla " & Format(V_start_kV, "0.###") & " kV patri do viacerych bazovych hladin. Skontrolujte konfiguraciu v data!K3:K8."
    End If
End Function

