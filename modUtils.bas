Attribute VB_Name = "modUtils"
'==========================
' Modul: modUtils
' Posledná úprava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

' Univerzálne spracovanie èísel
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

' Vyh¾adanie indexu uzla
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

' Nájde prvý vo¾ný riadok
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

' Naèítanie bázových hodnôt zo sheetu "index"
' B10: S_base [MVA]
' B11: U_base_VN [kV] (napr. 110)
' B12: U_base_NN [kV] (napr. 6.3)
Public Sub GetBaseValues(ByRef SBase_MVA As Double, ByRef UBase_VN As Double, ByRef UBase_NN As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("index")
    SBase_MVA = ParseDouble(ws.Range("B10").Value)
    UBase_VN = ParseDouble(ws.Range("B11").Value)
    UBase_NN = ParseDouble(ws.Range("B12").Value)
    
    If SBase_MVA <= 0# Then SBase_MVA = 100#
    If UBase_VN <= 0# Then UBase_VN = 110#
    If UBase_NN <= 0# Then UBase_NN = 6.3
End Sub

' Urèenie bázy pre uzol pod¾a jeho poèiatoèného napätia
Public Function GetBaseVoltageForBus(ByVal V_start_kV As Double, ByVal UBase_VN As Double, ByVal UBase_NN As Double) As Double
    Dim diffVN As Double, diffNN As Double
    
    If V_start_kV <= 0 Then GetBaseVoltageForBus = UBase_VN: Exit Function ' Default na VN ak nula
    
    diffVN = Abs(V_start_kV - UBase_VN)
    diffNN = Abs(V_start_kV - UBase_NN)
    
    ' Ak je napätie bližšie k NN báze (do 25% rozdielu), použi NN, inak VN
    If diffNN < diffVN And (diffNN / UBase_NN) < 0.5 Then
        GetBaseVoltageForBus = UBase_NN
    Else
        GetBaseVoltageForBus = UBase_VN
    End If
End Function


