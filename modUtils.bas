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
' K3:K8: zoznam napäťových hladín [kV] (prázdne bunky sa ignorujú)
Public Sub GetBaseValues(ByRef SBase_MVA As Double, ByRef VLevels() As Double)
    Dim ws As Worksheet
    Dim i As Long, n As Long
    Dim v As Double
    Dim raw As Variant

    Set ws = ThisWorkbook.Worksheets("data")

    SBase_MVA = ParseDouble(ws.Range("K11").Value)
    If SBase_MVA <= 0# Then
        Err.Raise vbObjectError + 2001, "GetBaseValues", _
            "Bázový výkon v data!K11 je neplatný alebo nezadaný (očakávam kladnú hodnotu v MVA)."
    End If

    n = 0
    ReDim VLevels(1 To 6)
    For i = 3 To 8
        raw = ws.Range("K" & i).Value
        If Not (IsEmpty(raw) Or IsNull(raw)) Then
            If Trim$(CStr(raw)) <> "" Then
                v = ParseDouble(raw)
                If v > 0# Then
                    n = n + 1
                    VLevels(n) = v
                End If
            End If
        End If
    Next i

    If n = 0 Then
        Err.Raise vbObjectError + 2002, "GetBaseValues", _
            "V bunkách data!K3:K8 nie je definovaná žiadna napäťová hladina [kV]."
    End If

    ReDim Preserve VLevels(1 To n)
End Sub

' Určenie bázy pre uzol podľa jeho počiatočného napätia.
' Vyberie hladinu z VLevels s minimálnou relatívnou odchýlkou (|V - lvl| / lvl).
' Pri odchýlke > 25 % alebo pri V_kV <= 0 vyhodí tvrdú chybu (vrátane mena uzla).
Public Function GetBaseVoltageForBus(ByVal V_kV As Double, ByRef VLevels() As Double, ByVal busName As String) As Double
    Const TOL As Double = 0.25
    Dim k As Long
    Dim bestK As Long, bestDiff As Double, diff As Double

    If V_kV <= 0# Then
        Err.Raise vbObjectError + 2003, "GetBaseVoltageForBus", _
            "Uzol '" & busName & "': v stĺpci D (|V| [kV]) nie je zadané kladné napätie, " & _
            "z ktorého by sa dala určiť bázová hladina."
    End If

    bestK = LBound(VLevels)
    bestDiff = Abs(V_kV - VLevels(bestK)) / VLevels(bestK)
    For k = LBound(VLevels) + 1 To UBound(VLevels)
        diff = Abs(V_kV - VLevels(k)) / VLevels(k)
        If diff < bestDiff Then
            bestDiff = diff
            bestK = k
        End If
    Next k

    If bestDiff > TOL Then
        Err.Raise vbObjectError + 2004, "GetBaseVoltageForBus", _
            "Uzol '" & busName & "': napätie " & V_kV & " kV nezodpovedá žiadnej hladine z data!K3:K8 " & _
            "(najbližšia " & VLevels(bestK) & " kV, relatívna odchýlka " & _
            Format$(bestDiff * 100#, "0.0") & " %, povolené max. " & Format$(TOL * 100#, "0") & " %)."
    End If

    GetBaseVoltageForBus = VLevels(bestK)
End Function


