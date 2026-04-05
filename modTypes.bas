Attribute VB_Name = "modTypes"
'==========================
' Modul: modTypes
' Posledná úprava: 15.02.2026 15:15 (Bratislava)
'==========================
Option Explicit

' Štruktúra pre komplexné èíslo v karteziánskom tvare
Public Type Complex
    Re As Double
    Im As Double
End Type

' Typ uzla
Public Enum BusType
    btSlack = 0
    btPQ = 1
    btPV = 2
End Enum

' Konštanty pre prácu s uhlami
Public Const PI As Double = 3.14159265358979
Public Const DEG2RAD As Double = PI / 180#
Public Const RAD2DEG As Double = 180# / PI


