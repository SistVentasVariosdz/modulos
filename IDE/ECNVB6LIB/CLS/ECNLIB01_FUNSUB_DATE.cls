VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB01_FUNSUB_DATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NumeroSemana(ByVal fecha As Date) As Integer
    Dim sFecha As String
    sFecha = "01/" & Month(fecha) & "/" & Year(fecha)
    NumeroSemana = DatePart("ww", fecha)
End Function

Function UltimoDiaDelMes(ByVal sMes As String, ByVal sAño As String) As Byte
    ReDim DiasMes(12) As Byte
    
    DiasMes(1) = 31
    DiasMes(2) = 28
    DiasMes(3) = 31
    DiasMes(4) = 30
    DiasMes(5) = 31
    DiasMes(6) = 30
    DiasMes(7) = 31
    DiasMes(8) = 31
    DiasMes(9) = 30
    DiasMes(10) = 31
    DiasMes(11) = 30
    DiasMes(12) = 31
    
    If (Val(sAño) Mod 4) = 0 Then DiasMes(2) = 29
    
    UltimoDiaDelMes = DiasMes(Val(sMes))
    
    Erase DiasMes
End Function
