VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB01_FUNSUB_STRING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CompletaPalabra
' -------------------
' Completa una cadena con el caracter recibido como parámetro

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CompletaPalabra(ByRef sCodOrigen As String, _
                           ByVal iCodTamaño As Integer, _
                           Optional ByVal sCarLlenar As String = " ", _
                           Optional ByVal blSW_Isquierda As Boolean = True)
    Dim i As Integer
    
    sCodOrigen = Trim(sCodOrigen)
    For i = 1 To (iCodTamaño - Len(sCodOrigen))
        Select Case blSW_Isquierda
            Case True
                sCodOrigen = sCarLlenar & sCodOrigen
            Case False
                sCodOrigen = sCodOrigen & sCarLlenar
        End Select
        sCodOrigen = sCarLlenar & sCodOrigen
    Next
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PadC
' -------------------
' Funciones que trabajan con tipos de datos String
' Función que centra "Cadena" en "Nlong" Posiciones
' Por defecto "Carrelleno" es " "
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function PadC(ByVal cadena As String, ByVal nLong As Integer, Optional ByVal CarRelleno As String) As String
    Dim CadLen As Integer
    Dim LonRe1 As Integer
    Dim LonRe2 As Integer
    
    If CarRelleno = "" Then CarRelleno = " "
    CadLen = Len(LTrim(RTrim(cadena)))
    LonRe1 = Int((nLong - CadLen) / 2)
    LonRe2 = nLong - (CadLen + LonRe1)
    PadC = Replicate(CarRelleno, LonRe1) + LTrim(RTrim(cadena)) + Replicate(CarRelleno, LonRe2)
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PadL
' -------------------
' Función que rellena de "CarRelleno" a la izquierda
' de "Cadena" en "Nlong" posiciones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function PadL(cadena As String, nLong As Integer, Optional CarRelleno As String) As String
    Dim NLongL As Integer
    
    If CarRelleno = "" Then CarRelleno = " "
    cadena = LTrim(RTrim(cadena))
    If nLong >= Len(cadena) Then
        NLongL = nLong - Len(cadena)
        PadL = Replicate(CarRelleno, NLongL) + cadena
    Else
        PadL = Replicate("*", nLong)
    End If
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PadR
' -------------------
' Función que rellena de "CarRelleno" a la derecha
' de "Cadena" en "Nlong" posiciones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function PadR(cadena As String, nLong As Integer, Optional CarRelleno As String) As String
    Dim NLongL As Integer
    
    If CarRelleno = "" Then CarRelleno = " "
    cadena = LTrim(RTrim(cadena))
    If nLong >= Len(cadena) Then
        NLongL = nLong - Len(cadena)
        PadR = cadena + Replicate(CarRelleno, NLongL)
    Else
        PadR = Replicate("*", nLong)
    End If
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Replicate
' -------------------
' Función que devuelve un String con Xite repeticiones de Xvar
'
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Replicate(XVar As String, Xite As Integer) As String
    Dim x As Integer
    Dim WStr As String
    
    WStr = ""
    If Xite = 0 Then
        Replicate = WStr
        Exit Function
    End If
    For x = 1 To Xite
        WStr = WStr + XVar
    Next x
    Replicate = WStr
End Function
