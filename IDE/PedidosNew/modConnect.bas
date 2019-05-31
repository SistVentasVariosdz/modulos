Attribute VB_Name = "modConnect"
Option Explicit
Declare Function DeleteFile _
        Lib "Kernel32" _
        Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Const BlockSize = 100000

Public sAnoPeriodoContableVigente As String

Public sAnoContableVigente        As String

Public sPeriodoContableVigente    As String

Private Function NotChar(ByVal vValor As String) As String

    Dim i       As Integer

    Dim sReturn As String

    If InStr(vValor, Chr(34)) Or InStr(vValor, Chr(39)) Then

        For i = 1 To Len(vValor)

            If Asc(Mid(vValor, i, 1)) <> 39 And Asc(Mid(vValor, i, 1)) <> 34 Then
                sReturn = sReturn + Mid(vValor, i, 1)
            End If

        Next

    Else
        sReturn = vValor
    End If

    NotChar = sReturn
End Function

Function StrZero(nDato As Long, nZeros As Integer)

    Dim wdato As String, wAncho As Integer, wDatoOk As String

    Dim i     As Integer

    wdato = Trim(Str(nDato))
    wAncho = Len(wdato)

    If wAncho < nZeros Then

        For i = 1 To nZeros - wAncho
            wDatoOk = wDatoOk + "0"
        Next i

        wDatoOk = wDatoOk + wdato
    Else
        wDatoOk = wdato
    End If

    StrZero = wDatoOk
End Function

Public Function ASearch(avArray As Variant, _
                        vSearchFor As Variant, _
                        iIndice As Integer, _
                        Optional base As Variant) As Integer
                        
    ' Control de Parametro opcional
    
    Dim iIndex  As Integer

    Dim iMaxLen As Integer
    
    ' Valor de retorno si no se encuentra el elemento
    ASearch = -1
    
    iMaxLen = UBound(avArray, 2)
    
    ' Inicio de busqueda del elemento
    For iIndex = 0 To iMaxLen
    
        If avArray(iIndice, iIndex) = vSearchFor Then
        
            ASearch = iIndex

            Exit Function
        
        End If
        
    Next

End Function

Public Function GetSubString(ByVal SourceString As String, _
                             iWhatString As Integer, _
                             Optional InChars As Variant) As String

    Dim CountSubString As Integer

    CountSubString = 1
    
    If IsMissing(InChars) Then
        InChars = "-+-"
    End If
    
    ' Si no existe el caracter de separacion, retornamos la cadena de entrada
    If InStr(SourceString, InChars) = 0 Then
        GetSubString = SourceString
    End If
    
    Do While (InStr(SourceString, InChars) > 0)
    
        GetSubString = Mid(SourceString, 1, InStr(SourceString, InChars) - 1)
        
        If (iWhatString = CountSubString) Then
            
            Exit Function

        End If
        
        SourceString = Mid(SourceString, InStr(SourceString, InChars) + Len(InChars))
        CountSubString = CountSubString + 1
        
        If Not (InStr(SourceString, InChars) > 0) Then
            If (iWhatString = CountSubString) Then
                GetSubString = SourceString
            Else
                GetSubString = ""
            End If

            'GetSubString = SourceString
        End If

    Loop
    
End Function

Public Function VBsprintf2(ByRef InString As String, ParamArray aInValues()) As String

    On Error GoTo Error

    Dim OutString   As String

    Dim ThisChar    As String

    Dim IndexString As Integer

    Dim IndexValues As Integer

    Dim iNotchar    As Integer

    Dim strCadena   As String

    Dim vValor      As Variant

    OutString = ""
    IndexValues = 0

    For IndexString = 1 To Len(InString)
        ThisChar = Mid(InString, IndexString, 1)

        If ThisChar <> "$" Then
            OutString = OutString + ThisChar
        Else

            If VarType(aInValues(IndexValues)) = vbstring Then
                vValor = aInValues(IndexValues)

                If Len(vValor) > 2 Then
                    strCadena = Mid(vValor, 2, Len(vValor) - 2)
                End If

                If InStr(strCadena, Chr(34)) Or InStr(strCadena, Chr(39)) Then
                    strCadena = NotChar(strCadena)
                    vValor = "'" & strCadena & "'"
                End If

                If Mid(vValor, 1, 1) <> Chr(39) Then
                    vValor = NotChar(vValor)
                End If

            Else
                vValor = CStr(aInValues(IndexValues))
                vValor = NotChar(vValor)
            End If
   
            OutString = OutString + vValor
            IndexValues = IndexValues + 1
        End If

    Next

    VBsprintf2 = OutString

    Exit Function

Error:
    MsgBox Err.Description

End Function

Function StrVacio(nDato As Variant, nZeros As Integer)

    Dim wdato As String, wAncho As Integer, wDatoOk As String

    Dim i     As Integer

    Dim tdato As Variant

    If TypeName(nDato) = "String" Then
        If nDato = "" Then
            StrVacio = ""

            Exit Function

        Else
            tdato = nDato
        End If

    Else
        tdato = nDato
    End If

    wdato = Trim(tdato)
    wAncho = Len(wdato)

    If wAncho < nZeros Then

        For i = 1 To nZeros - wAncho
            wDatoOk = wDatoOk + " "
        Next i

        wDatoOk = wdato & wDatoOk
    Else
        wDatoOk = wdato
    End If

    StrVacio = wDatoOk
End Function

