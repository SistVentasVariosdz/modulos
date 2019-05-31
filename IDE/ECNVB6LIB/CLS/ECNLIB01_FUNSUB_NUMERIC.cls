VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB01_FUNSUB_NUMERIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Ks_Moneda_Nuevo_Sol = "Nuevos Soles"
Private Const Ks_Moneda_Dolar_Ame = "Dólares"
Private Const Ks_Moneda_Euro = "Euros"

Public Enum EruMonedas
    [NUEVO SOL] = 1
    [US$] = 2
    [EURO] = 3
End Enum

Function Redondear(ByVal dblNumero As Double) As Double
    Dim dblResultado As Double
    Dim dblDiferencia As Double
    
    dblResultado = Int(dblNumero)
    dblDiferencia = dblNumero - dblResultado
    If dblDiferencia >= 0.5 Then
       Redondear = dblResultado + 1
    Else
       Redondear = dblResultado
    End If
End Function

'***************************************************************************************************************************************************************************************************************************************
' ConvertirNumeroEnLetras : Devuelve en letras el sNumero ingresado como parámetro
'***************************************************************************************************************************************************************************************************************************************
Public Function ConvertirNumeroEnLetras(ByVal sNumero As String, _
                                        ByVal opcMoneda As EruMonedas) As String
    '********Declara variables de tipo cadena************
    Dim palabras, entero, dec, flag As String
    Dim sMoneda As String

    '********Declara variables de tipo entero***********
    Dim num, x, y As Integer

    flag = "N"

    '**********Número Negativo***********
    If Mid(sNumero, 1, 1) = "-" Then
        sNumero = Mid(sNumero, 2, CStr(Len(sNumero) - 1))
        palabras = "menos "
    End If

    '**********Si tiene ceros a la izquierda*************
    For x = 1 To Len(sNumero)
        If Mid(sNumero, 1, 1) = "0" Then
            sNumero = Trim(Mid(sNumero, 2, CStr(Len(sNumero))))
            If Trim(Len(sNumero)) = 0 Then palabras = ""
        Else
            Exit For
        End If
    Next

    '*********Dividir parte entera y decimal************
    For y = 1 To Len(sNumero)
        If Mid(sNumero, y, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(sNumero, y, 1)
            Else
                dec = dec + Mid(sNumero, y, 1)
            End If
        End If
    Next y

    If Len(dec) = 1 Then dec = dec & "0"

    '**********proceso de conversión***********
    flag = "N"

    If Val(sNumero) <= 999999999 Then
        For y = Len(entero) To 1 Step -1
            num = Len(entero) - (y - 1)
            Select Case y
                Case 3, 6, 9
                    '**********Asigna las palabras para las centenas***********
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                palabras = palabras & "cien "
                            Else
                                palabras = palabras & "ciento "
                            End If
                        Case "2"
                            palabras = palabras & "doscientos "
                        Case "3"
                            palabras = palabras & "trescientos "
                        Case "4"
                            palabras = palabras & "cuatrocientos "
                        Case "5"
                            palabras = palabras & "quinientos "
                        Case "6"
                            palabras = palabras & "seiscientos "
                        Case "7"
                            palabras = palabras & "setecientos "
                        Case "8"
                            palabras = palabras & "ochocientos "
                        Case "9"
                            palabras = palabras & "novecientos "
                    End Select
                Case 2, 5, 8
                    '*********Asigna las palabras para las decenas************
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If Mid(entero, num + 1, 1) = "0" Then
                                flag = "S"
                                palabras = palabras & "diez "
                            End If
                            If Mid(entero, num + 1, 1) = "1" Then
                                flag = "S"
                                palabras = palabras & "once "
                            End If
                            If Mid(entero, num + 1, 1) = "2" Then
                                flag = "S"
                                palabras = palabras & "doce "
                            End If
                            If Mid(entero, num + 1, 1) = "3" Then
                                flag = "S"
                                palabras = palabras & "trece "
                            End If
                            If Mid(entero, num + 1, 1) = "4" Then
                                flag = "S"
                                palabras = palabras & "catorce "
                            End If
                            If Mid(entero, num + 1, 1) = "5" Then
                                flag = "S"
                                palabras = palabras & "quince "
                            End If
                            If Mid(entero, num + 1, 1) > "5" Then
                                flag = "N"
                                palabras = palabras & "dieci"
                            End If
                        Case "2"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "veinte "
                                flag = "S"
                            Else
                                palabras = palabras & "veinti"
                                flag = "N"
                            End If
                        Case "3"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "treinta "
                                flag = "S"
                            Else
                                palabras = palabras & "treinta y "
                                flag = "N"
                            End If
                        Case "4"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "cuarenta "
                                flag = "S"
                            Else
                                palabras = palabras & "cuarenta y "
                                flag = "N"
                            End If
                        Case "5"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "cincuenta "
                                flag = "S"
                            Else
                                palabras = palabras & "cincuenta y "
                                flag = "N"
                            End If
                        Case "6"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "sesenta "
                                flag = "S"
                            Else
                                palabras = palabras & "sesenta y "
                                flag = "N"
                            End If
                        Case "7"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "setenta "
                                flag = "S"
                            Else
                                palabras = palabras & "setenta y "
                                flag = "N"
                            End If
                        Case "8"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "ochenta "
                                flag = "S"
                            Else
                                palabras = palabras & "ochenta y "
                                flag = "N"
                            End If
                        Case "9"
                            If Mid(entero, num + 1, 1) = "0" Then
                                palabras = palabras & "noventa "
                                flag = "S"
                            Else
                                palabras = palabras & "noventa y "
                                flag = "N"
                            End If
                    End Select
                Case 1, 4, 7
                    '*********Asigna las palabras para las unidades*********
                    Select Case Mid(entero, num, 1)
                        Case "1"
                            If flag = "N" Then
                                If y = 1 Then
                                    palabras = palabras & "uno "
                                Else
                                    palabras = palabras & "un "
                                End If
                            End If
                        Case "2"
                            If flag = "N" Then palabras = palabras & "dos "
                        Case "3"
                            If flag = "N" Then palabras = palabras & "tres "
                        Case "4"
                            If flag = "N" Then palabras = palabras & "cuatro "
                        Case "5"
                            If flag = "N" Then palabras = palabras & "cinco "
                        Case "6"
                            If flag = "N" Then palabras = palabras & "seis "
                        Case "7"
                            If flag = "N" Then palabras = palabras & "siete "
                        Case "8"
                            If flag = "N" Then palabras = palabras & "ocho "
                        Case "9"
                            If flag = "N" Then palabras = palabras & "nueve "
                    End Select
            End Select

            '***********Asigna la palabra mil***************
            If y = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                Len(entero) <= 6) Then palabras = palabras & "mil "
            End If

            '**********Asigna la palabra millón*************
            If y = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    palabras = palabras & "millón "
                Else
                    palabras = palabras & "millones "
                End If
            End If
        Next y

        '**********Une la parte entera y la parte decimal*************
        
        Dim xTipMoneda As String

        Select Case opcMoneda
            Case [NUEVO SOL]: sMoneda = Ks_Moneda_Nuevo_Sol
            Case [US$]: sMoneda = Ks_Moneda_Dolar_Ame
            Case [EURO]: sMoneda = Ks_Moneda_Euro
        End Select
               

        If dec <> "" Then
            ConvertirNumeroEnLetras = palabras & "y " & dec & "/100" & sMoneda
        Else
            ConvertirNumeroEnLetras = palabras & sMoneda
        End If
    Else
            ConvertirNumeroEnLetras = ""
        End If
End Function



