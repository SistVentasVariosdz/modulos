Attribute VB_Name = "modRutinas"
'Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const Deshabilitado = &H8000000A
Public Const TODOS = "<TODOS>"

Declare Function GetcomputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
'Public Enum TipoRep
'    Observaciones = 1
'End Enum


Public Function ComputerName() As String
    Dim KeyName$
    Dim keylen&
    Dim iNull
            
    keylen& = 2000
    KeyName$ = String$(keylen, 0)
    
    GetcomputerName KeyName$, keylen&
    
    iNull = InStr(KeyName, Chr(0))
    ComputerName = Mid(KeyName$, 1, iNull - 1)
End Function

Function DevuelveFechaServidor() As Date
On Error GoTo hand
    DevuelveFechaServidor = DevuelveCampo("select getdate()", cCONNECT)
Exit Function
hand:
ErrorHandler Err, "DevuelveFechaServidor"
End Function


Public Function DevuelveMes(ByRef pMes As String, pIdioma As String) As Variant
On Error GoTo hand
DevuelveMes = DevuelveCampo("select dbo.uf_nombre_mes('" & Format(CInt(pMes), "0#") & "','" & pIdioma & "'", B_conexion)
Exit Function
hand:
ErrorHandler Err, "DevuelveMes"
End Function

Public Function DevuelveCampo(ByVal pQuerySql As String, ByVal pConexion As String) As Variant
On Error GoTo DevuelveCampoError
    Dim rstBuscaCampo As New ADODB.Recordset

   ' Set rstBuscaCampo.ActiveConnection = pConexion
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open pQuerySql, pConexion, adOpenKeyset, adLockOptimistic

    If rstBuscaCampo.RecordCount > 0 Then
        DevuelveCampo = rstBuscaCampo(0)
    Else
        DevuelveCampo = ""
    End If
    Set rstBuscaCampo = Nothing
Exit Function
DevuelveCampoError:
    ErrorHandler Err, "Funcion DevuelveCampo"
    Err.Clear
    DevuelveCampo = ""
    Set rstBuscaCampo = Nothing
End Function

'-------------------------------------------------------------
' Function  : EjecutarQuery()
' Propósito : Ejecutar una sentencia SQL Query
' Input     : pQuery: SQL Query
'             pCursorType: ADO Cursor Type
' Output    : ADO Recordset obtenido
'-------------------------------------------------------------
Public Function EjecutarQuery(ByVal pQuery As String, _
                              ByVal pCursorType As ADODB.CursorTypeEnum) _
                              As ADODB.Recordset
   Dim adoRs As ADODB.Recordset
   Dim adoRsUltimo As ADODB.Recordset

   Set adoRs = New ADODB.Recordset
   With adoRs
      .ActiveConnection = g_cnnConexion
      .CursorLocation = adUseClient
      .CursorType = pCursorType
      .LockType = adLockOptimistic
      .Open pQuery
   End With

   ' Se obtiene el ultimo resultado de Recordset
   Do While Not (adoRs Is Nothing)
      Set adoRsUltimo = adoRs
      Set adoRs = adoRsUltimo.NextRecordset
   Loop
   Set adoRs = adoRsUltimo

   Set EjecutarQuery = adoRs

End Function


'-------------------------------------------------------------
' Procedure : BuscaCampo()
' Propósito : Funcion que permite posicionarnos en un determinado registro
'             segun el valor ingresado
' Input     : pRs_Lista: Control ADODB.RECORDSET,
'             pCampo: Es la columna sobre la que se buscara,
'             pValor: Es el valor que se buscara
'-------------------------------------------------------------
Sub BuscaCampo(pRs_Lista As ADODB.Recordset, pCampo As String, pValor As String)
On Error GoTo hand
    Dim pIndice As Integer
    Dim pRs_Prov As New ADODB.Recordset
    
    If Not pRs_Lista.EOF And Not pRs_Lista.BOF Then
        Set pRs_Prov = pRs_Lista.Clone
        pIndice = 0
        pRs_Prov.MoveFirst
        While Not pRs_Prov.EOF
            If Mid(pRs_Prov(pCampo).Value, 1, Len(pValor)) = pValor Then
                pRs_Lista.MoveFirst
                pRs_Lista.Move (pIndice)
                
                pRs_Prov.Close
                Set pRs_Prov = Nothing
                
                Exit Sub
            End If
            pIndice = pIndice + 1
            pRs_Prov.MoveNext
        Wend
        If Not pRs_Prov Is Nothing Then
            pRs_Prov.Close
        End If
        Set pRs_Prov = Nothing
    End If

Exit Sub
hand:
ErrorHandler Err, "BuscaCampo"
pRs_Prov.Close
Set pRs_Prov = Nothing
End Sub


Sub FormateaGrid(pGrid As MSDataGridLib.DataGrid)
On Error GoTo hand
        pGrid.MarqueeStyle = dbgHighlightRow
        pGrid.HeadFont.Bold = True
        pGrid.Refresh
        pGrid.BackColor = -2147483624
Exit Sub
hand:
ErrorHandler Err, "FormateaGrid"
End Sub

Sub LlenaCombo(objObjeto As Object, strQuery As String, Conexion As String)
On Error GoTo LlenaComboError
    Dim rstBuscaCampo As New ADODB.Recordset
    'Set rstBuscaCampo.ActiveConnection = Conexion
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open strQuery, Conexion, adOpenDynamic, adLockOptimistic
    objObjeto.Clear
    If rstBuscaCampo.RecordCount > 0 Then
        With rstBuscaCampo
            If rstBuscaCampo.Fields.Count = 2 Then
                Do While Not .EOF
                    objObjeto.AddItem IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0)) & Space(3) & IIf(IsNull(rstBuscaCampo(1)), "", rstBuscaCampo(1))
                    .MoveNext
                Loop
            Else
                Do While Not .EOF
                    objObjeto.AddItem IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0))
                    .MoveNext
                Loop
            End If
        End With
    End If
Set rstBuscaCampo = Nothing
Exit Sub
LlenaComboError:
    ErrorHandler Err, "Procedimiento LlenaCombo"
    Err.Clear
    Set rstBuscaCampo = Nothing
End Sub


Sub BuscaCombo(strTexto As String, intPos As Integer, combo As ComboBox)
    Dim intCont As Integer
    Dim Encontro As Boolean
    Encontro = False
    If intPos = 1 Then
        For intCont = 0 To combo.ListCount - 1
            If strTexto = Mid(combo.List(intCont), 1, Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    Else
        For intCont = 0 To combo.ListCount - 1
            If UCase(strTexto) = Right(UCase(combo.List(intCont)), Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    End If
    If Encontro = False Then
        combo.ListIndex = -1
    End If
End Sub

'-------------------------------------------------------------
' Procedure : SoloNumeros()
' Propósito : Funcion que permite el ingreso de solo numeros
'             sobre un control Textbox
' Input     : pTextbox: Control Textbox,
'             pKeyAscii: La tecla ingresada,
'             pConDecimales: Si se usa o no decimales,
'             pNumDecimales: Numero de Decimales permitidos,
'             pNumEntero: Numero de Enteros permitidos
'-------------------------------------------------------------
Public Sub SoloNumeros(ByVal pTextbox As TextBox, _
                       ByRef pKeyAscii As Integer, _
                       Optional ByVal pConDecimales As Boolean, _
                       Optional ByVal pNumDecimales As Integer, _
                       Optional ByVal pNumEnteros As Integer)
   If pNumEnteros = 0 Then pNumEnteros = 10
   If pKeyAscii = 8 Then
      If pConDecimales And pTextbox.SelStart > 0 Then
         If Mid(pTextbox, pTextbox.SelStart, 1) = "." Then
            If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros And Len(Mid(pTextbox, pTextbox.SelStart + 1)) > 0 Then pKeyAscii = 0
         End If
      End If
      Exit Sub
   End If
   If pKeyAscii = 46 Then
      If pConDecimales Then
         If InStr(1, pTextbox, ".") > 0 Then
            pKeyAscii = 0
         Else
            If Len(Mid(pTextbox, pTextbox.SelStart + 1)) > pNumDecimales Then pKeyAscii = 0
            If pTextbox.SelStart > 0 Then If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros Then pKeyAscii = 0
         End If
      Else
         pKeyAscii = 0
      End If
   Else
      If Not (pKeyAscii >= 48 And pKeyAscii <= 57) Then pKeyAscii = 0
      If pKeyAscii = 39 Or pKeyAscii = 13 Then
         pKeyAscii = 0
      End If
      
      Dim iPos As Integer
      iPos = InStr(1, pTextbox, ".")
      If iPos > 0 And pConDecimales Then _
         If Len(Mid(pTextbox, iPos)) > pNumDecimales Then _
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") = 0 Then pKeyAscii = 0
            
      If pTextbox.SelStart < iPos Or iPos = 0 Then
         If pNumEnteros > 0 Then
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") > 0 Then
               If Len(Mid(pTextbox, 1, InStr(pTextbox.SelStart + 1, pTextbox, ".") - 1)) >= pNumEnteros Then pKeyAscii = 0
            Else
               If Len(pTextbox) >= pNumEnteros Then pKeyAscii = 0
            End If
         End If
      End If
   End If
End Sub



'-------------------------------------------------------------
' Procedure : ErrorHandler()
' Propósito : Manejo de Excepciones Genérico
' Input     : pErr: Objeto Error VB,
'             pProcedure: Nombre del Procedimiento
'-------------------------------------------------------------
Public Sub ErrorHandler(ByVal pErr As ErrObject, ByVal pProcedure As String)
   Dim sMsg As String
   
   Screen.MousePointer = vbDefault
   sMsg = pProcedure & " : " & Chr(13) & Chr(10) & _
          "Descripción : " & pErr.Description

  MsgBox sMsg, vbExclamation, App.Title
  
Err.Clear
End Sub

Function ExisteCampo(pCampo As String, pTabla As String, pValor As Variant, Conexion As String, Optional pEsStringValor As Boolean = True) As Boolean
On Error GoTo hand

If pEsStringValor Then
    If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = '" & pValor & "'", Conexion) > 0 Then
        ExisteCampo = True
    Else
        ExisteCampo = False
    End If
Else
    If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = " & pValor, Conexion) > 0 Then
        ExisteCampo = True
    Else
        ExisteCampo = False
    End If
End If
Exit Function
hand:
ErrorHandler Err, "ExisteCampo"
ExisteCampo = False
End Function


Public Sub ScrollHorizontal(ByRef varListBox As ListBox)

Dim X As Integer
Dim i As Integer
Dim maxLong As Integer
maxLong = 0
For i = 1 To varListBox.ListCount

    'List1.AddItem "El numero final de la selección es : " & i
    If maxLong < Len(Trim(varListBox.List(i))) Then
        maxLong = Len(Trim(varListBox.List(i)))
    End If

Next i
X = SendMessage(varListBox.hwnd, &H194, maxLong * 8, ByVal 0&)
'x = SendMessage(List1.hwnd, &H194, 200, ByVal 0&)

End Sub

Public Function ExecuteCommandSQL(ByVal Connect As String, ByVal sql As String) As Long
On Error GoTo errorx
Dim oCn As Object
  
  Set oCn = CreateObject("ADODB.Connection")
  oCn.Open Connect
  oCn.CommandTimeout = 900
  
  oCn.Execute sql, ExecuteCommandSQL, 128
  oCn.Close
  Set oCn = Nothing
  
Exit Function
errorx:
  ExecuteCommandSQL = -2
  Err.Raise Err.Number, "ExecuteCommandSQL", Err.Description
End Function
