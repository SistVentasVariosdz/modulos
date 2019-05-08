Attribute VB_Name = "modReportesExcel"
Option Explicit
Public Const Deshabilitado = &H8000000A
Public Enum TipoRep
    TrackingReporteDetail = 1
    DeliverySummary = 2
    Observaciones = 3
    Forecast = 4
End Enum
Public B_db As New ADODB.Connection
Sub BuscaCombo(strTexto As String, intPos As Integer, combo As ComboBox)
    Dim intCont As Integer
For intCont = 0 To combo.ListCount - 1
    If strTexto = Mid(combo.List(intCont), 1, Len(strTexto)) Then
        combo.ListIndex = intCont
        Exit For
    End If
Next
End Sub

Sub FormateaGrid(pGrid As MSDataGridLib.DataGrid)
On Error GoTo hand
        pGrid.MarqueeStyle = dbgHighlightRow
        pGrid.HeadFont.Bold = True
        pGrid.Refresh
Exit Sub
hand:
ErrorHandler Err, "FormateaGrid"
End Sub
Public Function DevuelveCampo(pQuerySql As String, pConexion As String) As Variant
On Error GoTo DevuelveCampoError
    Dim rstBuscaCampo As New ADODB.Recordset

   ' Set rstBuscaCampo.ActiveConnection = pConexion
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open pQuerySql, pConexion, adOpenKeyset, adLockOptimistic
   ' Debug.Print pQuerySql

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

Public Sub HabilitaMant(ctl As Object, botones As String)
ctl.FunctionsUser = botones

'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
End Sub

Sub LlenaCombo(objObjeto As Object, strQuery As String, Conexion As String)
On Error GoTo LlenaComboError
    Dim rstBuscaCampo As New ADODB.Recordset
    
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open strQuery, Conexion, adOpenDynamic, adLockOptimistic
    
    If rstBuscaCampo.RecordCount > 0 Then
        objObjeto.Clear
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
   sMsg = pProcedure & " : " & _
          Format(Now, "dd/mm/yyyy - hh:mm:ss") & Chr(13) & Chr(13) & _
          "Número : " & pErr.Number & Chr(13) & _
          "Descripción : " & pErr.Description & Chr(13) & _
          "Fuente : " & pErr.Source & Chr(13)
          Err.Clear
   MsgBox sMsg, vbCritical, App.Title
End Sub



Sub GeneraReportes(pTipo As TipoRep, Optional pMes As String, Optional pFabrica As String, Optional pCliente As String, Optional pRegistros As ADODB.Recordset)
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Select Case pTipo
    Case 1
        Ruta = App.Path & "\delivery.xlt"
    Case 2
        Ruta = App.Path & "\summary.xlt"
    Case 4
        Ruta = App.Path & "\Forecast.xlt"
        
End Select
'Usu = "Usuario : " & MDIPrincipal.pUsuario
Usu = vusu

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
'    oo.WindowState = xlMaximized
    oo.DisplayAlerts = False
    oo.Run "ArmarReporte", CStr(iLanguage), DSN_Empresa, Usu
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

