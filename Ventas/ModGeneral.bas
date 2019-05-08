Attribute VB_Name = "ModGeneral"

Global conn As New ADODB.Connection

Public Function GetRecordset(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object '
  Dim objAdoCn As Object '
  
' If vValid Then
  Set objADORs = CreateObject("ADODB.Recordset") 'CreateObject("ADODB.Recordset") '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  objADORs.Open SQL, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset = objADORs
  Set GetRecordset.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 'End If
Exit Function
ehGetRecordset:
  err.Raise err.Number, err.Source, err.Description
  MsgBox err.Description
End Function

Public Sub HabilitaMant(ctl As Object, botones As String)
ctl.FunctionsUser = botones

'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
End Sub
Public Function fecha(ByVal mes As String)
    If iLanguage <> "1" Then
        If mes = "01" Then
            fecha = "January"
        ElseIf mes = "02" Then
            fecha = "February"
        ElseIf mes = "03" Then
            fecha = "March"
        ElseIf mes = "04" Then
            fecha = "April"
        ElseIf mes = "05" Then
            fecha = "May"
        ElseIf mes = "06" Then
            fecha = "June"
        ElseIf mes = "07" Then
            fecha = "July"
        ElseIf mes = "08" Then
            fecha = "August"
        ElseIf mes = "09" Then
            fecha = "September"
        ElseIf mes = "10" Then
            fecha = "October"
        ElseIf mes = "11" Then
            fecha = "November"
        ElseIf mes = "12" Then
            fecha = "December"
        End If
    Else
        If mes = "01" Then
            fecha = "Enero"
        ElseIf mes = "02" Then
            fecha = "Febrero"
        ElseIf mes = "03" Then
            fecha = "Marzo"
        ElseIf mes = "04" Then
            fecha = "Abril"
        ElseIf mes = "05" Then
            fecha = "Mayo"
        ElseIf mes = "06" Then
            fecha = "Junio"
        ElseIf mes = "07" Then
            fecha = "Julio"
        ElseIf mes = "08" Then
            fecha = "Agosto"
        ElseIf mes = "09" Then
            fecha = "Setiembre"
        ElseIf mes = "10" Then
            fecha = "Octubre"
        ElseIf mes = "11" Then
            fecha = "Noviembre"
        ElseIf mes = "12" Then
            fecha = "Diciembre"
        End If
    End If
End Function
