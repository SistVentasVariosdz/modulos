Attribute VB_Name = "ModGeneral"

Global conn As New ADODB.Connection

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
