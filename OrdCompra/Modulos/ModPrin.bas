Attribute VB_Name = "Mod_principal"
Global B_db As New ADODB.Connection
Global B_sql As String
Global B_conexion As String
Global Const IDIOMA = "I"
Sub Main()
InitMessages
End Sub
Sub AVANZA(ByVal Tecla As Integer)
    Select Case Tecla
        Case 13, 40: SendKeys "{TAB}", True
        Case 38: SendKeys "+{TAB}", True
    End Select
End Sub
'Sub Informa(ByVal Mens As String, Optional ByVal amensaje As clsMensaje)
'If Mens <> "" Then
'    Dim rpta As Byte
'    rpta = MsgBox(Mens, vbInformation, "Informa")
'    Exit Sub
'End If
'Dim aMess(4)
'LoadMessage aMess, amensaje.Codigo
'amensaje.ShowMsg (aMess)
'End Sub
Function Pregunta(ByVal Mens As String) As Byte
    Pregunta = MsgBox(Mens, vbQuestion + vbYesNo, "Pregunta")
End Function
Public Sub Imprime(ByVal Menu As Integer)
Dim rpta As Byte
Dim Nom_rep As String
Select Case Menu
Case 1: Nom_rep = "Primer Reporte"
Case 2:
Case 3:
Case 4:
End Select
rpta = Pregunta("¿Desea enviar este reporte por correo?")
If rpta = 6 Then
    Informa ("Espere un momento")
End If
End Sub

