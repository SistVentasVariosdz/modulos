Attribute VB_Name = "modFunciones"
'JUAN MANUEL MIRANDA TORREALVA
Public Function generaEspacioCadena(ByVal strPalabra As String, ByVal numTamPalabra As Integer) As String
Dim strNuevaCadena As String
Dim intCantidadEspacioAGenerar As Integer
intCantidadEspacioAGenerar = 0
strNuevaCadena = ""

intCantidadEspacioAGenerar = numTamPalabra - Len(Trim(strPalabra))

strNuevaCadena = Trim(strPalabra) & strRepeat(" ", intCantidadEspacioAGenerar)

generaEspacioCadena = strNuevaCadena
End Function
Public Function GenerarEspaciosEnPalabra(strPalabra As String)
Dim i As Integer
Dim nuevaPalabra As String
nuevaPalabra = ""
For i = 1 To Len(strPalabra)
    nuevaPalabra = nuevaPalabra & Mid(strPalabra, i, 1) & " "
Next i
GenerarEspaciosEnPalabra = nuevaPalabra
End Function
Public Function strRepeat(ByVal strChar As String, ByVal intCantidad As Integer)
Dim i As Integer
Dim strNuevaCadena As String
strNuevaCadena = ""
For i = 1 To intCantidad
    strNuevaCadena = strNuevaCadena & strChar
Next
strRepeat = strNuevaCadena
End Function


Sub LlenaComboBox(objObjeto As Object, strQuery As String, Conexion As String)
On Error GoTo LlenaComboError
    Dim rstBuscaCampo As New ADODB.Recordset
    
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open strQuery, Conexion, adOpenDynamic, adLockOptimistic
        objObjeto.Clear
    If rstBuscaCampo.RecordCount > 0 Then

        With rstBuscaCampo
            If rstBuscaCampo.Fields.count = 2 Then
                Do While Not .EOF
                    objObjeto.AddItem Trim(IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0))) & Space(3) & Trim(IIf(IsNull(rstBuscaCampo(1)), "", rstBuscaCampo(1)))
                    .MoveNext
                Loop
            Else
                Do While Not .EOF
                    objObjeto.AddItem Trim(IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0)))
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

Sub BuscaCombobox(strTexto As String, intPos As Integer, combo As ComboBox)
    Dim intCont As Integer
combo.ListIndex = -1
If Len(Trim(strTexto)) = 0 Then Exit Sub
For intCont = 0 To combo.ListCount - 1
    If Trim(strTexto) = Trim(Mid(combo.List(intCont), intPos, Len(strTexto))) Then
        combo.ListIndex = intCont
        Exit For
    End If
Next
End Sub
Public Sub MSG_EXIT_FORM(ByVal Frm As Form)
    If MsgBox("¿Desea cerrar la ventana?", vbYesNo + vbInformation, Frm.Caption) = vbYes Then
        Unload Frm
    End If
End Sub

'"MODIFICAR/IMPRIMIR/DESHACER" son el valor de la propiedad tag de los controles
' por default todos los controles se encuentran con un estado Enabled = false
Public Sub permisos(Frm As Form, accesos As String)

    Dim access As String ' variable temporal

    Dim Cadena() As String ' array de string

    Cadena = Split(accesos, "/") 'convierto la cadena en un array

    For i = 0 To UBound(Cadena) ' recorremos el array

        Dim c As Control

        For Each c In Frm.Controls ' recorremos los controles del formulario

            ' preguntamos si el valor de la propiedad Tag del control que actualmente
            ' se está recorriendo es el valor del array actual
            If c.Tag = Cadena(i) Then  ' habilitamos los controles que se encuentran en el array
                c.Enabled = True ' y el control está deshabilitado
                Call Deshabilitar_Frame(c, True)
            End If

        Next
    Next

End Sub

Private Sub Deshabilitar_Frame(UnFrame As Control, estado As Boolean)

    On Error Resume Next

    'Variable de tipo Control Para los controles del contenedor en este caso del Frame
    Dim ElControl As Control

    'recorre los controles
    If (TypeOf UnFrame Is Frame) Then

        For Each ElControl In UnFrame.Controls

            'si está dentro lo deshabilita o habilita
            If ElControl.Container Is UnFrame Then
                ElControl.Enabled = estado
            End If

        Next

    End If

End Sub

Public Function ENCUENTRA_EN_TEXTO(ByVal Cadena As String, ByVal cadenaBuscar As String) As Boolean
    Dim intEncontroPos As Integer
    Dim boolEncontroPos As Boolean
    intEncontroPos = InStr(Cadena, cadenaBuscar)
        If intEncontroPos > 0 Then
            boolEncontroPos = True
        Else
            boolEncontroPos = False
        End If
    ENCUENTRA_EN_TEXTO = boolEncontroPos
End Function
Public Sub mensajeError(Frm As Form, _
                        strMetodo As String, _
                        strNumero As String, _
                        strDescripcion As String, _
                        strMensaje As String)

    Dim strMsgErr As String

    strMsgErr = "FORMULARIO : " & Frm.Name
    strMsgErr = strMsgErr & vbNewLine & "METODO : " & strMetodo
    strMsgErr = strMsgErr & vbNewLine & "N° de Error : " & strNumero
    strMsgErr = strMsgErr & vbNewLine & "Descripción       : " & strDescripcion
    strMsgErr = strMsgErr & vbNewLine & "Mensaje      : " & UCase(strMensaje)
    MsgBox strMsgErr, vbCritical, "Formulario : " & Frm.Caption
End Sub

'Juan Manuel Miranda Torrealva 2014
Public Function CompletaCodigo(CodOrigen As String, _
                               longcodfinal As Integer, _
                               PosfinalCod As Integer) As String

    ' CodOrigen     = Es el codigo que sera pasado por parametro
    ' LongCodFinal  = Es el tamaño del Codigo a devolver
    ' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim contador As Integer

    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)

    For contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next

    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Len(CodOrigen) - PosfinalCod)
End Function

Public Sub focoControl(c As Control, tieneFoco As Boolean)

    If (tieneFoco And TypeOf c Is TextBox) Then
        c.BackColor = &HC0FFFF
    Else
        c.BackColor = &H80000005
    End If

End Sub

Public Function validarFormulario(Frm As Form, Control As String)

    Dim error  As String '

    Dim ctrl() As String ' array de string

    ctrl = Split(Control, "\") 'convierto la cadena en un array

    Dim c As Control
        
    For Each c In Frm.Controls ' recorremos los controles del formulario

        ' preguntamos si el valor de la propiedad Tag del control que actualmente
        ' se está recorriendo es el valor del array actual
        ' MsgBox (ctrl(i) & "-")
          
        For i = 0 To UBound(ctrl) ' recorremos el array
             
            If c.Tag = ctrl(i) Then  ' validamos los controles que se encuentran en el array

                If (TypeOf c Is TextBox) Then
                    If (Len(Trim(c.Text)) = 0) Then
                        error = c.Tag
                        c.SetFocus
                        c.BackColor = &HC0FFFF

                        Exit For

                    End If

                ElseIf (TypeOf c Is DataCombo) Then

                    Dim lenCod As Integer

                    Dim iCtrl  As Integer

                    iCtrl = InStr(1, c.Tag, "|", vbTextCompare)
                    lenCod = Right(Trim(c.Tag), Len(Trim(c.Tag)) - iCtrl)

                    If (Len(Trim(c.BoundText)) <> lenCod) Then
                        error = Mid(c.Tag, 1, iCtrl - 1)
                        c.SetFocus

                        Exit For

                    End If

                    '  ElseIf (TypeOf c Is KEXPCheck) Then
                
                    '                    Dim lenGrupo As Integer
                    '                    Dim iCtrl As Integer
                    '                    iCtrl = InStr(1, c.Tag, "|", vbTextCompare)
                    '                    lenGrupo = Len(Trim(c.Tag)) - iCtrl + 1
                    '
                    '                    If (Mid(c.Tag, 1, iCtrl - 1) = Mid(ctrl(i), 1, iCtrl - 1)) Then
                    '
                    '                    End If
                    '                    If (c.Value = False) Then
                    '                        error = Mid(c.Tag, 1, iCtrl - 1)
                    '                        c.SetFocus
                    '                        Exit For
                    '                    End If
                    
                End If
            End If

            '            For Each ElControl In UnFrame.Controls
            '            'si está dentro lo deshabilita o habilita
            '            If ElControl.Container Is UnFrame Then
            '                ElControl.Enabled = Estado
            '            End If

        Next

        If (error <> "") Then

            Exit For

        End If

    Next

    validarFormulario = error
End Function
