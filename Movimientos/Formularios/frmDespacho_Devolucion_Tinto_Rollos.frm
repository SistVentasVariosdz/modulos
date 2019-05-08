VERSION 5.00
Begin VB.Form frmDespacho_Devolucion_Tinto_Rollos 
   Caption         =   "Form1"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDespacho_Devolucion_Tinto_Rollos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String, sFec_MovStk As String, iNum_Despacho As Integer, _
       sCod_Cliente_Tex As String, Codigo As String, Descripcion As String, TipoAdd As String, sNum_MovStk As String, sNum_Secuencia As String
       
       Dim StrSQL As String, rstAux As ADODB.Recordset

Private Sub chkRepesar_Click()
    txtPeso = 0
    txtPeso.Visible = (chkRepesar = 1)
    lblPeso.Visible = (chkRepesar = 1)
    cmdCapturarPeso.Visible = (chkRepesar = 1)
End Sub

Private Sub chkRepesar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdCapturarPeso_Click()
    txtPeso.Text = CapturaPeso
    
    'EjecBatch vRuta & "\TOL02.PIF"
    'Call LeerArchivo
    
    If RTrim(txtPeso.Text) <> "0" Then
        'FunctButt1.SetFocus
        FunctButt1_ActionClick 0, 0, "ACEPTAR"
    End If
End Sub

Private Sub LeerArchivo()
Dim Archivo$
Dim f, a As Variant

Archivo = vRuta & "\TEXTO.TXT"
If Dir(Archivo) <> "" Then
    f = FreeFile
    Open Archivo For Input As #f
        Line Input #f, a
        txtPeso = Trim(a)
        Close #f
End If
End Sub

Private Sub cmdCapturarPeso_GotFocus()
    cmdCapturarPeso_Click
End Sub

Private Sub Form_Load()
    chkRepesar = 1
    chkRepesar = 0
    optBarra = True
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        AddDspDet
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub AddDspDet()
On Error GoTo Fin
Dim sTit As String, sPrefMaq As String, sCodRollo As String, sRolloAdd As String, sSufRollo As String
    
    sTit = "Agregar Detalle de Despacho"
    
'    If lblNom_Cliente = "" Then
'        MsgBox "No se encontraron los datos de la O.T.", vbExclamation + vbOKOnly, sTit
'        txtCod_OrdTra.SetFocus
'        Exit Sub
'    End If
    
    Select Case True
    Case optBarra
        txtCod_Barra = LTrim(RTrim(txtCod_Barra))
        If Len(txtCod_Barra) < 7 Then
            MsgBox "Codigo de Barra Inválido", vbExclamation + vbOKOnly, sTit
            txtCod_Barra.SetFocus
            Exit Sub
        End If
        If Len(txtCod_Barra) > 10 Then
            sRolloAdd = Left(txtCod_Barra, 1)
            sPrefMaq = Right(txtCod_Barra, 7)
            sCodRollo = Left(sPrefMaq, 5)
            sPrefMaq = Right(sPrefMaq, 2)
        End If
        
        If Len(txtCod_Barra) < 10 Then
            sPrefMaq = Left(txtCod_Barra, 2)
            sCodRollo = Mid(txtCod_Barra, 3, 7)
        End If
        
        
        
'        If Not IsNumeric(sRolloAdd) Then
'            sCodRollo = sCodRollo & sRolloAdd
'        End If
    If Len(txtCod_Barra) > 10 Then
        If IsNumeric(Mid(txtCod_Barra, 1, 1)) Then
            sSufRollo = ""
        Else
            sSufRollo = Mid(txtCod_Barra, 1, 1)
            If Not IsNumeric(Mid(txtCod_Barra, 2, 1)) Then
                sSufRollo = sSufRollo & Mid(txtCod_Barra, 2, 1)
            End If
        End If
    End If
    Case optDirecto
        txtPrefijo_Maquina = Trim(txtPrefijo_Maquina)
        txtCodigo_Rollo = Trim(txtCodigo_Rollo)
'        If Len(txtCodigo_Rollo) <> txtCodigo_Rollo.MaxLength Or _
'        Len(txtPrefijo_Maquina) <> txtPrefijo_Maquina.MaxLength Then
'            MsgBox "Codigo de Rollo Inválido", vbExclamation + vbOKOnly, sTit
'            txtPrefijo_Maquina.SetFocus
'            Exit Sub
'        End If
        sPrefMaq = txtPrefijo_Maquina
        sCodRollo = Left(txtCodigo_Rollo, 5)
        sSufRollo = Mid(txtCodigo_Rollo, 6, 2)
    End Select
    
    If Not IsNumeric(txtPeso) Then
        MsgBox "Se debe especificar el valor del repeso", vbExclamation + vbOKOnly, sTit
        txtPeso.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtPeso) <= 0 And chkRepesar = 1 Then
        MsgBox "El repeso debe ser mayor a 0", vbExclamation + vbOKOnly, sTit
        txtPeso.SetFocus
        Exit Sub
    End If
    
    sCodRollo = sCodRollo & sSufRollo
    
    StrSQL = "LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS_DESPACHO_PARTIDA 'I', '" & sCod_Almacen & _
    "', '" & sNum_MovStk & "', " & sNum_Secuencia & ", '" & _
    sPrefMaq & "', '" & sCodRollo & "', " & txtPeso & ", " & _
    txtUnidades & ", '" & IIf(chkRepesar = 1, "S", "N") & "', '" & sUsuario & "' "

    ExecuteSQL cConnect, StrSQL
    
    LimpiaDet
    
    Select Case True
    Case optBarra: txtCod_Barra.SetFocus
    Case optDirecto: txtPrefijo_Maquina.SetFocus
    End Select
    Beep
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
    txtCod_Barra = ""
    Select Case True
    Case optBarra: txtCod_Barra.SetFocus
    Case optDirecto: txtPrefijo_Maquina.SetFocus
    End Select
    Beep
    Beep
    Beep
End Sub

Private Sub BuscaOT()
On Error GoTo Fin
Dim iCol As Long, sTit As String
    
    sTit = "Busqueda de O.T."
    
    txtCod_OrdTra = Trim(txtCod_OrdTra)
    
    StrSQL = "SELECT Cod_OrdTra, Fec_Generacion, Ser_OrdComp + '-' + Cod_OrdComp " & _
             "+ '-' + Sec_OrdComp AS OrdComp FROM TX_ORDTRA_TEJEDURIA " & _
             "WHERE Cod_Cliente_Tex = '" & sCod_Cliente_Tex & "' " & _
             "AND   Cod_OrdTra like '%" & txtCod_OrdTra & "%' " & _
             "AND   Flg_Status = 'E' ORDER BY Cod_OrdTra"
    
    txtCod_OrdTra = ""
    
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Cod_OrdTra").Caption = "O.T."
        .DGridLista.Columns("Cod_OrdTra").Width = 700
        .DGridLista.Columns("Fec_Generacion").Caption = "Fec.Gen."
        .DGridLista.Columns("Fec_Generacion").Width = 1200
        .DGridLista.Columns("OrdComp").Caption = "Orden"
        .DGridLista.Columns("OrdComp").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod_OrdTra = Trim(rstAux!Cod_OrdTra)
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Cliente"
End Sub

Private Sub DatosOT()
On Error GoTo Fin
Dim sTit As String, rstAux As ADODB.Recordset
    sTit = "Mostrar Datos de O.T."
    
    StrSQL = "SELECT a.Cod_Cliente_Tex, a.Cod_FamGrupo, b.Nom_Cliente, c.Des_FamGrupo " & _
             "FROM   TX_ORDTRA_TEJEDURIA a, TX_CLIENTE b, TX_FAMGRUPO_TEJEDURIA c " & _
             "WHERE  Cod_OrdTra = '" & txtCod_OrdTra & "' " & _
             "AND    b.Cod_Cliente_Tex = a.Cod_Cliente_Tex " & _
             "AND    c.Cod_FamGrupo = a.Cod_FamGrupo"
    
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    lblNom_Cliente = ""
    lblDes_Fam = ""
    lblCod_Fam = ""
    With rstAux
    If .RecordCount > 0 Then
        .MoveFirst
        lblNom_Cliente = !Nom_Cliente
        lblDes_Fam = !Des_FamGrupo
        lblCod_Fam = !Cod_FamGrupo
    End If
    .Close
    End With
    
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub optBarra_Click()
    OpcionesVisibles
End Sub

Private Sub optBarra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCod_Barra.SetFocus
End Sub

Private Sub optDirecto_Click()
    OpcionesVisibles
End Sub

Private Sub optDirecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtPrefijo_Maquina.SetFocus
End Sub

Private Sub txtCod_Barra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCod_Barra = Trim(txtCod_Barra)
        If txtCod_Barra = "" Then
            optDirecto = True
            SendKeys "{TAB}"
        Else
            If cmdCapturarPeso.Visible Then
                cmdCapturarPeso.SetFocus
            Else
                FunctButt1_ActionClick 0, 0, "ACEPTAR"
                'FunctButt1.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtCod_OrdTra_Change()
    lblNom_Cliente = ""
    lblDes_Fam = ""
End Sub

Private Sub txtCod_OrdTra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(sCod_Cliente_Tex) <> "" Then
            BuscaOT
        Else
            txtCod_OrdTra = Format(txtCod_OrdTra, "00000")
        End If
        DatosOT
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtCodigo_Rollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtPrefijo_Maquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub OpcionesVisibles()
    LimpiaDet
    'txtCod_Barra.Visible = optBarra
    'txtPrefijo_Maquina.Visible = optDirecto
    'txtCodigo_Rollo.Visible = optDirecto
End Sub

Private Sub LimpiaDet()
On Error GoTo errx
Dim mRS As ADODB.Recordset
    txtCod_Barra = ""
    txtPrefijo_Maquina = ""
    txtCodigo_Rollo = ""
    txtPeso = 0
    
    'StrSql = "SELECT COUNT(Num_Rollo) FROM TX_ORDTRA_TEJEDURIA_ROLLOS " & _
    '         "WHERE Cod_Almacen_Despacho = '" & sCod_Almacen & "' " & _
    '         "AND   Fec_MovStk_Despacho = '" & sFec_MovStk & "' " & _
    '         "AND   Num_Despacho = '" & iNum_Despacho & "'"
    
    StrSQL = " SM_RESUMEN_DESPACHO '" & sCod_Almacen & "','" & _
               sFec_MovStk & "','" & iNum_Despacho & "'"
    Set mRS = GetRecordset(cConnect, StrSQL)
    
    If Not mRS.EOF Then
        lblKgs = FixNulos(mRS!Kgs, vbDouble)
        lblUnid = FixNulos(mRS!Unidades, vbDouble)
        lblRollos = FixNulos(mRS!Rollos, vbDouble)
    End If
    
    mRS.Close
    Set mRS = Nothing
Exit Sub
errx:
    If Not mRS Is Nothing Then
        mRS.Close
    End If
    Set mRS = Nothing
    
    MsgBox err.Description, vbCritical + vbOKOnly, StrSQL
End Sub

Private Sub txtUnidades_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Function CapturaPeso()
  On Error GoTo ControlErrores
    Dim sBuffer As String
    sBuffer = String(19, 0)
    If (Captura) Then
       CapturaPeso = Captura / 100
  Else
      MsgBox "Error en Lectura. Comuníquese con Sistemas", vbExclamation
   End If
  
  Exit Function
ControlErrores:
  CapturaPeso = -1
End Function



