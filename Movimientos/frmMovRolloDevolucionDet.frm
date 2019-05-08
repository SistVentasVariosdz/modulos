VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmMovRolloDevolucionDet 
   Caption         =   "Movimiento Devolucion Detalle"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMovRolloDet 
      Height          =   3765
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdCapturarPeso 
         Caption         =   "Capturar Peso"
         Height          =   270
         Left            =   3960
         TabIndex        =   3
         Top             =   750
         Width           =   1425
      End
      Begin VB.TextBox TxtNom_Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1455
         Width           =   2955
      End
      Begin VB.TextBox txtPrefijo_Maquina 
         Height          =   285
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   1
         Top             =   735
         Width           =   705
      End
      Begin VB.TextBox txtCodigo_Rollo 
         Height          =   285
         Left            =   2370
         MaxLength       =   6
         TabIndex        =   2
         Top             =   735
         Width           =   1440
      End
      Begin VB.TextBox txtKgs_Rollo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1605
         TabIndex        =   4
         Text            =   "0"
         Top             =   1095
         Width           =   705
      End
      Begin VB.TextBox txtUnidades 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3645
         TabIndex        =   5
         Text            =   "0"
         Top             =   1095
         Width           =   705
      End
      Begin VB.TextBox txtObservacion 
         Enabled         =   0   'False
         Height          =   930
         Left            =   1605
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2685
         Width           =   5205
      End
      Begin VB.TextBox txtCod_Barra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1605
         MaxLength       =   37
         TabIndex        =   0
         Top             =   315
         Width           =   5505
      End
      Begin VB.OptionButton optDirecto 
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   345
         TabIndex        =   17
         Top             =   765
         Width           =   1260
      End
      Begin VB.OptionButton optBarra 
         Caption         =   "Cod. Barra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   345
         TabIndex        =   16
         Top             =   330
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox txtCod_Calidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1605
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1845
         Width           =   675
      End
      Begin VB.TextBox txtDes_Calidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1845
         Width           =   2115
      End
      Begin VB.TextBox lblAbr_Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1455
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Peso (Kgs)"
         Height          =   210
         Left            =   450
         TabIndex        =   25
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidades Rollo"
         Height          =   195
         Left            =   2490
         TabIndex        =   24
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "O.T."
         Height          =   180
         Left            =   450
         TabIndex        =   23
         Top             =   2310
         Width           =   870
      End
      Begin VB.Label txtCod_OrdTra 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1605
         TabIndex        =   11
         Top             =   2265
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "Tela"
         Height          =   180
         Left            =   2595
         TabIndex        =   22
         Top             =   2325
         Width           =   405
      End
      Begin VB.Label txtDes_Tela 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3090
         TabIndex        =   12
         Top             =   2265
         Width           =   3720
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   450
         TabIndex        =   21
         Top             =   2685
         Width           =   1095
      End
      Begin VB.Label lblCalidad 
         Caption         =   "Calidad"
         Height          =   180
         Left            =   450
         TabIndex        =   20
         Top             =   1890
         Width           =   555
      End
      Begin VB.Label lblTipoOC 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   5205
         TabIndex        =   10
         Top             =   1845
         Width           =   1650
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo O/C"
         Height          =   180
         Left            =   4470
         TabIndex        =   19
         Top             =   1890
         Width           =   660
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   450
         TabIndex        =   18
         Top             =   1485
         Width           =   555
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2430
      TabIndex        =   14
      Top             =   3960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmMovRolloDevolucionDet.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmMovRolloDevolucionDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String, Descripcion As String, TipoAdd As String
Public sAccion As String, sCod_Almacen As String, sCod_TipMov As String, _
       sCod_Calidad As String, sCod_ClaMov As String, sCod_TipAnx As String, _
       sNum_MovStk As String, sNum_Secuencia As String, sDes_TipMov As String
Dim rstAux As ADODB.Recordset, StrSql As String
Dim bExigeUnidades As Boolean
Public sTip_PtMP As String

Public sFlg_Devolucion_Rollos_Tejeduria As String

Private Sub BuscaDatosRollo()
On Error GoTo Fin
Dim sPrefMaq As String, sCodRollo As String
Dim bExigeUnidades As Boolean
Dim sCod_Tela_Ot As String

sPrefMaq = "": sCodRollo = ""
    Select Case True
    Case optBarra
        txtCod_Barra = Trim(txtCod_Barra)
        If Len(txtCod_Barra) < 7 Then
'            MsgBox "Codigo de Barra Inválido", vbExclamation + vbOKOnly, "Busca Datos Rollo"
'            txtCod_Barra.SetFocus
'            Exit Sub
        End If
        sPrefMaq = Right(txtCod_Barra, 7)
        sCodRollo = Left(sPrefMaq, 5)
        sPrefMaq = Right(sPrefMaq, 2)
    Case optDirecto
        txtPrefijo_Maquina = Trim(txtPrefijo_Maquina)
        txtCodigo_Rollo = Trim(txtCodigo_Rollo)
        If Len(txtCodigo_Rollo) <> txtCodigo_Rollo.MaxLength Or _
        Len(txtPrefijo_Maquina) <> txtPrefijo_Maquina.MaxLength Then
'            MsgBox "Codigo de Rollo Inválido", vbExclamation + vbOKOnly, "Busca Datos Rollo"
'            txtPrefijo_Maquina.SetFocus
'            Exit Sub
        End If
        sPrefMaq = txtPrefijo_Maquina
        sCodRollo = txtCodigo_Rollo
    End Select
    
    StrSql = "TX_MUESTRA_DATOS_BASICOS_ROLLO_TEJEDURIA '" & sPrefMaq & "','" & sCodRollo & "'"
    
    txtCod_OrdTra = ""
    txtDes_Tela = ""
    lblTipoOC = ""
    'lblKilos_Requeridos = ""
    lblAbr_Cliente = ""
    
    Set rstAux = CargarRecordSetDesconectado(StrSql, cCONNECT)
    If rstAux.RecordCount > 0 Then
        txtCod_Calidad.Text = rstAux!Cod_Calidad
        txtDes_Calidad.Text = rstAux!Des_Calidad
        txtCod_OrdTra = rstAux!Cod_Ordtra
        'lblTipoOC = rstAux!cod_ordtra
        txtDes_Tela = rstAux!DES_TELA
        txtKgs_Rollo = rstAux!Peso_Rollo_Actual
        txtUnidades = rstAux!Unidades_Rollo_Actual
        lblAbr_Cliente = rstAux!abr_cliente
        TxtNom_Cliente = rstAux!nom_cliente
    End If
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Mostrar Datos de Rollo"
End Sub

'Public Sub HabilitaCantidades()
'    'Entrada de un Proveedor
'
'    txtKgs_Rollo.Enabled = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'    txtUnidades.Enabled = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'
'    txtCod_Calidad.Visible = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'    txtDes_Calidad.Visible = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'    lblCalidad.Visible = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'    'cboCalidad.Visible = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
'    If sFlg_Devolucion_Rollos_Tejeduria = "S" Then
'        'cmdCapturarPeso.Enabled = False
'    End If
'End Sub

Public Sub LimpiaForm()
    txtCod_Barra.Text = ""
    txtPrefijo_Maquina = ""
    txtCodigo_Rollo = ""
    txtKgs_Rollo = "0"
    txtUnidades = "0"
    txtObservacion = ""
    txtDes_Calidad = ""
    txtCod_Calidad = ""
    StrSql = "SELECT TOP 1 Cod_Calidad FROM TX_CALIDAD_ROLLOS " & _
             "WHERE Flg_Default = '*'"
    'txtCod_Calidad = DevuelveCampo(StrSql, cCONNECT)
    'MostrarCalidad
    BuscaDatosRollo
End Sub

'Private Sub cboCalidad_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
'End Sub

Private Sub cmdCapturarPeso_Click()
    txtKgs_Rollo.Text = CapturaPeso
    
    'EjecBatch vRuta & "\TOL02.PIF"
    'Call LeerArchivo
    
    If RTrim(txtKgs_Rollo.Text) <> "0" Then
        If txtUnidades.Enabled Then
            If bExigeUnidades Then
               If txtUnidades.Enabled And Me.Visible Then
                  txtUnidades.SetFocus
                End If
            Else
                If txtCod_Calidad.Enabled = True Then
                    txtCod_Calidad.SetFocus
                End If
            End If
        Else
            FunctButt1.SetFocus
       End If
    End If
End Sub

Private Sub cmdCapturarPeso_GotFocus()
    cmdCapturarPeso_Click
End Sub

Private Sub Form_Load()
    'FillCalidad
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        GrabarDetalle
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub GrabarDetalle()
On Error GoTo Fin
Dim sTit As String
Dim sPrefMaq As String, sCodRollo As String
    
    sTit = "Guardar Detalle de Movimiento Devolucion de Rollos"
    If Not IsNumeric(txtKgs_Rollo) Then
        MsgBox "Se debe especificar el Peso", vbExclamation + vbOKOnly, sTit
        If txtKgs_Rollo.Enabled Then txtKgs_Rollo.SetFocus
        Exit Sub
    End If
    If CDbl(txtKgs_Rollo) <= 0 Then
        MsgBox "El Peso debe ser mayor a 0", vbExclamation + vbOKOnly, sTit
        If txtKgs_Rollo.Enabled Then txtKgs_Rollo.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtUnidades) Then
        MsgBox "Se deben especificar las unidades", vbExclamation + vbOKOnly, sTit
        If txtUnidades.Enabled Then txtUnidades.SetFocus
        Exit Sub
    End If
    If CDbl(txtUnidades) < 0 Then
        MsgBox "La Unidades deben ser 0 o mas", vbExclamation + vbOKOnly, sTit
        If txtUnidades.Enabled Then txtUnidades.SetFocus
        Exit Sub
    End If
    
    If optBarra.Value Then
        txtCod_Barra = Trim(txtCod_Barra)
        If Len(txtCod_Barra) < 7 Then
            MsgBox "Codigo de Barra Inválido", vbExclamation + vbOKOnly, sTit
            txtCod_Barra.SetFocus
            Exit Sub
        End If
        sPrefMaq = Right(txtCod_Barra, 7)
        sCodRollo = Left(sPrefMaq, 5)
        sPrefMaq = Right(sPrefMaq, 2)
    Else
        txtPrefijo_Maquina = Trim(txtPrefijo_Maquina)
        txtCodigo_Rollo = Trim(txtCodigo_Rollo)
        'Len(txtCodigo_Rollo) <> txtCodigo_Rollo.MaxLength
        If Len(txtPrefijo_Maquina) <> txtPrefijo_Maquina.MaxLength Then
            MsgBox "Codigo de Barra Inválido", vbExclamation + vbOKOnly, sTit
            txtCod_Barra.SetFocus
            Exit Sub
        End If
        
        If txtCod_OrdTra = "" Or txtDes_Tela = "" Then
            MsgBox "No se encontraron los datos de la O.T.", vbExclamation + vbOKOnly, sTit
            txtPrefijo_Maquina.SetFocus
            Exit Sub
        End If
        
        sPrefMaq = txtPrefijo_Maquina
        sCodRollo = txtCodigo_Rollo
    End If
    
    StrSql = "EXEC TX_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS '" & sAccion & "', '" & _
    sCod_Almacen & "', '" & sNum_MovStk & "', '" & sNum_Secuencia & "', '" & _
    sPrefMaq & "', '" & sCodRollo & "', '" & txtKgs_Rollo & "', '" & _
    txtUnidades & "', '" & txtObservacion & "', '', '" & txtCod_Calidad & "'"
    
    ExecuteSQL cCONNECT, StrSql
    
    LimpiaForm
    If optBarra.Value Then
        txtCod_Barra.SetFocus
    Else
        txtPrefijo_Maquina.SetFocus
    End If
    Beep
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
    txtCod_Barra = ""
    If optBarra.Value Then
        txtCod_Barra.SetFocus
    Else
        txtPrefijo_Maquina.SetFocus
    End If
    Beep
    Beep
    Beep
End Sub

Private Sub txtCod_Calidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MostrarCalidad
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCodigo_Rollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaDatosRollo
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtKgs_Rollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPrefijo_Maquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaDatosRollo
        SendKeys "{TAB}"
    End If
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

Private Sub txtCod_Barra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtCod_Barra <> "" Then
            BuscaDatosRollo
            'cmdCapturarPeso.SetFocus
        Else
            optDirecto = True
            SendKeys "{TAB}"
        End If
    End If
End Sub

'Private Sub FillCalidad()
'On Error GoTo Fin
'Dim rstAux As ADODB.Recordset
'    StrSql = "SELECT Cod_Calidad, Des_Calidad, Flg_Default " & _
'             "FROM TX_CALIDAD_ROLLOS"
'    Set rstAux = CargarRecordSetDesconectado(StrSql, cCONNECT)
'    With rstAux
'    If .RecordCount > 0 Then .MoveFirst
'    Do Until .EOF
'        cboCalidad.AddItem !Cod_Calidad & " " & !Des_Calidad
'        If !Flg_Default = "*" Then cboCalidad.ListIndex = cboCalidad.ListCount - 1
'        .MoveNext
'    Loop
'    End With
'Exit Sub
'Fin:
'    MsgBox Err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
'End Sub

Private Sub LeerArchivo()
Dim Archivo$
Dim f, a As Variant

Archivo = vRuta & "\TEXTO.TXT"
If Dir(Archivo) <> "" Then
    f = FreeFile
    Open Archivo For Input As #f
        Line Input #f, a
        txtKgs_Rollo = Trim(a)
        Close #f
End If
End Sub

Private Sub MostrarCalidad()
On Error GoTo Fin
    StrSql = "SELECT Des_Calidad FROM TX_CALIDAD_ROLLOS " & _
             "WHERE Cod_Calidad = '" & txtCod_Calidad & "'"
    txtDes_Calidad = DevuelveCampo(StrSql, cCONNECT)
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Mostrar Calidad"
End Sub


Private Function RetornaTipoFamtela(Cod_Tela As String) As String
Dim cn As New ADODB.Connection
Dim CMD As New ADODB.Command
Dim Rs As ADODB.Recordset

Dim Cod_TipFamTela As String


cn.Open cCONNECT
       
Set Rs = New ADODB.Recordset
With CMD
     Set .ActiveConnection = cn
     .CommandType = adCmdStoredProc
     .CommandText = "SM_ENCUENTRA_TIPOFAMTELA "
     
     'Pasa los parametros para el procedimiento UP_GENERA_COTIZACION
     .Parameters.Append .CreateParameter("@cod_TELA", adVarChar, adParamInput, 8, Cod_Tela)
     .Parameters.Append .CreateParameter("@cod_TIPFAMTELA", adVarChar, adParamOutput, 1, Cod_TipFamTela)
     .Parameters.Append .CreateParameter("@Select", adVarChar, adParamInput, 1, "N")
   .Execute
 End With
RetornaTipoFamtela = CMD.Parameters.Item("@Cod_tipfamtela").Value

cn.Close
Set cn = Nothing

End Function

Public Sub HabilitaCantidades()
    'Entrada de un Proveedor
    txtKgs_Rollo.Enabled = (sCod_ClaMov = "E" And sFlg_Devolucion_Rollos_Tejeduria = "S")
    txtUnidades.Enabled = (sCod_ClaMov = "E" And sFlg_Devolucion_Rollos_Tejeduria = "S")
    
    txtCod_Calidad.Visible = (sCod_ClaMov = "E" And sFlg_Devolucion_Rollos_Tejeduria = "S")
    txtDes_Calidad.Visible = (sCod_ClaMov = "E" And sFlg_Devolucion_Rollos_Tejeduria = "S")
    lblCalidad.Visible = (sCod_ClaMov = "E" And sFlg_Devolucion_Rollos_Tejeduria = "S")
    'cboCalidad.Visible = (sCod_ClaMov = "E" And sCod_TipAnx = "P")
End Sub

