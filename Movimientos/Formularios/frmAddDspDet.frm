VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAddDspDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Despacho"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6390
      Begin VB.CheckBox chkRepesar 
         Caption         =   "Repesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   1155
      End
      Begin VB.TextBox txtCod_OrdTra 
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
         Left            =   840
         TabIndex        =   2
         Top             =   750
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblUnid 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   4095
         TabIndex        =   26
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label8 
         Caption         =   "Unidades "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   25
         Top             =   285
         Width           =   900
      End
      Begin VB.Label lblKgs 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   1770
         TabIndex        =   24
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "Kgs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1410
         TabIndex        =   23
         Top             =   285
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "Rollos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5010
         TabIndex        =   22
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblRollos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   5640
         TabIndex        =   21
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblCod_Fam 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3810
         TabIndex        =   18
         Top             =   1110
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblDes_Fam 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   14
         Top             =   1110
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
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
         Left            =   3075
         TabIndex        =   13
         Top             =   1125
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblNom_Cliente 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   12
         Top             =   1110
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
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
         Left            =   135
         TabIndex        =   11
         Top             =   1140
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "O.T."
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
         Left            =   150
         TabIndex        =   1
         Top             =   765
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2025
      TabIndex        =   10
      Top             =   2820
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAddDspDet.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   75
      TabIndex        =   15
      Top             =   795
      Width           =   6375
      Begin VB.TextBox txtUnidades 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   9
         Text            =   "0"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
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
         Left            =   2535
         TabIndex        =   8
         Text            =   "0"
         Top             =   1095
         Width           =   1380
      End
      Begin VB.CommandButton cmdCapturarPeso 
         Caption         =   "Capturar Peso"
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   1095
         Width           =   1440
      End
      Begin VB.TextBox txtCodigo_Rollo 
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
         Left            =   3195
         MaxLength       =   7
         TabIndex        =   6
         Top             =   705
         Width           =   1170
      End
      Begin VB.TextBox txtPrefijo_Maquina 
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
         Left            =   2535
         MaxLength       =   2
         TabIndex        =   5
         Top             =   705
         Width           =   615
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
         Left            =   2535
         TabIndex        =   4
         Top             =   315
         Width           =   3420
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
         Left            =   915
         TabIndex        =   3
         Top             =   765
         Width           =   1455
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
         Left            =   915
         TabIndex        =   20
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Unidades"
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
         Left            =   1185
         TabIndex        =   17
         Top             =   1485
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblPeso 
         Caption         =   "Peso"
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
         Left            =   1200
         TabIndex        =   16
         Top             =   1110
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmAddDspDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String, sFec_MovStk As String, iNum_Despacho As Integer, _
       sCod_Cliente_Tex As String, Codigo As String, Descripcion As String, TipoAdd As String, sAccion As String, sCod_TipMov As String, sNum_MovStk As String
Dim StrSql As String, rstAux As ADODB.Recordset
Public StipoStore As String
Public sNum_Secuencia As String

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
        'If Len(txtCod_Barra) > 10 Then
        '    sRolloAdd = Left(txtCod_Barra, 1)
        '    sPrefMaq = Right(txtCod_Barra, 7)
        '    sCodRollo = Left(sPrefMaq, 5)
        '    sPrefMaq = Right(sPrefMaq, 2)
        'End If
        
        'If Len(txtCod_Barra) >= 11 Or Len(txtCod_Barra) >= 10 Then
            sCodRollo = Left(txtCod_Barra, 5)
            sPrefMaq = Mid(txtCod_Barra, 6, 2)
'            sCodRollo = Mid(txtCod_Barra, 3, 7)
        'End If
        
        
        
'        If Not IsNumeric(sRolloAdd) Then
'            sCodRollo = sCodRollo & sRolloAdd
'        End If
   ' If Len(txtCod_Barra) > 10 Then
   '     If IsNumeric(Mid(txtCod_Barra, 1, 1)) Then
   '         sSufRollo = ""
   '     Else
   '         sSufRollo = Mid(txtCod_Barra, 6, 2)
   '         If Not IsNumeric(Mid(txtCod_Barra, 2, 1)) Then
   '             sSufRollo = sSufRollo
   '             '& Mid(txtCod_Barra, 6, 2)
   '         End If
       ' End If
    'End If
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
    
    
    If StipoStore = "01" Then
    
    StrSql = "EXEC LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS 'I', '" & sCod_Almacen & _
    "', '" & sNum_MovStk & "', '" & sNum_Secuencia & "','" & sPrefMaq & "', '" & sCodRollo & "', " & txtPeso & ", " & _
    txtUnidades & ", '','','','" & vusu & "'"
    
    End If
    
    If StipoStore = "02" Then
    
        StrSql = "LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS_DESPACHO_PARTIDA 'I', '" & sCod_Almacen & _
        "', '" & sNum_MovStk & "',' " & sNum_Secuencia & "', '" & _
        sPrefMaq & "', '" & sCodRollo & "', " & txtPeso & ", " & _
        txtUnidades & ", '" & IIf(chkRepesar = 1, "S", "N") & "', '" & vusu & "' "
        
    End If
    
    If StipoStore = "03" Then
    
        StrSql = "LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS_OTROS_MOVS 'I', '" & sCod_Almacen & _
        "', '" & sNum_MovStk & "',' " & sNum_Secuencia & "', '" & _
        sPrefMaq & "', '" & sCodRollo & "', " & txtPeso & ", " & _
        txtUnidades & ", '" & IIf(chkRepesar = 1, "S", "N") & "', '" & vusu & "' "
        
    End If
    


    ExecuteSQL cConnect, StrSql
    
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

'Private Sub BuscaOT()
'On Error GoTo Fin
'Dim iCol As Long, sTit As String
'
'    sTit = "Busqueda de O.T."
'
'    txtCod_OrdTra = Trim(txtCod_OrdTra)
'
'    StrSQL = "SELECT Cod_OrdTra, Fec_Generacion, Ser_OrdComp + '-' + Cod_OrdComp " & _
'             "+ '-' + Sec_OrdComp AS OrdComp FROM TX_ORDTRA_TEJEDURIA " & _
'             "WHERE Cod_Cliente_Tex = '" & sCod_Cliente_Tex & "' " & _
'             "AND   Cod_OrdTra like '%" & txtCod_OrdTra & "%' " & _
'             "AND   Flg_Status = 'E' ORDER BY Cod_OrdTra"
'
'    txtCod_OrdTra = ""
'
'    With frmBusqGeneral
'        Set .oParent = Me
'        .sQuery = StrSQL
'        .Cargar_Datos
'        Codigo = ".."
'        Set rstAux = .DGridLista.ADORecordset
'
'        .DGridLista.Columns("Cod_OrdTra").Caption = "O.T."
'        .DGridLista.Columns("Cod_OrdTra").Width = 700
'        .DGridLista.Columns("Fec_Generacion").Caption = "Fec.Gen."
'        .DGridLista.Columns("Fec_Generacion").Width = 1200
'        .DGridLista.Columns("OrdComp").Caption = "Orden"
'        .DGridLista.Columns("OrdComp").Width = 1500
'
'        If rstAux.RecordCount > 1 Then .Show vbModal
'
'        If Codigo <> "" And rstAux.RecordCount > 0 Then
'            txtCod_OrdTra = Trim(rstAux!Cod_OrdTra)
'        End If
'    End With
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'On Error Resume Next
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
'    "Búsqueda de Cliente"
'End Sub

Private Sub DatosOT()
On Error GoTo Fin
Dim sTit As String, rstAux As ADODB.Recordset
    sTit = "Mostrar Datos de O.T."
    
    StrSql = "SELECT a.Cod_Cliente_Tex, a.Cod_FamGrupo, b.Nom_Cliente, c.Des_FamGrupo " & _
             "FROM   TX_ORDTRA_TEJEDURIA a, TX_CLIENTE b, TX_FAMGRUPO_TEJEDURIA c " & _
             "WHERE  Cod_OrdTra = '" & txtCod_OrdTra & "' " & _
             "AND    b.Cod_Cliente_Tex = a.Cod_Cliente_Tex " & _
             "AND    c.Cod_FamGrupo = a.Cod_FamGrupo"
    
    Set rstAux = CargarRecordSetDesconectado(StrSql, cConnect)
    lblNom_Cliente = ""
    lblDes_Fam = ""
    lblCod_Fam = ""
    With rstAux
    If .RecordCount > 0 Then
        .MoveFirst
        lblNom_Cliente = !nom_cliente
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
            'BuscaOT
        Else
            txtCod_OrdTra = Format(txtCod_OrdTra, "00000")
        End If
        DatosOT
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCodigo_Rollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPrefijo_Maquina_KeyPress(KeyAscii As Integer)
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
Dim mRs As ADODB.Recordset
    txtCod_Barra = ""
    txtPrefijo_Maquina = ""
    txtCodigo_Rollo = ""
    txtPeso = 0
    
    'StrSql = "SELECT COUNT(Num_Rollo) FROM TX_ORDTRA_TEJEDURIA_ROLLOS " & _
    '         "WHERE Cod_Almacen_Despacho = '" & sCod_Almacen & "' " & _
    '         "AND   Fec_MovStk_Despacho = '" & sFec_MovStk & "' " & _
    '         "AND   Num_Despacho = '" & iNum_Despacho & "'"
    
    StrSql = " SM_RESUMEN_DESPACHO '" & sCod_Almacen & "','" & _
               sFec_MovStk & "','" & iNum_Despacho & "'"
    Set mRs = GetRecordset(cConnect, StrSql)
    
    If Not mRs.EOF Then
        lblKgs = FixNulos(mRs!Kgs, vbDouble)
        lblUnid = FixNulos(mRs!Unidades, vbDouble)
        lblRollos = FixNulos(mRs!Rollos, vbDouble)
    End If
    
    mRs.Close
    Set mRs = Nothing
Exit Sub
errx:
    If Not mRs Is Nothing Then
        mRs.Close
    End If
    Set mRs = Nothing
    
    MsgBox err.Description, vbCritical + vbOKOnly, StrSql
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

