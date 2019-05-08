VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmManEstCliTem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estilo Cliente Temporada"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraListado 
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   8460
      Begin GridEX20.GridEX DGridLista 
         Height          =   2640
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   4657
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "FrmManEstCliTem.frx":0000
         Column(2)       =   "FrmManEstCliTem.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmManEstCliTem.frx":016C
         FormatStyle(2)  =   "FrmManEstCliTem.frx":02A4
         FormatStyle(3)  =   "FrmManEstCliTem.frx":0354
         FormatStyle(4)  =   "FrmManEstCliTem.frx":0408
         FormatStyle(5)  =   "FrmManEstCliTem.frx":04E0
         FormatStyle(6)  =   "FrmManEstCliTem.frx":0598
         FormatStyle(7)  =   "FrmManEstCliTem.frx":0678
         FormatStyle(8)  =   "FrmManEstCliTem.frx":0724
         ImageCount      =   0
         PrinterProperties=   "FrmManEstCliTem.frx":07D4
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   6660
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmManEstCliTem.frx":09AC
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "FrmManEstCliTem.frx":0B1E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmManEstCliTem.frx":0C90
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1470
         Picture         =   "FrmManEstCliTem.frx":0E02
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MFEstCli 
      Height          =   540
      Left            =   2400
      TabIndex        =   14
      Top             =   6480
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManEstCliTem.frx":0F74
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame fraDetalles 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   8340
      Begin VB.CheckBox chk_Excel 
         Alignment       =   1  'Right Justify
         Caption         =   "Visualiza en Excel"
         Height          =   195
         Left            =   3960
         TabIndex        =   28
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtnum_estprorea 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "1"
         Top             =   1620
         Width           =   615
      End
      Begin VB.Frame FrmOpcionales 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   480
         TabIndex        =   20
         Top             =   1850
         Width           =   7575
         Begin VB.ComboBox cboflg_status 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   120
            Width           =   2055
         End
         Begin VB.ComboBox cboCod_MotPrePro 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txtComentario 
            Height          =   735
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   6255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   0
            TabIndex        =   26
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mot. Pre Producc"
            Height          =   195
            Left            =   3480
            TabIndex        =   25
            Top             =   165
            Width           =   1245
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Comentario"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   525
            Width           =   795
         End
      End
      Begin VB.TextBox TxtTela 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1290
         Width           =   6255
      End
      Begin VB.TextBox txtNom_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtAbr_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDes_EstCli 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txtCod_EstCli 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Est.Prop"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1665
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1335
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1005
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Cliente"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   645
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmManEstCliTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim opcion As Integer
Dim sTipo As String
Dim StrSQL As String
Dim Rs_Lista As ADODB.Recordset
Public varCod_Cliente, varCod_TemCli As String
Dim Cod_EstCliProv As String
Dim VarCod_EstCliMod As Boolean 'Identifica si se va a modificar o no el codestcli
'Esta es solo una variable opcional
Public varCod_EstCli As String
Dim TipoFab As Integer

'Private Sub cmdFirst_Click()
'    If Not Rs_Lista.BOF Then
'        Rs_Lista.MoveFirst
'    End If
'End Sub
'
'Private Sub cmdLast_Click()
'    If Not Rs_Lista.EOF Then
'        Rs_Lista.MoveLast
'    End If
'End Sub
'
'Private Sub cmdNext_Click()
'    If Not Rs_Lista.EOF Then
'        Rs_Lista.MoveNext
'        If Rs_Lista.EOF Then
'            Rs_Lista.MoveLast
'        End If
'    End If
'End Sub
'
'Private Sub cmdPrevious_Click()
'    If Not Rs_Lista.BOF Then
'        Rs_Lista.MovePrevious
'        If Rs_Lista.BOF Then
'            Rs_Lista.MoveFirst
'        End If
'    End If
'End Sub

Public Sub RECARGA_LISTA()
    'Set Rs_Lista = Nothing
    Call CARGA_LISTA
End Sub

Public Sub CARGA_LISTA()
    Dim StrSQL As String
'    Set Rs_Lista = New ADODB.Recordset
'    Rs_Lista.ActiveConnection = cCONNECT
'    Rs_Lista.CursorType = adOpenStatic
'    Rs_Lista.CursorLocation = adUseClient
'    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_ESTCLITEM '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "'"
    'Rs_Lista.Open StrSQL
    Set DGridLista.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)
    'Set DGridLista.DataSource = Rs_Lista

    If DGridLista.RowCount > 0 Then
        HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MFEstCli, "ADICIONAR"
        Call LIMPIA_DATOS
    End If
End Sub

Public Sub Carga_Datos()
    If DGridLista.RowCount > 0 Then
        txtCod_EstCli = Trim(DGridLista.Value(DGridLista.Columns("Cod_EstCli").Index))
        txtDes_EstCli = Trim(DGridLista.Value(DGridLista.Columns("Des_EstCli").Index))
        txtnum_estprorea = Trim(DGridLista.Value(DGridLista.Columns("num_estprorea").Index))
        
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("flg_status").Index), 2, cboflg_status)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_MotPrePro").Index), 2, cboCod_MotPrePro)
        txtComentario = Trim(DGridLista.Value(DGridLista.Columns("Comentario").Index))
        Me.TxtTela = Trim(DGridLista.Value(DGridLista.Columns("Tela").Index))
        chk_Excel = IIf(Trim(DGridLista.Value(DGridLista.Columns("Flg_Excel").Index)) = True, 1, 0)
    End If
End Sub
Public Sub HABILITA_DATOS()
    txtCod_EstCli.Enabled = True
    txtDes_EstCli.Enabled = True
    txtnum_estprorea.Enabled = True
    cboflg_status.Enabled = True
    cboCod_MotPrePro.Enabled = True
    txtComentario.Enabled = True
    Me.TxtTela.Enabled = True
End Sub
Public Sub DESABILITA_DATOS()
    txtCod_EstCli.Enabled = False
    txtDes_EstCli.Enabled = False
    txtnum_estprorea.Enabled = False
    cboflg_status.Enabled = False
    cboCod_MotPrePro.Enabled = False
    txtComentario.Enabled = False
    Me.TxtTela.Enabled = False
End Sub

Public Sub LIMPIA_DATOS()
    txtCod_EstCli.Text = ""
    txtDes_EstCli.Text = ""
    txtnum_estprorea.Text = ""
    cboflg_status.ListIndex = -1
    cboCod_MotPrePro.ListIndex = -1
    txtComentario.Text = ""
    Me.TxtTela = ""
End Sub

Public Sub CARGA_COMBOS()
    
    'Combo Flag Estatus
    StrSQL = "SELECT des_status + space(100) + flg_status  FROM TG_StaDes"
    Call LlenaCombo(cboflg_status, StrSQL, cCONNECT)
    
    'Combo Motivo Preproduccion
    StrSQL = "SELECT des_motprepro + space(100) + cod_motprepro  FROM TG_MotPrePro"
    Call LlenaCombo(cboCod_MotPrePro, StrSQL, cCONNECT)
    
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo = "I" Then
        If Trim(txtCod_EstCli.Text) = "" Then
            Call MsgBox("Sirvase ingresar un codigo de estilo", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        If Val(txtnum_estprorea.Text) < 1 Then
            Call MsgBox("El nro de estilos propios debe ser mayor a 0, sirvase verificar", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(txtDes_EstCli.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        StrSQL = "SELECT * FROM TG_ESTCLITEM WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & txtCod_EstCli.Text & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            Call MsgBox("El código ingresado ya existe. Sirvase verificar", vbCritical)
            txtCod_EstCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    If sTipo = "U" Then
        If Val(txtnum_estprorea.Text) < 1 Then
            Call MsgBox("El nro de estilos propios debe ser mayor a 0, sirvase verificar", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(txtDes_EstCli.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    
    
End Function

Public Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
       
        If sTipo = "I" Then
            Cod_EstCliProv = Trim(txtCod_EstCli.Text)
        End If
       
        'Esta es la sentencia que realizara el salvado de datos
        StrSQL = "UP_MAN_ESTCLITEM " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        Cod_EstCliProv & "','" & _
        Trim(txtDes_EstCli.Text) & "'," & _
        txtnum_estprorea & ",'" & _
        Right(cboflg_status.Text, 1) & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        Trim(txtComentario) & "','" & _
        Me.TxtTela & "','" & _
        IIf(VarCod_EstCliMod, Trim(Me.txtCod_EstCli.Text), "") & "'," & chk_Excel
        
        Con.Execute StrSQL
        
        If VarCod_EstCliMod Then
            varCod_EstCli = Trim(Me.txtCod_EstCli.Text)
        End If
        
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call DESABILITA_DATOS
    Call LIMPIA_DATOS

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
    StrSQL = "SELECT Cod_EstCli FROM tg_estcliest WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & txtCod_EstCli.Text & "'"

    If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
           
        'Esta es la sentencia que realiza la eliminacion del Registro
        StrSQL = "UP_MAN_ESTCLITEM " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        txtCod_EstCli & "','" & _
        txtDes_EstCli & "'," & _
        txtnum_estprorea & ",'" & _
        Right(cboflg_status.Text, 1) & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        txtComentario & "'"
        
        Con.Execute StrSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIA_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Private Sub DGridLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If DGridLista.RowCount > 0 Then
        Call Carga_Datos
    End If
End Sub

'Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If Rs_Lista.State <> 1 Then
'        Exit Sub
'    End If
'    If DGridLista.RowCount > 0 Then
'        Call Carga_Datos
'    End If
'End Sub

Private Sub Form_Load()
On Error GoTo hand
    Call FormSet(Me)
    Call CARGA_COMBOS
    Call DESABILITA_DATOS
    'Call FormateaGrid(DGridLista)
    MFEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    TipoFab = DevuelveCampo("select tip_fabrica from tg_control", cCONNECT)
    If TipoFab = 1 Then
        FrmOpcionales.Visible = True
    Else
        FrmOpcionales.Visible = False
    End If
Exit Sub
hand:
ErrorHandler Err, "Form_Load()"
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If Rs_Lista.RecordCount > 0 Then
'        With oParent
'            .Valor = DGridLista.Columns(0)
'        End With
'    End If
'End Sub

Public Sub MFEstCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIA_DATOS
            HABILITA_DATOS
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            txtnum_estprorea = "1"
            
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            StrSQL = "SELECT COUNT(*) FROM TG_ESTCLIEST WHERE Cod_Cliente = '" & Me.varCod_Cliente & "' AND Cod_TemCli = '" & Me.varCod_TemCli & "' AND Cod_EstCli = '" & DGridLista.Value(DGridLista.Columns("Cod_EstCli").Index) & "'"
            If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
                txtCod_EstCli.Enabled = False
                'txtDes_EstCli.SetFocus
            Else
                'txtCod_EstCli.SetFocus
                VarCod_EstCliMod = True
            End If
            'Aqui guardamos en esta varialbe temporal el codigo antiguo del estilo
            Cod_EstCliProv = DGridLista.Value(DGridLista.Columns("Cod_EstCli").Index)
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            Eliminar = MsgBox("Usted desea eliminar el registro seleccionado", vbExclamation + vbYesNo)
            If Eliminar = vbYes Then
                Call ELIMINAR_DATOS
                Call RECARGA_LISTA
            Else
                Exit Sub
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                If sTipo = "I" Then Call Asigna_EP
                Call RECARGA_LISTA
                sTipo = ""
                HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                Call Carga_Datos
            End If
        Case "DESHACER"
            DESABILITA_DATOS
            sTipo = ""
            LIMPIA_DATOS
            Call Carga_Datos
            HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Private Sub txtCod_EstCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
End Sub

Private Sub txtDes_EstCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
End Sub

Private Sub txtnum_estprorea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
    Call SoloNumeros(txtnum_estprorea, KeyAscii, False, 0, 2)
End Sub

Private Sub txtnum_estprorea_LostFocus()
    If Trim(txtnum_estprorea.Text) = "" Then
        txtnum_estprorea.Text = "1"
    End If
End Sub

Private Sub TxtTela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
End Sub

Sub Asigna_EP()
Load FrmAddEstiloPropio
FrmAddEstiloPropio.vCod_Cliente = Me.varCod_Cliente
FrmAddEstiloPropio.txtAbr_Cliente = DevuelveCampo("select abr_cliente from tg_cliente where cod_cliente='" & Me.varCod_Cliente & "'", cCONNECT)
FrmAddEstiloPropio.TxtNom_Cliente = DevuelveCampo("select nom_cliente from tg_cliente where cod_cliente='" & Me.varCod_Cliente & "'", cCONNECT)
FrmAddEstiloPropio.vcod_TemCli = Me.varCod_TemCli
FrmAddEstiloPropio.txtCod_TemCli = Me.varCod_TemCli
FrmAddEstiloPropio.TxtDes_TemCli = DevuelveCampo("select nom_temcli from TG_TemCli where cod_cliente ='" & varCod_Cliente & "' and cod_temcli='" & Me.varCod_TemCli & "'", cCONNECT)
FrmAddEstiloPropio.vCod_EstCli = Cod_EstCliProv
FrmAddEstiloPropio.txtCod_EstCli.Text = Cod_EstCliProv
FrmAddEstiloPropio.txtDes_EstCli.Text = DevuelveCampo("select Des_EstCli from tg_estclitem where cod_cliente ='" & varCod_Cliente & "' and cod_temcli='" & Me.varCod_TemCli & "' and cod_estcli='" & Cod_EstCliProv & "'", cCONNECT)
FrmAddEstiloPropio.sDes_Estilo = DevuelveCampo("select Des_EstCli from tg_estclitem where cod_cliente ='" & varCod_Cliente & "' and cod_temcli='" & Me.varCod_TemCli & "' and cod_estcli='" & Cod_EstCliProv & "'", cCONNECT)
FrmAddEstiloPropio.vDes_Tela = DevuelveCampo("select Des_Tela from tg_estclitem where cod_cliente ='" & varCod_Cliente & "' and cod_temcli='" & Me.varCod_TemCli & "' and cod_estcli='" & Cod_EstCliProv & "'", cCONNECT)
FrmAddEstiloPropio.Show vbModal
Set FrmAddEstiloPropio = Nothing
End Sub
