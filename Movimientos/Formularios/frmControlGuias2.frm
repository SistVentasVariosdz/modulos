VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmControlGuias 
   Caption         =   "Control de Guias"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   Icon            =   "frmControlGuias2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   8070
      TabIndex        =   22
      Top             =   8100
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~ASIGNA~True~True~&Asigna Guias~0~0~1~~0~False~False~&Asigna Guias~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   135
      TabIndex        =   15
      Top             =   6780
      Width           =   11175
      Begin VB.ComboBox CmbEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TxtNom_Usuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton CmdUsuario 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxtCod_Usuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtCod_Guia 
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtSer_Guia 
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   6480
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nº de Guia :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1710
      TabIndex        =   10
      Top             =   8100
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   953
      Custom          =   $"frmControlGuias2.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame Frame2 
      Height          =   5625
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   11190
      Begin GridEX20.GridEX gexList 
         Height          =   5280
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   9313
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigatorString=   "Registro:|de"
         HoldSortSettings=   -1  'True
         GridLineStyle   =   2
         ColumnAutoResize=   -1  'True
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         Options         =   8
         RecordsetType   =   1
         GroupByBoxInfoText=   ""
         AllowEdit       =   0   'False
         BorderStyle     =   2
         GroupByBoxVisible=   0   'False
         ImageCount      =   3
         ImagePicture1   =   "frmControlGuias2.frx":046E
         ImagePicture2   =   "frmControlGuias2.frx":0580
         ImagePicture3   =   "frmControlGuias2.frx":089A
         RowHeaders      =   -1  'True
         DataMode        =   1
         HeaderFontName  =   "Tahoma"
         FontName        =   "Tahoma"
         GridLines       =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         SortKeysCount   =   1
         SortKey(1)      =   "frmControlGuias2.frx":0BB4
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmControlGuias2.frx":0C1C
         FormatStyle(2)  =   "frmControlGuias2.frx":0CFC
         FormatStyle(3)  =   "frmControlGuias2.frx":0E24
         FormatStyle(4)  =   "frmControlGuias2.frx":0ED4
         FormatStyle(5)  =   "frmControlGuias2.frx":0F88
         FormatStyle(6)  =   "frmControlGuias2.frx":1060
         ImageCount      =   3
         ImagePicture(1) =   "frmControlGuias2.frx":1118
         ImagePicture(2) =   "frmControlGuias2.frx":122A
         ImagePicture(3) =   "frmControlGuias2.frx":1544
         PrinterProperties=   "frmControlGuias2.frx":185E
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   11175
      Begin VB.TextBox txtSerie 
         Height          =   315
         Left            =   7995
         TabIndex        =   24
         Top             =   285
         Width           =   615
      End
      Begin VB.OptionButton optSerie 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   6810
         TabIndex        =   23
         Top             =   345
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   9870
         TabIndex        =   5
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox TxtCodUsuario 
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton CmdUsuarioB 
         Caption         =   "..."
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox TxtCodGuia 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   320
         Width           =   1335
      End
      Begin VB.TextBox TxtSerGuia 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   320
         Width           =   615
      End
      Begin VB.OptionButton OptUsu 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   345
         Width           =   975
      End
      Begin VB.OptionButton OptGuia 
         Caption         =   "Nº Guia :"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   315
         Width           =   135
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   630
      Top             =   8100
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmControlGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String

Private Sub CmdUsuario_Click()
Set frmBusqGeneral2.oParent = Me
frmBusqGeneral2.sQuery = "select cod_usuario as Codigo ,nom_usuario as Descripcion from seguridad..seg_usuarios order by 2"
frmBusqGeneral2.CARGAR_DATOS
frmBusqGeneral2.Show 1
If Codigo <> "" Then
    Me.TxtCod_Usuario.Text = Codigo
    Me.TxtNom_Usuario.Text = Descripcion
End If
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub CmdUsuarioB_Click()
Set frmBusqGeneral2.oParent = Me
frmBusqGeneral2.sQuery = "select cod_usuario as Codigo ,nom_usuario as Descripcion from seguridad..seg_usuarios order by 2"
frmBusqGeneral2.CARGAR_DATOS
frmBusqGeneral2.Show 1
If Codigo <> "" Then
    Me.TxtCodUsuario.Text = Codigo
End If
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub Form_Load()
    OptGuia_Click
    CARGA_GRID
    CmbEstado.AddItem "Entregado", 0
    CmbEstado.AddItem "Devuelto", 1
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
If OptGuia.Value = True Then
    If Trim(TxtSerGuia) = "" Or Trim(TxtCodGuia.Text) = "" Then Exit Sub
End If
If OptUsu.Value = True Then
    If Trim(TxtCodUsuario.Text) = "" Then Exit Sub
End If
If optSerie.Value Then
    If Trim(txtSerie.Text) = "" Then Exit Sub
End If
    
    CARGA_GRID
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    frmAsignaGuias.Show 1
End Sub

Private Sub gexList_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
If gexList.RowCount = 0 And ActionName <> "SALIR" Then Exit Sub
Select Case ActionName
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Habilita True
    Case "GRABAR"
        If Trim(TxtCod_Usuario.Text) = "" Then
            MsgBox "Ingrese el usuario", vbInformation, Me.Caption
            TxtCod_Usuario.SetFocus
            Exit Sub
        End If
        If Trim(CmbEstado.Text) = "" Then
            MsgBox "Ingrese el estado", vbInformation, Me.Caption
            CmbEstado.SetFocus
            Exit Sub
        End If
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        SALVAR_DATOS
        Habilita False
        CARGA_GRID
'        Datos "V", False
'        varNum_Mov = ""
    Case "DESHACER"
'        sTipo = ""
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Habilita False
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub

Private Sub OptGuia_Click()
    TxtSerGuia.Enabled = True
    TxtCodGuia.Enabled = True
    TxtCodUsuario.Enabled = False
    TxtCodUsuario.Text = ""
    CmdUsuarioB.Enabled = False
End Sub

Private Sub optSerie_Click()
    txtSerie.SetFocus
End Sub

Private Sub OptUsu_Click()
    TxtSerGuia.Enabled = False
    TxtCodGuia.Enabled = False
    TxtSerGuia.Text = ""
    TxtCodGuia.Text = ""
    TxtCodUsuario.Enabled = True
    CmdUsuarioB.Enabled = True
    TxtCodUsuario.SetFocus
End Sub

Private Sub TxtCodGuia_Change()
If Len(TxtCodGuia.Text) = 6 Then FunctButt1.SetFocus
End Sub

Private Sub TxtCodGuia_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtCodGuia, KeyAscii, False, 0, 6)
End Sub

Private Sub TxtCodUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("cod_usuario", "seg_usuarios", TxtCodUsuario.Text, cSEGURIDAD) Then
        FunctButt1.SetFocus
    Else
        MsgBox "Codigo no existe", vbInformation, Me.Caption
    End If
End If
End Sub

Private Sub TxtSerGuia_Change()
    If Len(TxtSerGuia.Text) = 3 Then TxtCodGuia.SetFocus
End Sub

Private Sub TxtSerGuia_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtSerGuia, KeyAscii, False, 0, 3)
    If KeyAscii = 13 Then
        TxtCodGuia.SetFocus
    End If
End Sub

Sub CARGA_GRID()
Dim Rs_Carga As New ADODB.Recordset
Dim sSQl As String

On Error GoTo Cargar_DatosErr
Rs_Carga.ActiveConnection = cConnect
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.CursorLocation = adUseClient
Rs_Carga.LockType = adLockReadOnly
If OptGuia.Value = True Then
    sSQl = "exec UP_SEL_USERGUIA '1','" & TxtSerGuia.Text & "','" & TxtCodGuia.Text & "','" & Trim(TxtCodUsuario.Text) & "'"
Else
    If OptUsu.Value Then
        sSQl = "exec UP_SEL_USERGUIA '2','" & TxtSerGuia.Text & "','" & TxtCodGuia.Text & "','" & Trim(TxtCodUsuario.Text) & "'"
    Else
        sSQl = "exec UP_SEL_USERGUIA '3','" & txtSerie.Text & "','" & TxtCodGuia.Text & "','" & Trim(TxtCodUsuario.Text) & "'"
    End If
End If

Rs_Carga.Open sSQl

Set gexList.ADORecordset = Rs_Carga
'ConfiguraGrid
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "CARGA_GRID"
End Sub

Sub CARGA_DATOS()
On Error GoTo Cargar_DatosErr

If gexList.RowCount = 0 Then
    TxtSer_Guia.Text = ""
    TxtCod_Guia.Text = ""
    TxtCod_Usuario = ""
    TxtNom_Usuario.Text = ""
    Exit Sub
End If

TxtSer_Guia.Text = gexList.Value(gexList.Columns("serie").Index)
TxtCod_Guia.Text = gexList.Value(gexList.Columns("numero").Index)
TxtCod_Usuario = gexList.Value(gexList.Columns("Usuario").Index)
TxtNom_Usuario.Text = DevuelveCampo("select isnull(nom_usuario,'') from seg_usuarios where cod_usuario='" & TxtCod_Usuario.Text & "'", cSEGURIDAD)

If gexList.Value(gexList.Columns("Estado").Index) = "D" Then
    CmbEstado.ListIndex = 1
Else
    CmbEstado.ListIndex = 0
End If

Exit Sub
Cargar_DatosErr:
    ErrorHandler Err, "CARGA_GRID"
End Sub

Sub Habilita(vBoolean As Boolean)
    TxtCod_Usuario.Enabled = vBoolean
    TxtNom_Usuario.Enabled = vBoolean
    CmdUsuario.Enabled = vBoolean
    CmbEstado.Enabled = vBoolean
End Sub

Sub SALVAR_DATOS()
Dim Rs As ADODB.Recordset
On Error GoTo Cargar_DatosErr

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

Rs.Open "exec UP_MAN_USERGUIA 'U','" & gexList.Value(gexList.Columns("serie").Index) & "','" & gexList.Value(gexList.Columns("numero").Index) & "','" & TxtCod_Usuario.Text & "','" & Mid(CmbEstado.Text, 1, 1) & "'"

Exit Sub
Cargar_DatosErr:
    ErrorHandler Err, "CARGA_GRID"
    Set Rs = Nothing
End Sub

Private Sub txtSerie_GotFocus()
    optSerie.Value = True
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CARGA_GRID
    End If
    
End Sub
