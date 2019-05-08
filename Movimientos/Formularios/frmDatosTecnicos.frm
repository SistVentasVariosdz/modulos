VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmDatosTecnicos 
   Caption         =   "Datos Técnicos de Tela Acabada"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   Icon            =   "frmDatosTecnicos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   3360
      TabIndex        =   10
      Top             =   5520
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   900
      Custom          =   $"frmDatosTecnicos.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1430
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultado"
      Height          =   4215
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   10815
      Begin GridEX20.GridEX gexList 
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6800
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   14
         Column(1)       =   "frmDatosTecnicos.frx":0500
         Column(2)       =   "frmDatosTecnicos.frx":05EC
         Column(3)       =   "frmDatosTecnicos.frx":06BC
         Column(4)       =   "frmDatosTecnicos.frx":0790
         Column(5)       =   "frmDatosTecnicos.frx":085C
         Column(6)       =   "frmDatosTecnicos.frx":0930
         Column(7)       =   "frmDatosTecnicos.frx":0A04
         Column(8)       =   "frmDatosTecnicos.frx":0AD0
         Column(9)       =   "frmDatosTecnicos.frx":0BA4
         Column(10)      =   "frmDatosTecnicos.frx":0C70
         Column(11)      =   "frmDatosTecnicos.frx":0D3C
         Column(12)      =   "frmDatosTecnicos.frx":0E08
         Column(13)      =   "frmDatosTecnicos.frx":0EE0
         Column(14)      =   "frmDatosTecnicos.frx":0FB8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDatosTecnicos.frx":105C
         FormatStyle(2)  =   "frmDatosTecnicos.frx":1194
         FormatStyle(3)  =   "frmDatosTecnicos.frx":1244
         FormatStyle(4)  =   "frmDatosTecnicos.frx":12F8
         FormatStyle(5)  =   "frmDatosTecnicos.frx":13D0
         FormatStyle(6)  =   "frmDatosTecnicos.frx":1488
         ImageCount      =   0
         PrinterProperties=   "frmDatosTecnicos.frx":1568
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
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
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   10815
      Begin VB.OptionButton OptOT 
         Caption         =   "Partida Tintoreria"
         Height          =   375
         Left            =   7200
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtOT 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8280
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   9480
         TabIndex        =   8
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox TxtPartida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   62193665
         CurrentDate     =   37484
      End
      Begin VB.TextBox TxtGrupo 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OpPartida 
         Caption         =   "Partida"
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OpGrupo 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   480
      Top             =   5640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmDatosTecnicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOpcion As Integer
Public Codigo As String
Public Descripcion As String

Private Sub Form_Load()
    FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
    OpGrupo_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
If Valida_Busqueda Then CARGA_GRID
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim vRow As Variant
    
    
    
    Select Case ActionName
    Case "ACTDATOSPREN"
        If gexList.RowCount = 0 Then Exit Sub
        vRow = gexList.Row
        Load FrmDatosTecnicosPrendas
        With FrmDatosTecnicosPrendas
            .Tipo = "U"
            .vCod_TipOrdTra = gexList.Value(gexList.Columns("Cod_TipOrdTra").Index)
            .vCod_OrdTra = gexList.Value(gexList.Columns("O/T").Index)
            .vCod_Tela = gexList.Value(gexList.Columns("cod_tela").Index)
            .vCod_Comb = gexList.Value(gexList.Columns("cod_comb").Index)
            .vCod_Color = gexList.Value(gexList.Columns("cod_color").Index)
            .vPartida = gexList.Value(gexList.Columns("Partida").Index)
            .TxtPartida.Text = gexList.Value(gexList.Columns("Partida").Index)
            .TxtTela.Text = gexList.Value(gexList.Columns("cod_tela").Index) & " - " & gexList.Value(gexList.Columns("Tela").Index)
            .TxtColor.Text = gexList.Value(gexList.Columns("cod_color").Index) & " - " & gexList.Value(gexList.Columns("color").Index)
            .TxtComb.Text = gexList.Value(gexList.Columns("cod_comb").Index) & " - " & gexList.Value(gexList.Columns("comb").Index)
            .TxtEncogAncho.Text = CDbl(gexList.Value(gexList.Columns("encogimiento_ancho_prenda").Index))
            .TxtEncogLargo.Text = CDbl(gexList.Value(gexList.Columns("encogimiento_largo_prenda").Index))
            .TxtRevirado.Text = CDbl(gexList.Value(gexList.Columns("Revirado_prenda").Index))
            .TxtGramaje.Text = CDbl(gexList.Value(gexList.Columns("Gramaje_prenda").Index))
            .TxtObservaciones = Trim(gexList.Value(gexList.Columns("observaciones_prenda").Index))
            .FunctButt1.Visible = True
            .FunctButt2.Visible = False
            .Show vbModal
            If .vOk = True Then CARGA_GRID
        End With
        Set FrmDatosTecnicosPrendas = Nothing
    Case "CONSULTADATOSPREN"
        If gexList.RowCount = 0 Then Exit Sub
        vRow = gexList.Row
        Load FrmDatosTecnicosPrendas
            With FrmDatosTecnicosPrendas
                .Tipo = "C"
                .vCod_TipOrdTra = gexList.Value(gexList.Columns("Cod_TipOrdTra").Index)
                .vCod_OrdTra = gexList.Value(gexList.Columns("O/T").Index)
                .vCod_Tela = gexList.Value(gexList.Columns("cod_tela").Index)
                .vCod_Comb = gexList.Value(gexList.Columns("cod_comb").Index)
                .vCod_Color = gexList.Value(gexList.Columns("cod_color").Index)
                .vPartida = gexList.Value(gexList.Columns("Partida").Index)
                .TxtPartida.Text = gexList.Value(gexList.Columns("Partida").Index)
                .TxtTela.Text = gexList.Value(gexList.Columns("cod_tela").Index) & " - " & gexList.Value(gexList.Columns("Tela").Index)
                .TxtColor.Text = gexList.Value(gexList.Columns("cod_color").Index) & " - " & gexList.Value(gexList.Columns("color").Index)
                .TxtComb.Text = gexList.Value(gexList.Columns("cod_comb").Index) & " - " & gexList.Value(gexList.Columns("comb").Index)
                .TxtEncogAncho.Text = CDbl(gexList.Value(gexList.Columns("encogimiento_ancho_prenda").Index))
                .TxtEncogLargo.Text = CDbl(gexList.Value(gexList.Columns("encogimiento_largo_prenda").Index))
                .TxtRevirado.Text = CDbl(gexList.Value(gexList.Columns("Revirado_prenda").Index))
                .TxtGramaje.Text = CDbl(gexList.Value(gexList.Columns("Gramaje_prenda").Index))
                .TxtObservaciones = Trim(gexList.Value(gexList.Columns("observaciones_prenda").Index))
                .FunctButt2.Visible = True
                .FunctButt1.Visible = False
                .fraDatos.Enabled = False
                .Show vbModal
            End With
            Set FrmDatosTecnicosPrendas = Nothing
    Case "DATOSTEC"
        If gexList.RowCount = 0 Then Exit Sub
        vRow = gexList.Row
        With frmDetDatosTecnicos
            .vCod_TipOrdTra = gexList.Value(gexList.Columns("Cod_TipOrdTra").Index)
            .vCod_OrdTra = gexList.Value(gexList.Columns("O/T").Index)
            .vCod_Tela = gexList.Value(gexList.Columns("cod_tela").Index)
            .vCod_Comb = gexList.Value(gexList.Columns("cod_comb").Index)
            .vCod_Color = gexList.Value(gexList.Columns("cod_color").Index)
            .vPartida = gexList.Value(gexList.Columns("Partida").Index)
            .vDes_tela = gexList.Value(gexList.Columns("Tela").Index)
            .TxtCod_MotRechazo = RTrim(gexList.Value(gexList.Columns("cod_MotRechazo").Index))
            .TxtDes_MotRechazo = RTrim(gexList.Value(gexList.Columns("Des_MotRechazo").Index))
            .TxtElongacionLargo = gexList.Value(gexList.Columns("ElongacionLargo").Index)
            .TxtElongacionAncho = gexList.Value(gexList.Columns("ElongacionAncho").Index)
            If gexList.Value(gexList.Columns("FLG_TINTORERIA_PROPIA").Index) = "S" Then
                Aviso "Datos Técnicos de tintorería Propia sólo pueden ser Consultados", 2
                .FunctButt1.ChangeProperty "ENABLED", 0, False
            End If
            .Show 1
        End With
        CARGA_GRID
        Set frmDetDatosTecnicos = Nothing
        gexList.Row = vRow
    Case "DATOSDET"
        If gexList.RowCount = 0 Then Exit Sub
        vRow = gexList.Row
        Load frmDatosTecDetalle
        With frmDatosTecDetalle
            .vCod_TipOrdTra = gexList.Value(gexList.Columns("Cod_TipOrdTra").Index)
            .vCod_OrdTra = gexList.Value(gexList.Columns("O/T").Index)
            .vCod_Tela = gexList.Value(gexList.Columns("cod_tela").Index)
            .vCod_Comb = gexList.Value(gexList.Columns("cod_comb").Index)
            .vCod_Color = gexList.Value(gexList.Columns("cod_color").Index)
            .vPartida = gexList.Value(gexList.Columns("Partida").Index)
            .vDes_tela = gexList.Value(gexList.Columns("Tela").Index)
            .Show 1
        End With
        CARGA_GRID
        Set frmDatosTecDetalle = Nothing
        gexList.Row = vRow
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub OpFecha_Click()
    Limpia_Busquedas
    
    vOpcion = 2
    TxtGrupo.Enabled = False
    TxtPartida.Enabled = False
    TxtOT.Enabled = False
    DTPFecha.Enabled = True
    
    DTPFecha.SetFocus
End Sub

Private Sub OpGrupo_Click()
    Limpia_Busquedas
    
    vOpcion = 1
    DTPFecha.Enabled = False
    TxtPartida.Enabled = False
    TxtOT.Enabled = False
    
    TxtGrupo.Enabled = True
    
    'TxtGrupo.SetFocus
End Sub

Private Sub OpPartida_Click()
    Limpia_Busquedas
    
    vOpcion = 3
    TxtGrupo.Enabled = False
    DTPFecha.Enabled = False
    TxtOT.Enabled = False
    TxtPartida.Enabled = True
    
    TxtPartida.SetFocus
End Sub

Sub Limpia_Busquedas()
TxtGrupo.Text = ""
TxtPartida.Text = ""
DTPFecha.Value = Date
TxtOT.Text = ""
End Sub

Public Sub CARGA_GRID()
Dim Rs_Lista As ADODB.Recordset
Dim strSQL As String
'On Error GoTo hand
 Set Rs_Lista = New ADODB.Recordset

Rs_Lista.CursorLocation = adUseClient
Rs_Lista.ActiveConnection = cConnect

'If OpGrupo.Value Then
'    strSQL = "EXEC SM_TRAE_PARTIDAS_DATOS_TECNICOS '1','" & Trim(TxtGrupo.Text) & "',NULL,''"
'End If
'
'If OpFecha.Value Then
'    strSQL = "EXEC SM_TRAE_PARTIDAS_DATOS_TECNICOS '2','','" & DTPFecha.Value & "',''"
'End If
'
'If OpPartida.Value Then
'    strSQL = "EXEC SM_TRAE_PARTIDAS_DATOS_TECNICOS '3','','','" & Trim(TxtPartida.Text) & "'"
'End If
strSQL = "EXEC SM_TRAE_PARTIDAS_DATOS_TECNICOS '" & vOpcion & "','" & Trim(TxtGrupo.Text) & "','" & DTPFecha.Value & "','" & Trim(TxtPartida.Text) & "','" & Trim(TxtOT.Text) & "'"

Rs_Lista.Open strSQL
'If Rs_Lista.RecordCount Then
    Set gexList.ADORecordset = Rs_Lista
    Configurar_Grid
'End If

'hand:
'ErrorHandler Err, "Datos"
End Sub


Private Sub OptOT_Click()
Limpia_Busquedas
    
    vOpcion = 4
    TxtGrupo.Enabled = False
    DTPFecha.Enabled = False
    TxtPartida.Enabled = False
    
    TxtOT.Enabled = True
        
    TxtOT.SetFocus
End Sub

Private Sub TxtGrupo_Change()
    If Trim(Codigo) <> "" Or Not OpGrupo Then
        Exit Sub
    End If
    
    Load frmBuscaGrupo
    Set frmBuscaGrupo.oParent = Me
    frmBuscaGrupo.varTipo = "1"
    frmBuscaGrupo.txtCod_GrupoTex = Me.TxtGrupo.Text
    frmBuscaGrupo.CARGA_GRID
    frmBuscaGrupo.Show 1
    
    Set frmBuscaGrupo = Nothing
    
    If Trim(Codigo) <> "" Then
        Me.TxtGrupo.Text = Codigo
        FunctButt1.SetFocus
    End If
    Codigo = ""
    Descripcion = ""

End Sub

Sub Configurar_Grid()
    gexList.Columns("Cod_TipOrdTra").Visible = False
    'gexList.Columns("Cod_tela").Visible = False
    'gexList.Columns("Cod_Comb").Visible = False
    'gexList.Columns("Cod_Color").Visible = False
    gexList.Columns("O/T").Visible = False
    
    gexList.Columns("OP").Caption = DevuelveCampo("select tipo_orden from tg_control", cConnect)
    
    gexList.Columns("Cod_tela").Width = 850
    gexList.Columns("Cod_Comb").Width = 850
    gexList.Columns("Cod_Color").Width = 850
    
    gexList.Columns("Partida").Width = 800
    gexList.Columns("Tela").Width = 3500
    gexList.Columns("Comb").Width = 2000
    gexList.Columns("Gramaje").Width = 900
    gexList.Columns("Ancho").Width = 900
    gexList.Columns("Encog.Ancho").Width = 900
    gexList.Columns("Encog.Largo").Width = 900
    gexList.Columns("Revirado").Width = 900
    
    gexList.FrozenColumns = 5
End Sub

Function Valida_Busqueda() As Boolean
If vOpcion = 1 And Trim(TxtGrupo.Text) = "" Then
    MsgBox "Ingrese el grupo a buscar", vbInformation, "Busqueda"
    Valida_Busqueda = False
    Exit Function
End If

If vOpcion = 3 And Trim(TxtPartida.Text) = "" Then
    MsgBox "Ingrese la partida a buscar", vbInformation, "Busqueda"
    Valida_Busqueda = False
    Exit Function
End If

Valida_Busqueda = True
End Function

Private Sub TxtOT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtOT.Text = Right("00000" & Trim(TxtOT.Text), 5)
    FunctButt1.SetFocus
End If
End Sub

Private Sub TxtPartida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FunctButt1.SetFocus
End If
End Sub
