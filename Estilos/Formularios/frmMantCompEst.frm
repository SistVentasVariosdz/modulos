VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantCompEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Componente Estilo"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Style Component"
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   75
      TabIndex        =   16
      Tag             =   "Detail"
      Top             =   3900
      Width           =   8025
      Begin VB.CheckBox chkExigeComponenteAsociado 
         Alignment       =   1  'Right Justify
         Caption         =   "Exige Componente Asociado"
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   1780
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "Servicios de Manufactura"
         Height          =   855
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   7815
         Begin VB.CheckBox ckccom 
            Caption         =   "Componentes de Manufactura"
            Height          =   435
            Left            =   360
            TabIndex        =   32
            Top             =   285
            Width           =   1665
         End
         Begin VB.TextBox txtdes_familia 
            Height          =   285
            Left            =   4605
            TabIndex        =   31
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdBusFamItem 
            Caption         =   "..."
            Height          =   300
            Left            =   4290
            TabIndex        =   30
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtfamilia 
            Height          =   285
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   29
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Etiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Familia Asociada:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   2280
            TabIndex        =   33
            Tag             =   "Mat. Prima :"
            Top             =   285
            Width           =   1260
         End
      End
      Begin VB.CheckBox ChkAccesorio 
         Alignment       =   1  'Right Justify
         Caption         =   "Accesorio Tela"
         Height          =   195
         Left            =   285
         TabIndex        =   27
         Top             =   1785
         Width           =   1665
      End
      Begin VB.TextBox txtcod_ctacont 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         MaxLength       =   14
         TabIndex        =   9
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox TxtConsumo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   8
         Text            =   "0"
         Top             =   990
         Width           =   840
      End
      Begin VB.ComboBox TxtSector 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMantCompEst.frx":0000
         Left            =   6390
         List            =   "frmMantCompEst.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox TxtOrden 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "0"
         Top             =   285
         Width           =   840
      End
      Begin VB.TextBox txtPathIcono 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1365
         Width           =   3435
      End
      Begin VB.TextBox txtDesCompEst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   2
         Top             =   630
         Width           =   3435
      End
      Begin VB.TextBox txtDesTipComp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   19
         Top             =   990
         Width           =   2310
      End
      Begin VB.TextBox txtIdCompEst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         MaxLength       =   4
         TabIndex        =   1
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton cmdBuscaTipo 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2085
         TabIndex        =   4
         Tag             =   "..."
         Top             =   975
         Width           =   300
      End
      Begin VB.TextBox txtIdTipComp 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         MaxLength       =   1
         TabIndex        =   3
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cta Contable:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   4785
         TabIndex        =   25
         Tag             =   "Mat. Prima :"
         Top             =   1425
         Width           =   960
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Consumo por Prenda (Rectilineos p.e)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   4785
         TabIndex        =   24
         Tag             =   "Type:"
         Top             =   945
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4785
         TabIndex        =   23
         Tag             =   "Type:"
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4785
         TabIndex        =   22
         Tag             =   "Icon Path"
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Ruta Icono :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   21
         Tag             =   "Icon Path"
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   20
         Tag             =   "Description :"
         Top             =   675
         Width           =   1020
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   285
         TabIndex        =   18
         Tag             =   "Type:"
         Top             =   1035
         Width           =   390
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   285
         TabIndex        =   17
         Tag             =   "Code"
         Top             =   330
         Width           =   1020
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      TabIndex        =   15
      Tag             =   "List"
      Top             =   0
      Width           =   8025
      Begin GridEX20.GridEX DGridLista 
         Height          =   3525
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6218
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantCompEst.frx":0025
         Column(2)       =   "frmMantCompEst.frx":00ED
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantCompEst.frx":0191
         FormatStyle(2)  =   "frmMantCompEst.frx":02C9
         FormatStyle(3)  =   "frmMantCompEst.frx":0379
         FormatStyle(4)  =   "frmMantCompEst.frx":042D
         FormatStyle(5)  =   "frmMantCompEst.frx":0505
         FormatStyle(6)  =   "frmMantCompEst.frx":05BD
         ImageCount      =   0
         PrinterProperties=   "frmMantCompEst.frx":069D
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1245
      TabIndex        =   0
      Top             =   7065
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantCompEst.frx":0875
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantCompEst.frx":09E7
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantCompEst.frx":0B59
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantCompEst.frx":0CCB
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3360
      TabIndex        =   10
      Top             =   7140
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantCompEst.frx":0E3D
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantCompEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Carga As New ADODB.Recordset
Dim strSQL As String
Dim snum As Integer

Dim indicegrilla  As Long
Dim exigeComponenteAsociado As String


Private Sub chkExigeComponenteAsociado_Click()
    If Me.chkExigeComponenteAsociado = 0 Then
        exigeComponenteAsociado = "N"
    Else
        exigeComponenteAsociado = "S"
    End If
End Sub

Private Sub ckccom_Click()
    If ckccom.Value = 1 Then
        Etiqueta(5).Enabled = True
        txtfamilia.Enabled = True
        txtdes_familia.Enabled = True
        cmdBusFamItem.Enabled = True
    Else
        Etiqueta(5).Enabled = False
        txtfamilia.Text = ""
        txtdes_familia.Text = ""
        txtfamilia.Enabled = False
        txtdes_familia.Enabled = False
        cmdBusFamItem.Enabled = False
    End If
    
End Sub

Private Sub cmdBuscaTipo_Click()
Dim oTipo As New frmBusqGeneral
Dim rs As New ADODB.Recordset
Set oTipo.oParent = Me
oTipo.sQuery = "SELECT cod_tipcompest as Codigo, des_tipcompest as Descripcion FROM ES_TipCompEst"
oTipo.Cargar_Datos
oTipo.Show 1
If Codigo <> "" Then
    txtIdTipComp.Text = Codigo
    txtDesTipComp.Text = Descripcion
    Codigo = ""
    If txtIdTipComp = "T" Then TxtOrden.Enabled = True: TxtSector.Enabled = False
    If txtIdTipComp = "I" Then TxtOrden.Enabled = False: TxtSector.Enabled = True
End If
SendKeys "{TAB}"
Set oTipo = Nothing
End Sub
Private Sub cmdBuscaTipo_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub cmdFirst_Click()
If Not Rs_Carga.BOF Then
  Rs_Carga.MoveFirst
End If
End Sub
Private Sub cmdLast_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveLast
End If
End Sub
Private Sub cmdNext_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveNext
End If
End Sub
Private Sub cmdPrevious_Click()
If Not Rs_Carga.BOF Then
 Rs_Carga.MovePrevious
End If
End Sub

 
Sub Carga_Datos()
    Dim strSQL As String
    On Error GoTo Cargar_DatosErr
    'strSQL = "SG_Act_CompEst '','','','','L'"
        strSQL = "SG_Act_CompEst '','','','','L','','',0,'','','',''"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    ''Rs_Carga.CursorType = adOpenStatic
    ''Rs_Carga.CursorLocation = adUseClient
    ''Rs_Carga.LockType = adLockReadOnly
    ''Rs_Carga.Open strSQL
    Set Rs_Carga = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    'Set DGridLista.DataSource = Rs_Carga
    Set DGridLista.ADORecordset = Rs_Carga
    
    If DGridLista.RowCount > 0 Then DGridLista.Row = 1
    DGridLista_RowColChange 0, 0
    
    If DGridLista.RowCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
        DGridLista.Columns(1).Caption = "Codigo"
        DGridLista.Columns(2).Caption = "Descripcion"
        DGridLista.Columns(3).Caption = "Tipo"
        
        DGridLista.Columns(1).Width = 800
        DGridLista.Columns(2).Width = 2000
        DGridLista.Columns(3).Width = 1400
           
        DGridLista.RowSelected(indicegrilla) = True
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)

If DGridLista.Row <= 0 Then Exit Sub
Rs_Carga.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)

txtIdCompEst.Text = Rs_Carga!cod_compest
txtDesCompEst.Text = Rs_Carga!des_compest
txtIdTipComp.Text = Rs_Carga!cod_tipcompest

If IsNull(Rs_Carga!pat_icono) Then
    txtPathIcono.Text = ""
Else
    txtPathIcono.Text = Rs_Carga!pat_icono
End If
txtDesTipComp.Text = Rs_Carga!des_tipcompest
TxtOrden = Rs_Carga!orden
BuscaCombo Rs_Carga!SectorConfec, 1, TxtSector
TxtConsumo = Rs_Carga!consumo
txtcod_ctacont.Text = Trim(Rs_Carga!Cod_CtaCon)

If Trim(Rs_Carga!accesorio_tela) = "N" Then
    ChkAccesorio.Value = Unchecked
Else
    ChkAccesorio.Value = Checked
End If

If Trim(Rs_Carga!Servicio_Manufactura) = "N" Then
    ckccom.Value = Unchecked
Else
    ckccom.Value = Checked
End If

txtfamilia.Text = Rs_Carga!Cod_CtaCon

If IsNull(Rs_Carga!des_famitem) Then
    txtdes_familia.Text = ""
Else
    txtdes_familia.Text = Rs_Carga!des_famitem
End If

If Trim(Rs_Carga!Flg_Exige_Componente_Asociado) = "N" Then
    Me.chkExigeComponenteAsociado.Value = Unchecked
Else
    Me.chkExigeComponenteAsociado.Value = Checked
End If

DESHABILITA_DATOS
End Sub

Private Sub Form_Load()
Call FormSet(Me)
'FormateaGrid Me.DGridLista

txtfamilia.Text = ""
indicegrilla = 1

exigeComponenteAsociado = "N"
Carga_Datos
'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub
Sub SALVAR_DATOS()
Dim Con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdTipComp.Text <> "" Then
    Con.BeginTrans
    Con.Execute "SG_Act_CompEst '" & _
    txtIdCompEst.Text & "','" & _
    txtDesCompEst.Text & "','" & _
    txtIdTipComp.Text & "','" & _
    txtPathIcono.Text & "','" & _
    sTipo & "'," & _
    TxtOrden & ",'" & _
    Left(TxtSector, 1) & "'," & _
    Me.TxtConsumo & ", '" & _
    txtcod_ctacont & "','" & _
    IIf(ChkAccesorio.Value, "N", "S") & "','" & _
    IIf(ckccom.Value, "S", "N") & "','" & _
    txtfamilia.Text & "','" & _
    exigeComponenteAsociado & "'"
    
    
    
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    snum = 2
    ErrorHandler Err, "Salvar_Datos"
   
End Sub
Sub ELIMINAR_DATOS()
Dim Con As New ADODB.Connection
On Error GoTo Eliminar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdTipComp.Text <> "" Then
    Con.BeginTrans
    Con.Execute "SG_Act_CompEst '" & txtIdCompEst.Text & "','','','','D','','',0,'','','',''"
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
    txtIdCompEst.Text = ""
    txtDesCompEst.Text = ""
    txtIdTipComp.Text = ""
    txtDesTipComp.Text = ""
    txtPathIcono.Text = ""
    TxtOrden = "0"
    TxtSector.ListIndex = -1
    Me.TxtConsumo = "0.00"
    txtcod_ctacont.Text = ""
    ChkAccesorio.Value = Unchecked
    ckccom.Value = Unchecked
    txtfamilia.Text = ""
    txtdes_familia.Text = ""
End Sub
Private Sub DGridLista_Click()
If Rs_Carga.State <> 1 Then
    Exit Sub
End If

indicegrilla = DGridLista.Row
'If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
'    txtIdCompEst.Text = Rs_Carga!cod_compest
'    txtDesCompEst.Text = Rs_Carga!des_compest
'    txtIdTipComp.Text = Rs_Carga!cod_tipcompest
'    txtPathIcono.Text = Rs_Carga!pat_icono
'    txtDesTipComp.Text = Rs_Carga!des_tipcompest
'    TxtOrden = Rs_Carga!orden
'    BuscaCombo Rs_Carga!SectorConfec, 1, TxtSector
'    DESHABILITA_DATOS
'End If
End Sub
Sub HABILITA_DATOS()
    txtIdCompEst.Enabled = True
    txtDesCompEst.Enabled = True
    txtIdTipComp.Enabled = True
    txtPathIcono.Enabled = True
    
    TxtOrden.Enabled = True
    TxtSector.Enabled = True
    txtIdCompEst.SetFocus
    Me.TxtConsumo.Enabled = True
    txtcod_ctacont.Enabled = True
    ChkAccesorio.Enabled = True
    Me.chkExigeComponenteAsociado.Enabled = True
End Sub
Sub DESHABILITA_DATOS()
    txtIdCompEst.Enabled = False
    txtDesCompEst.Enabled = False
    txtIdTipComp.Enabled = False
    txtPathIcono.Enabled = False
    TxtOrden.Enabled = False
    TxtSector.Enabled = False
    Me.TxtConsumo.Enabled = False
    txtcod_ctacont.Enabled = False
    ChkAccesorio.Enabled = False
    Me.chkExigeComponenteAsociado.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
'Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If Rs_Carga.State <> 1 Then
'    Exit Sub
'End If
'If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
'    txtIdCompEst.Text = Rs_Carga!cod_compest
'    txtDesCompEst.Text = Rs_Carga!des_compest
'    txtIdTipComp.Text = Rs_Carga!cod_tipcompest
'    If IsNull(Rs_Carga!pat_icono) Then
'        txtPathIcono.Text = ""
'    Else
'        txtPathIcono.Text = Rs_Carga!pat_icono
'    End If
'    txtDesTipComp.Text = Rs_Carga!des_tipcompest
'    TxtOrden = Rs_Carga!orden
'    BuscaCombo Rs_Carga!SectorConfec, 1, TxtSector
'    TxtConsumo = Rs_Carga!consumo
'    txtcod_ctacont.Text = Trim(Rs_Carga!cod_ctacon)
'
'    DESHABILITA_DATOS
'End If
'End Sub
Sub RECARGAR_DATOS()
Rs_Carga.Close
Carga_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Rs_Carga = Nothing
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub
Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        sTipo = "I"
        LIMPIAR_DATOS
        HABILITA_DATOS
        cmdBuscaTipo.Enabled = True
        txtIdCompEst.SetFocus
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "MODIFICAR"
        sTipo = "U"
        txtDesCompEst.Enabled = True
        txtIdTipComp.Enabled = True
        txtPathIcono.Enabled = True
        cmdBuscaTipo.Enabled = True
        Me.TxtConsumo.Enabled = True
        txtcod_ctacont.Enabled = True
        ChkAccesorio.Enabled = True
        Me.chkExigeComponenteAsociado.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
        If txtIdTipComp = "T" Then TxtOrden.Enabled = True: TxtSector.Enabled = False
        If txtIdTipComp = "I" Then TxtOrden.Enabled = False: TxtSector.Enabled = True

    Case "ELIMINAR"
        ELIMINAR_DATOS
    Case "GRABAR"
    snum = 0
        If VALIDA_DATOS Then
            SALVAR_DATOS
            If snum = 0 Then
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
            DGridLista.Enabled = True
            cmdBuscaTipo.Enabled = False
            End If
        End If
    Case "DESHACER"
        LIMPIAR_DATOS
        'RECARGAR_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
        DGridLista.Enabled = True
        cmdBuscaTipo.Enabled = False
    Case "SALIR"
        Unload Me
End Select
End Sub
Function VALIDA_DATOS() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
VALIDA_DATOS = True
If Len(Trim(txtDesCompEst.Text)) = 0 Then
   MsgBox "Ingrese la descripcion", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Len(Trim(txtIdCompEst.Text)) = 0 Then
   MsgBox "Ingrese el Tipo", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If

If Len(Trim(txtIdTipComp.Text)) = 0 Then
    Call MsgBox("El tipo de componente no puede estar vacio. Sirvase verificar", vbInformation, "Tipo de Componente")
    txtIdTipComp.SetFocus
    VALIDA_DATOS = False
End If

If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
End If


End Function

Private Sub txtcod_ctacont_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub TxtConsumo_KeyPress(KeyAscii As Integer)
    SoloNumeros ActiveControl, KeyAscii, True, 3, 6
    AVANZA (KeyAscii)
End Sub

Private Sub txtDesCompEst_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub

Private Sub txtfamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'txtgrupo.Text = ""
        If Trim(txtfamilia.Text) = "" Then
            cmdBusFamItem_Click
        Else
            If ValidaFamilia = False Then
                 Exit Sub
            Else
                strSQL = "SELECT DES_FAMITEM FROM LG_FAMITE WHERE COD_FAMITEM='" & txtfamilia.Text & "'"
                txtdes_familia.Text = DevuelveCampo(strSQL, cCONNECT)
                
'                txtgrupo.Enabled = True
'                cmdBusgrupo.Enabled = True
                MantFunc1.SetFocus
                
            End If
        End If
    End If
End Sub

Public Function ValidaFamilia() As Boolean
    Dim rs As New ADODB.Recordset
    Dim opcmessage As Integer
    rs.ActiveConnection = cCONNECT
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockReadOnly
    rs.Open "SELECT COD_FAMITEM as Código, DES_FAMITEM as Descripción FROM LG_FAMITE WHERE COD_FAMITEM='" & Trim(txtfamilia.Text) & "'"
    If rs.EOF Then
        opcmessage = MsgBox("La familia ingresada no existe, Desea Crearla?", vbInformation + vbYesNo)
        If opcmessage = vbYes Then
            Load frmMantFamTela
            frmMantFamTela.Show 1
            
        Else
        ValidaFamilia = False
        End If
    Else
        ValidaFamilia = True
    End If
    Set rs = Nothing
End Function
Private Sub cmdBusFamItem_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT COD_FAMITEM as Código, DES_FAMITEM as Descripción FROM LG_FAMITE"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtfamilia.Text = Codigo
        txtdes_familia.Text = Descripcion
        
        'txtgrupo.Enabled = True
        'cmdBusgrupo.Enabled = True
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub txtIdCompEst_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdCompEst_LostFocus()
Busca_CompEst
End Sub
Private Sub txtIdTipComp_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdTipComp_LostFocus()
If Len(Trim(txtIdTipComp)) <> 0 Then
    Busca_TipComp
End If
End Sub
Sub Busca_TipComp()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM ES_TipCompEst " & _
"WHERE cod_tipcompest = '" & txtIdTipComp.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesTipComp.Text = Rs_busca!des_tipcompest
Else
    txtDesTipComp.Text = ""
    txtIdTipComp.Text = ""
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub
Sub Busca_CompEst()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM ES_CompEst " & _
"WHERE cod_compest = '" & txtIdCompEst.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesCompEst.Text = Rs_busca!des_compest
    txtIdTipComp.Text = Rs_busca!cod_tipcompest
    txtPathIcono.Text = Rs_busca!pat_icono
    Busca_TipComp
    DESHABILITA_DATOS
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    DGridLista.Enabled = True
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub



Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
    SoloNumeros ActiveControl, KeyAscii, False, 0, 4
    AVANZA (KeyAscii)
End Sub


Private Sub TxtOrden_LostFocus()
    If Trim(TxtOrden.Text) = "" Then
        TxtOrden.Text = "0"
    End If
End Sub

Private Sub txtPathIcono_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub TxtSector_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub


