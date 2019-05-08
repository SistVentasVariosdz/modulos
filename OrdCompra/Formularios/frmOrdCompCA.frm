VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmOrdComp 
   Caption         =   "Ordenes de Compra"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9495
   Icon            =   "frmOrdComp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   6705
      TabIndex        =   63
      Top             =   7500
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~CERRAR~True~True~&Cerrar~0~0~1~~0~False~False~&Cerrar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2610
      TabIndex        =   23
      Top             =   7500
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmOrdComp.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame fraOpciones 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   8040
      TabIndex        =   22
      Top             =   3210
      Width           =   1455
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   2910
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   5133
         Custom          =   $"frmOrdComp.frx":046A
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   30
      TabIndex        =   21
      Top             =   3240
      Width           =   7905
      Begin VB.TextBox TxtNumImportacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   84
         Top             =   3195
         Width           =   840
      End
      Begin VB.TextBox TxtCodProv 
         Height          =   315
         Left            =   4650
         TabIndex        =   82
         Top             =   2835
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2835
         Width           =   1590
      End
      Begin VB.ComboBox cboCod_CenCost 
         Height          =   315
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2490
         Width           =   1695
      End
      Begin VB.TextBox txtPorc_IGV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6645
         TabIndex        =   40
         Top             =   690
         Width           =   915
      End
      Begin VB.ComboBox cboCod_ProTex 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2490
         Width           =   1590
      End
      Begin VB.TextBox TxtDes_Grupo 
         Height          =   315
         Left            =   3285
         MaxLength       =   50
         TabIndex        =   56
         Top             =   2160
         Width           =   3090
      End
      Begin VB.CommandButton cmdBuscaGrupo 
         Caption         =   "..."
         Height          =   330
         Left            =   2985
         TabIndex        =   55
         Top             =   2115
         Width           =   330
      End
      Begin VB.TextBox txtCod_Grupo 
         Height          =   315
         Left            =   1710
         MaxLength       =   8
         TabIndex        =   54
         Top             =   2130
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Fin 
         Height          =   315
         Left            =   4650
         TabIndex        =   52
         Top             =   1770
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   23592963
         CurrentDate     =   37267
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Inicio 
         Height          =   315
         Left            =   1710
         TabIndex        =   50
         Top             =   1770
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   23592963
         CurrentDate     =   37267
      End
      Begin VB.ComboBox cboCod_ClaOrdComp 
         Height          =   315
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1410
         Width           =   2925
      End
      Begin VB.ComboBox cboCod_StaOrdComp 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1410
         Width           =   1560
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   465
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   3615
         Width           =   6105
      End
      Begin VB.ComboBox cboCod_LugEntr 
         Height          =   315
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1050
         Width           =   2925
      End
      Begin VB.ComboBox cboCod_Moneda 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1050
         Width           =   1560
      End
      Begin VB.ComboBox cboCod_Descuento 
         Height          =   315
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   690
         Width           =   1260
      End
      Begin VB.ComboBox cboCod_CondVent 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   690
         Width           =   1560
      End
      Begin VB.TextBox txtDes_Proveedor 
         Height          =   315
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   34
         Top             =   330
         Width           =   1725
      End
      Begin VB.TextBox txtCod_Proveedor 
         Height          =   315
         Left            =   4650
         MaxLength       =   12
         TabIndex        =   33
         Top             =   330
         Width           =   1200
      End
      Begin VB.TextBox txtCod_OrdComp 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   31
         Top             =   330
         Width           =   1560
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Num. Importación:"
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   3270
         Width           =   1290
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Prov :"
         Height          =   195
         Left            =   3510
         TabIndex        =   83
         Top             =   2910
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Total/Parcial:"
         Height          =   195
         Left            =   105
         TabIndex        =   80
         Top             =   2940
         Width           =   1320
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "C. de Costo :"
         Height          =   195
         Left            =   3510
         TabIndex        =   59
         Top             =   2565
         Width           =   915
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Textíl :"
         Height          =   195
         Left            =   105
         TabIndex        =   57
         Top             =   2565
         Width           =   1125
      End
      Begin VB.Label Label17 
         Caption         =   "Grupo :"
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Fin :"
         Height          =   195
         Left            =   3510
         TabIndex        =   51
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Inicio :"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1845
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Clase OC :"
         Height          =   195
         Left            =   3510
         TabIndex        =   47
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado de la OC :"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1485
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones :"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   3735
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Lugar Entrega :"
         Height          =   195
         Left            =   3510
         TabIndex        =   43
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.:"
         Height          =   195
         Left            =   6045
         TabIndex        =   39
         Top             =   765
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dsctos :"
         Height          =   195
         Left            =   3510
         TabIndex        =   37
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cond. Venta :"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   3510
         TabIndex        =   32
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   15
      TabIndex        =   1
      Top             =   45
      Width           =   9450
      Begin VB.Frame fraoptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   160
         Width           =   7335
         Begin VB.OptionButton optOP 
            Caption         =   "Op"
            Height          =   195
            Left            =   6660
            TabIndex        =   72
            Top             =   105
            Width           =   855
         End
         Begin VB.OptionButton OpGrupo 
            Caption         =   "Grupo"
            Height          =   195
            Left            =   4005
            TabIndex        =   65
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   2205
            TabIndex        =   20
            Top             =   120
            Width           =   1425
         End
         Begin VB.OptionButton optEstado 
            Caption         =   "Estado"
            Height          =   195
            Left            =   5445
            TabIndex        =   19
            Top             =   120
            Width           =   1185
         End
         Begin VB.OptionButton optOrdCompra 
            Caption         =   "Orden de Compra"
            Height          =   195
            Left            =   45
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   8040
         TabIndex        =   16
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Frame FraOrdComp 
         Height          =   645
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtCodOrdComp 
            Height          =   285
            Left            =   4275
            MaxLength       =   6
            TabIndex        =   7
            Top             =   270
            Width           =   1425
         End
         Begin VB.TextBox txtSerOrdComp 
            Height          =   285
            Left            =   1500
            MaxLength       =   3
            TabIndex        =   4
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Numero"
            Height          =   195
            Left            =   3075
            TabIndex        =   6
            Top             =   345
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            Height          =   195
            Left            =   300
            TabIndex        =   3
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame FraEstado 
         Height          =   640
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtCodStaOrdComp 
            Height          =   285
            Left            =   1500
            MaxLength       =   1
            TabIndex        =   13
            Top             =   270
            Width           =   1005
         End
         Begin VB.TextBox txtDesStaOrdComp 
            Height          =   285
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   15
            Top             =   255
            Width           =   4200
         End
         Begin VB.CommandButton cmdBusEstado 
            Caption         =   "..."
            Height          =   330
            Left            =   2520
            TabIndex        =   14
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label2 
            Caption         =   "Estado :"
            Height          =   240
            Left            =   300
            TabIndex        =   12
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame FraProveedor 
         Height          =   640
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtDesProveedor 
            Height          =   285
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   11
            Top             =   270
            Width           =   4155
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   285
            Left            =   1500
            MaxLength       =   12
            TabIndex        =   9
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor :"
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame FraGrupo 
         Height          =   645
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   7455
         Begin VB.OptionButton OpLog 
            Caption         =   "Logistico"
            Height          =   255
            Left            =   6360
            TabIndex        =   71
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OpTex 
            Caption         =   "Textil"
            Height          =   255
            Left            =   6360
            TabIndex        =   70
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox TxtCodGrupo 
            Height          =   315
            Left            =   1320
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtDesGrupo 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   2520
            TabIndex        =   67
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label20 
            Caption         =   "Grupo :"
            Height          =   255
            Left            =   480
            TabIndex        =   69
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraOP 
         Height          =   645
         Left            =   240
         TabIndex        =   73
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtDes_OrdPro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4350
            TabIndex        =   77
            Top             =   210
            Width           =   2925
         End
         Begin VB.TextBox txtCod_OrdPro 
            Height          =   285
            Left            =   3645
            TabIndex        =   76
            Top             =   210
            Width           =   645
         End
         Begin VB.TextBox txtNom_Fabrica 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1245
            TabIndex        =   75
            Top             =   210
            Width           =   1575
         End
         Begin VB.TextBox txtCod_Fabrica 
            Height          =   285
            Left            =   705
            TabIndex        =   74
            Top             =   210
            Width           =   480
         End
         Begin VB.Label lblFabrica 
            Caption         =   "Fábrica"
            Height          =   195
            Left            =   90
            TabIndex        =   79
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lblorden 
            Caption         =   "Orden"
            Height          =   195
            Left            =   2985
            TabIndex        =   78
            Top             =   270
            Width           =   585
         End
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   30
      TabIndex        =   0
      Top             =   1320
      Width           =   9435
      Begin GridEX20.GridEX gexLista 
         Height          =   1485
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   2619
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         EmptyRows       =   -1  'True
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         GridLines       =   1
         ColumnHeaderHeight=   285
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "frmOrdComp.frx":05F9
         Column(2)       =   "frmOrdComp.frx":074D
         Column(3)       =   "frmOrdComp.frx":0881
         Column(4)       =   "frmOrdComp.frx":09C9
         Column(5)       =   "frmOrdComp.frx":0A6D
         Column(6)       =   "frmOrdComp.frx":0B11
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmOrdComp.frx":0BB5
         FormatStyle(2)  =   "frmOrdComp.frx":0CED
         FormatStyle(3)  =   "frmOrdComp.frx":0D9D
         FormatStyle(4)  =   "frmOrdComp.frx":0E51
         FormatStyle(5)  =   "frmOrdComp.frx":0F29
         FormatStyle(6)  =   "frmOrdComp.frx":0FE1
         ImageCount      =   0
         PrinterProperties=   "frmOrdComp.frx":10C1
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   195
      TabIndex        =   24
      Top             =   7395
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1455
         Picture         =   "frmOrdComp.frx":1299
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ultimo"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmOrdComp.frx":140B
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Siguiente"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   495
         Picture         =   "frmOrdComp.frx":157D
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Anterior"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmOrdComp.frx":16EF
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Primero"
         Top             =   105
         Width           =   495
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   8400
      Top             =   7140
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmOrdComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
Dim sTipo As String
Dim Opcion As Integer
Public codigo, Descripcion As String
'VAriables del Form
Public varCod_TipRequ As Integer
Dim varSer_OrdComp As String
Dim varProvCod_ClaOrdComp As String
Dim varFlg_Requerimiento As Boolean
'Variables para la impresion
Public varCadena_Familias As String
Public varCancelImpresion As Integer
Dim sTituliAbrOP As String
Public varAyuda As Integer

Dim Pregunta As Variant

Private Sub cboCod_ClaOrdComp_Click()
   Dim varCod_Protex As String
   Dim varTip_Item As String
   Dim varTip_Presentacion As String
    'si no tiene proceso relacionado entonces es un proceso post tenido
    strSQL = "SELECT ISNULL(Cod_Protex,'') FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
    varCod_Protex = DevuelveCampo(strSQL, cConnect)
    If Trim(varCod_Protex) = "" Then
        If sTipo = "I" Or sTipo = "U" Then
            cboCod_ProTex.Enabled = True
        End If
        
        strSQL = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varTip_Item = DevuelveCampo(strSQL, cConnect)
        
        strSQL = "SELECT Tip_Presentacion FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varTip_Presentacion = DevuelveCampo(strSQL, cConnect)
        
        If varTip_Item = "T" And varTip_Presentacion = "T" Then
            strSQL = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Flg_TejTen = 'T' AND Flg_principal = ''"
            Call LlenaCombo(cboCod_ProTex, strSQL, cConnect)
        Else
            If varTip_Item = "T" And varTip_Presentacion = "C" Then
                strSQL = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Flg_TejTen = 'J' AND Flg_principal = ''"
                Call LlenaCombo(cboCod_ProTex, strSQL, cConnect)
            Else
                cboCod_ProTex.Clear
            End If
        End If
    Else
        'Aqui llenamos los codigos de los procesos textiles
        strSQL = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Cod_ProTex = '" & varCod_Protex & "'"
        Call LlenaCombo(cboCod_ProTex, strSQL, cConnect)
        cboCod_ProTex.Enabled = False
        cboCod_ProTex.ListIndex = 0
    End If
   
    
    strSQL = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
    If DevuelveCampo(strSQL, cConnect) = "S" Then
        strSQL = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varCod_TipRequ = DevuelveCampo(strSQL, cConnect)
    End If
    
'    Strsql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
'    If DevuelveCampo(Strsql, cCONNECT) = "S" Then
'        txtCod_Grupo.Enabled = True
'        TxtDes_Grupo.Enabled = True
'        cmdBuscaGrupo.Enabled = True
'
'        varFlg_Requerimiento = True
'        'ProvCod_ClaOrdComp = Right(cboCod_ClaOrdComp.Text, 2)
'
'        Strsql = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
'        varCod_TipRequ = DevuelveCampo(Strsql, cCONNECT)
'
'    Else
'        txtCod_Grupo.Text = ""
'        TxtDes_Grupo.Text = ""
'        txtCod_Grupo.Enabled = False
'        TxtDes_Grupo.Enabled = False
'        cmdBuscaGrupo.Enabled = False
'
'        varFlg_Requerimiento = False
'    End If
'
'    If sTipo = "" Then
'        txtCod_Grupo.Enabled = False
'        TxtDes_Grupo.Enabled = False
'        cmdBuscaGrupo.Enabled = False
'    End If
    
End Sub



Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Sub LIMPIAR_DATOS()
    
    txtCod_OrdComp.Text = ""
    txtCod_Proveedor.Text = ""
    txtDes_Proveedor.Text = ""
    TxtNumImportacion.Text = ""
    
    cboCod_CondVent.ListIndex = -1
    cboCod_Descuento.ListIndex = -1
    
    cboCod_Moneda.ListIndex = -1
    cboCod_LugEntr.ListIndex = -1
    txtObservaciones.Text = ""
    cboCod_StaOrdComp.ListIndex = -1
    cboCod_ClaOrdComp.ListIndex = -1
    dtpFec_Entrega_Inicio.Value = Date
    dtpFec_Entrega_Fin.Value = Date
    cboCod_CenCost.ListIndex = -1
    txtCod_Grupo.Text = ""
    TxtDes_Grupo.Text = ""
    cboCod_ProTex.ListIndex = -1
    CboEstado.ListIndex = -1

    'Aqui llenamos a los valores por defecto
    strSQL = "SELECT Porc_IGV FROM TG_IGV WHERE ANO=YEAR(GETDATE()) AND MES=RIGHT('0'+CONVERT(VARCHAR,MONTH(GETDATE())),2) "
    txtPorc_IGV.Text = DevuelveCampo(strSQL, cConnect)

End Sub

Sub CARGA_COMBOS()

    'Aqui llenamos las condiciones de Venta
    strSQL = "SELECT Des_CondVent + SPACE(100)+ Cod_CondVent FROM LG_CONDVENT"
    Call LlenaCombo(cboCod_CondVent, strSQL, cConnect)
    
    'Aqui llenamos los Descuentos
    strSQL = "SELECT CONVERT(VARCHAR,Porcentaje1) + ' - '+ CONVERT(VARCHAR,Porcentaje2) + SPACE(100) + COD_DESCUENTO FROM LG_DSCTOS"
    Call LlenaCombo(cboCod_Descuento, strSQL, cConnect)
    
    'Aqui llenamos las Monedas
    strSQL = "SELECT Nom_Moneda + SPACE(100) + Cod_Moneda FROM TG_MONEDA"
    Call LlenaCombo(cboCod_Moneda, strSQL, cConnect)
    
    
    strSQL = "SELECT Des_LugEntr + SPACE(100) + Cod_LugEntr FROM LG_LUGENTR"
    Call LlenaCombo(cboCod_LugEntr, strSQL, cConnect)
    
    strSQL = "SELECT Des_StaOrdComp + SPACE(100) + Cod_StaOrdComp FROM LG_STAORDCOMP"
    Call LlenaCombo(cboCod_StaOrdComp, strSQL, cConnect)
    
    strSQL = "SELECT a.Des_ClaOrdComp + SPACE(100) + a.Cod_ClaOrdComp FROM LG_CLAORDCOMP a,lg_segordcomp b where a.cod_claordcomp = b.cod_claordcomp and b.cod_usuario ='" & vusu & "'"
    Call LlenaCombo(cboCod_ClaOrdComp, strSQL, cConnect)
    
    'Aqui llenamos los codigos de los procesos textiles
    strSQL = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS"
    Call LlenaCombo(cboCod_ProTex, strSQL, cConnect)
    
    'Aqui llenamos nos centros de costo
    strSQL = "SELECT Des_CenCost + SPACE(100) + Cod_CenCost FROM TG_CENCOSTO"
    Call LlenaCombo(cboCod_CenCost, strSQL, cConnect)
    
    'Aqui llenamos los estados (Total y parcial)
    
    strSQL = "SELECT DESCRIPCION + SPACE(100) + FLG_TOTAL_PARCIAL FROM LG_MODORDCOMP"
    Call LlenaCombo(CboEstado, strSQL, cConnect)
End Sub

Function VALIDA_DATOS() As Boolean
    Dim NombreTabla As String
    Dim CodigoTabla As String
    

    VALIDA_DATOS = True
    If sTipo <> "D" Then
'
'        If sTipo = "I" Then
'            If ExisteCampo("Cod_StaOrdComp", "Lg_StaOrdComp", Trim(txtcod_StaOrdComp.Text), cCONNECT, True) Then
'                MsgBox "El código de Status de Orden de Compra ya se encuentra registrado. Sirvase verificar", vbInformation, "Status de Orden de Compra"
'                txtcod_StaOrdComp.SetFocus
'                VALIDA_DATOS = False
'                Exit Function
'            End If
'        End If
'
'        If Trim(txtcod_StaOrdComp.Text) = "" Then
'            MsgBox "El código de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtcod_StaOrdComp.Text = ""
'            txtcod_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If
'
'        If Trim(txtDes_StaOrdComp.Text) = "" Then
'            MsgBox "La descripción de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtDes_StaOrdComp.Text = ""
'            txtDes_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If

        If Trim(txtCod_Proveedor.Text) = "" Then
            MsgBox "El Código de Proveedor no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCod_Proveedor.Text = ""
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        strSQL = "SELECT count(*) FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
        If DevuelveCampo(strSQL, cConnect) = "0" Then
            MsgBox "El código de proveedor ingresado no es válido. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
        If Trim(cboCod_Descuento.Text) = "" Then
            MsgBox "El descuento no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_Descuento.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
    
        If Trim(cboCod_CondVent.Text) = "" Then
            MsgBox "La condición de venta no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_CondVent.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_Moneda.Text) = "" Then
            MsgBox "La moneda no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_Moneda.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_LugEntr.Text) = "" Then
            MsgBox "El campo lugar de entrega no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_LugEntr.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_ClaOrdComp.Text) = "" Then
            MsgBox "La clase de orden de compra no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_ClaOrdComp.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If dtpFec_Entrega_Fin.Value < dtpFec_Entrega_Inicio.Value Then
            MsgBox "La fecha de entrega final no puede ser menor a la inicial. Sirvase verificar", vbInformation, "Ordenes de Compra"
            dtpFec_Entrega_Fin.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        'Preguntamos por la variable si es requerida o no
        strSQL = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        
        If DevuelveCampo(strSQL, cConnect) <> "S" Then
            If Trim(cboCod_CenCost.Text) = "" Then
                MsgBox "El centro de costo no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                cboCod_CenCost.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
        
        If varFlg_Requerimiento = True Then
        
            If Trim(txtCod_Grupo.Text) = "" Then
                MsgBox "El grupo no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                txtCod_Grupo.Text = ""
                txtCod_Grupo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
            
            'Como el grupo puede ser textil o log, determinamos primero de quien se trata
            strSQL = "SELECT Tip_Grupo FROM LG_TIPREQ WHERE Cod_TipRequ='" & varCod_TipRequ & "'"
            If DevuelveCampo(strSQL, cConnect) = "I" Then
                NombreTabla = "ES_GRUPOLOG"
                CodigoTabla = "Cod_GrupoLog"
            Else
                NombreTabla = "ES_GRUPOTEX"
                CodigoTabla = "Cod_GrupoTex"
            End If
            'Una vez determ el grupo preguntamos si el codigo existe
            strSQL = "SELECT count(*) FROM " & NombreTabla & " WHERE " & CodigoTabla & " = '" & txtCod_Grupo.Text & "'"
            If DevuelveCampo(strSQL, cConnect) = "0" Then
                MsgBox "El codigo de grupo ingresado no es válido. Sirvase verificar", vbInformation, "Ordenes de Compra"
                txtCod_Grupo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If

        strSQL = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        If DevuelveCampo(strSQL, cConnect) <> "I" Then
            If Trim(cboCod_ProTex.Text) = "" Then
                MsgBox "El proceso textil no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                cboCod_ProTex.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If

    Else
        'Aqui se valida que no tenga registros dependientes
        strSQL = "SELECT COUNT(*) FROM LG_ORDCOMPITEM WHERE Ser_OrdComp='" & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "' AND Cod_OrdComp='" & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "'"
        If DevuelveCampo(strSQL, cConnect) > 0 Then
            MsgBox "El registro seleccionado posee registros relacionados. Sirvase verificar", vbInformation, "Ordenes de Compra"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub CARGA_DATOS()

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        
        varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
        txtCod_OrdComp.Text = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
        txtPorc_IGV.Text = gexLista.Value(gexLista.Columns("I.G.V.").Index)
        txtObservaciones.Text = gexLista.Value(gexLista.Columns("Observaciones").Index)
        TxtCodProv.Text = gexLista.Value(gexLista.Columns("Observaciones").Index)
        dtpFec_Entrega_Inicio.Value = gexLista.Value(gexLista.Columns("F.Entrega Inicial").Index)
        TxtNumImportacion.Text = gexLista.Value(gexLista.Columns("Num_Importacion").Index)
        'dtpFec_Entrega_Fin.Value = gexLista.Value(gexLista.Columns("F.Entrega Final").Index)
        
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_CondVent").Index), 2, cboCod_CondVent)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_Descuento").Index), 2, cboCod_Descuento)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_CenCost").Index), 2, cboCod_CenCost)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_Moneda").Index), 2, cboCod_Moneda)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_LugEntr").Index), 2, cboCod_LugEntr)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index), 2, cboCod_StaOrdComp)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index), 2, cboCod_ClaOrdComp)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_ProTex").Index), 2, cboCod_ProTex)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Flg_Total_Parcial").Index), 2, CboEstado)
        
        txtCod_Proveedor.Text = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
        Call BUSCA_PROVEEDOR(1, 2)
        txtCod_Grupo.Text = gexLista.Value(gexLista.Columns("Cod.Grupo").Index)
        Call BUSCA_GRUPO(1)
        
    End If
End Sub

Sub HABILITA_DATOS()
Dim RsDet As ADODB.Recordset
    If sTipo = "I" Then
        cboCod_ClaOrdComp.Enabled = True
        txtCod_Grupo.Enabled = True
        TxtDes_Grupo.Enabled = True
        cmdBuscaGrupo.Enabled = True
   Else
        Set RsDet = Nothing
        Set RsDet = New ADODB.Recordset
        RsDet.CursorLocation = adUseClient
        RsDet.Open "SELECT * FROM lg_ordcompitem WHERE Ser_OrdComp='" & Trim(gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)) & "' AND Cod_OrdComp='" & Trim(gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)) & "'", cConnect
        
        If RsDet.RecordCount = 0 Then
            txtCod_Grupo.Enabled = True
            TxtDes_Grupo.Enabled = True
            cmdBuscaGrupo.Enabled = True
'        Else
'            txtCod_Grupo.Enabled = False
'            txtDes_Grupo.Enabled = False
'            cmdBuscaGrupo.Enabled = False
        End If
    End If
    
    txtCod_Proveedor.Enabled = True
    txtDes_Proveedor.Enabled = True
    cboCod_CondVent.Enabled = True
    cboCod_Descuento.Enabled = True
    cboCod_Moneda.Enabled = True
    cboCod_LugEntr.Enabled = True
    txtObservaciones.Enabled = True
        
    cboCod_CenCost.Enabled = True
    cboCod_ProTex.Enabled = True
    
    dtpFec_Entrega_Fin.Enabled = True
    dtpFec_Entrega_Inicio.Enabled = True
End Sub

Sub INHABILITA_DATOS()
    
    txtCod_Proveedor.Enabled = False
    txtDes_Proveedor.Enabled = False
    cboCod_CondVent.Enabled = False
    cboCod_Descuento.Enabled = False
    cboCod_Moneda.Enabled = False
    cboCod_LugEntr.Enabled = False
    txtObservaciones.Enabled = False
    cboCod_StaOrdComp.Enabled = False
    cboCod_ClaOrdComp.Enabled = False
    cboCod_CenCost.Enabled = False
    txtCod_Grupo.Enabled = False
    TxtDes_Grupo.Enabled = False
    cmdBuscaGrupo.Enabled = False
    cboCod_ProTex.Enabled = False
    CboEstado.Enabled = False
    dtpFec_Entrega_Fin.Enabled = False
    dtpFec_Entrega_Inicio.Enabled = False

End Sub

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC UP_SEL_ORDCOMP " & CStr(Opcion) & ",'" & Trim(txtSerOrdComp.Text) & "','" & Trim(txtCodOrdComp.Text) & "','" & Trim(txtCodProveedor.Text) & "','" & Trim(txtCodStaOrdComp.Text) & "','" & Trim(TxtCodGrupo.Text) & "','" & vusu & "','" & Trim(txtCod_Fabrica.Text) & "','" & Trim(txtCod_OrdPro.Text) & "'"
    
    Rs_Lista.Open strSQL
    Set gexLista.ADORecordset = Rs_Lista

    If Rs_Lista.RecordCount > 0 Then
        gexLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    CONFIGURAR_GRID
End Sub

Private Sub CONFIGURAR_GRID()
    gexLista.Columns("Ser_OrdComp").Visible = False
    gexLista.Columns("Cod_OrdComp").Visible = False
    gexLista.Columns("Cod_Proveedor").Visible = False
    gexLista.Columns("Cod_CondVent").Visible = False
    gexLista.Columns("Cod_LugEntr").Visible = False
    gexLista.Columns("Cod_StaOrdComp").Visible = False
    gexLista.Columns("Cod_ClaOrdComp").Visible = False
    gexLista.Columns("Cod_ProTex").Visible = False
    gexLista.Columns("Cod_CenCost").Visible = False
    gexLista.Columns("Cod_Moneda").Visible = False
    gexLista.Columns("Cod_Descuento").Visible = False
    gexLista.Columns("Observaciones").Visible = False
    gexLista.Columns("Flg_Total_Parcial").Visible = False
    
    gexLista.Columns("Proveedor").Width = 2500
    gexLista.Columns("I.G.V.").Width = 700
    gexLista.Columns("O.C.").Width = 1100
    gexLista.Columns("Descuentos").Width = 900
    gexLista.Columns("Cod.Grupo").Width = 900
    gexLista.Columns("Moneda").Width = 2000
    gexLista.Columns("L.Entrega").Width = 2000
    gexLista.Columns("CondVenta").Width = 2000
End Sub

Sub CAMBIO_ESTADO()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        strSQL = "EXEC UP_MAN_ORDCOMPCAMBIOESTADO '" & _
        gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "','" & _
        vusu & "'"
        
        Con.Execute strSQL

        Con.CommitTrans
        'Dim amensaje As New clsMensaje
        'amensaje.Codigo = CodeMsg.KMESSAGE_INF_DATA_SAVE
        'Informa "", amensaje
        
        MsgBox "El cambio de estado resultó exitoso.", vbOKOnly, "Ordenes de Compra"
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub



Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans
    'If RTrim(txtPorc_IGV.Text) = "" Then
    '    txtPorc_IGV.Text = "0"
    'End If
        strSQL = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "'"
        
        If sTipo = "I" Then
            Rs.Open strSQL, cConnect, adOpenStatic
            optOrdCompra.Value = True
            txtSerOrdComp.Text = Rs(0)
            txtCodOrdComp.Text = Rs(1)
            CARGA_GRID
        Else
            Con.Execute strSQL
        End If
       
        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        Informa "", amensaje
        
'        If sTipo = "I" Then
'            optOrdCompra.Value = True
'            Strsql = "SELECT MAX(Ser_OrdComp) FROM lg_ordcomp"
'            txtSerOrdComp.Text = DevuelveCampo(Strsql, cCONNECT)
'            Strsql = "SELECT MAX(Cod_OrdComp) FROM lg_ordcomp WHERE Ser_OrdComp ='" & Trim(txtSerOrdComp.Text) & "'"
'            txtCodOrdComp.Text = DevuelveCampo(Strsql, cCONNECT)
'            CARGA_GRID
'        End If
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        strSQL = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "'"
        
        Con.Execute strSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

'Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Call CARGA_DATOS
'End Sub

Sub BUSCA_PROVEEDOR(Tipo As Integer, Ubic As Integer)
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCodProveedor.Text & "'"
                    txtDesProveedor.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDesProveedor.Text & "'"
                    'txtCodProveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                    FunctBuscar.SetFocus
                Else
                    strSQL = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
                    txtDes_Proveedor.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDes_Proveedor.Text & "'"
                    'txtCod_Proveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                    If cboCod_CondVent.Enabled = True Then
                        
                        'Aqui poscionaremos por defecto al cond, venta del prov
                        strSQL = "SELECT Cod_CondVENT FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSQL, cConnect), 2, cboCod_CondVent)
                        strSQL = "SELECT Cod_Descuento FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSQL, cConnect), 2, cboCod_Descuento)
                        
                        cboCod_CondVent.SetFocus
                    End If
                End If
                'FunctBuscar.SetFocus
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.sQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDesProveedor.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDes_Proveedor.Text) & "%'"
                End If
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCodProveedor.Text = Trim(codigo)
                        txtDesProveedor.Text = Trim(Descripcion)
                        FunctBuscar.SetFocus
                        codigo = ""
                        Descripcion = ""
                    Else
                        txtCod_Proveedor.Text = Trim(codigo)
                        txtDes_Proveedor.Text = Trim(Descripcion)
                        
                        'Aqui posicionaremos por defecto al cond, venta del prov
                        strSQL = "SELECT Cod_CondVENT FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSQL, cConnect), 2, cboCod_CondVent)
                        strSQL = "SELECT Cod_Descuento FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSQL, cConnect), 2, cboCod_Descuento)
                        
                        cboCod_CondVent.SetFocus
                    End If
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
                
    End Select
End Sub

Sub BUSCA_GRUPO(Tipo As Integer)
    Dim NombreTabla As String
    Dim CodigoTabla As String
    strSQL = "SELECT Tip_Grupo FROM LG_TIPREQ WHERE Cod_TipRequ='" & varCod_TipRequ & "'"
    If DevuelveCampo(strSQL, cConnect) = "I" Then
        NombreTabla = "ES_GRUPOLOG"
        CodigoTabla = "Cod_GrupoLog"
    Else
        NombreTabla = "ES_GRUPOTEX"
        CodigoTabla = "Cod_GrupoTex"
    End If
    
    
    Select Case Tipo
        Case 1:
                strSQL = "SELECT Des_Grupo FROM " & NombreTabla & " WHERE " & CodigoTabla & " = '" & txtCod_Grupo.Text & "'"
                TxtDes_Grupo.Text = Trim(DevuelveCampo(strSQL, cConnect))
                
                'Strsql = "SELECT " & CodigoTabla & " FROM " & NombreTabla & " WHERE Des_Grupo = '" & txtDes_Grupo.Text & "'"
                'txtCod_Grupo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, Des_Grupo as Descripción FROM " & NombreTabla & " WHERE Des_Grupo LIKE '" & Trim(TxtDes_Grupo.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, Des_Grupo as Descripción FROM " & NombreTabla
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    txtCod_Grupo.Text = Trim(codigo)
                    TxtDes_Grupo.Text = Trim(Descripcion)
                    If cboCod_ProTex.Enabled Then
                        cboCod_ProTex.SetFocus
                        codigo = ""
                        Descripcion = ""
                    End If
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Sub BUSCA_ESTADO(Tipo As Integer)
    'Dim TipGrupo As Integer
    'Strsql = ""
    'TipGrupo = DevuelveCampo(Strsql, cCONNECT)
    
    Select Case Tipo
        Case 1:
                strSQL = "SELECT Des_StaOrdComp FROM LG_STAORDCOMP WHERE  Cod_StaOrdComp = '" & txtCodStaOrdComp.Text & "'"
                txtDesStaOrdComp.Text = Trim(DevuelveCampo(strSQL, cConnect))
                strSQL = "SELECT Cod_StaOrdComp FROM LG_STAORDCOMP WHERE Des_StaOrdComp = '" & txtDesStaOrdComp.Text & "'"
                txtCodStaOrdComp.Text = Trim(DevuelveCampo(strSQL, cConnect))
                FunctBuscar.SetFocus
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT Cod_StaOrdComp as Código, Des_StaOrdComp as Descripción FROM LG_STAORDCOMP WHERE Des_StaOrdComp LIKE '" & txtDesStaOrdComp.Text & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_StaOrdComp as Código, Des_StaOrdComp as Descripción FROM LG_STAORDCOMP"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    txtCodStaOrdComp.Text = Trim(codigo)
                    txtDesStaOrdComp.Text = Trim(Descripcion)
                    FunctBuscar.SetFocus
                    codigo = ""
                    Descripcion = ""
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Private Sub Form_Load()
    Opcion = 1
'    Call FormateaGrid(DGridLista)
    Call CARGA_COMBOS
    Call CARGA_GRID
    Call INHABILITA_DATOS
    
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctBuscar.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    
    VerificaFabrica txtCod_Fabrica, txtNom_Fabrica
    sTituliAbrOP = RTrim(DevuelveCampo("select Titulo_Abr_Orden from TG_Control", cConnect))
    optOP.Caption = sTituliAbrOP
    lblorden.Caption = sTituliAbrOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call CARGA_GRID
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim vRow As Long
Dim vOrdCompBusq As String
On Error GoTo AceptaError:

    Dim varOrigen As String
    Dim varTipFabrica As Integer

    Dim varCambioEstado As Integer
    If Rs_Lista.EOF And Rs_Lista.EOF Then
        MsgBox "Debe seleccionar un registro, para poder acceder a esta opción. Sirvase verificar", vbInformation, "Ordenes de Compra"
        Exit Sub
    End If
    vOrdCompBusq = gexLista.Value(gexLista.Columns("O.C.").Index)
    'vRow = gexLista.RowIndex(gexLista.Row)
    Select Case ActionName
        Case "IMPRESION":
                        
'                        Load frmSelecFamilias
'                        frmSelecFamilias.varSer_OrdComp = Rs_Lista("Ser_OrdComp").Value
'                        frmSelecFamilias.varCod_OrdComp = Rs_Lista("Cod_OrdComp").Value
'                        frmSelecFamilias.CARGA_FAMILIAS
'                        Set frmSelecFamilias.oParent = Me
'                        frmSelecFamilias.Show 1
'
'                        If Me.varCancelImpresion = 1 Then
'                            Exit Sub
'                        End If
                        
                        'MsgBox ("IMPRESION")
                        'varAyuda = CStr(varCod_TipRequ)
                        strSQL = "SELECT Origen FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(txtCod_Proveedor.Text) & "'"
                        varOrigen = DevuelveCampo(strSQL, cConnect)
                        
                        'Strsql = "SELECT tip_fabrica FROM TG_CONTROL"
                        'varTipFabrica = DevuelveCampo(Strsql, cCONNECT)
                        
                        'Strsql = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
                        'varAyuda = IIf(IsNull(DevuelveCampo(Strsql, cConnect)), "", DevuelveCampo(Strsql, cConnect))
                        
                        strSQL = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
                        Dim oo As Object
                        Set oo = CreateObject("excel.application")
'                        oo.Workbooks.Open vRuta & "\RptOCompra.xlt"
                        If varOrigen = "N" Then
                            oo.Workbooks.Open vRuta & "\RptOCompra.xlt"
                            'oo.Workbooks.Open App.Path & "\RptOCompra.xlt"
                        Else
                            oo.Workbooks.Open vRuta & "\RptOCompraIng.xlt"
                            'oo.Workbooks.Open App.Path & "\RptOCompraIng.xlt"
                        End If
                        oo.Visible = True
                        'oo.Run "PRUEBA", CStr(varCod_Cliente), CStr(varCod_Fabrica), CStr(txtCod_EstCli.Text), CStr(txtAbr_Cliente.Text & " - " & txtNom_Cliente.Text), CStr(txtAbr_Fabrica.Text & " - " & txtNom_Fabrica.Text), CStr(txtCod_EstCli.Text & " - " & txtDes_EstCli.Text), cCONNECT
                        oo.Run "REPORTE", gexLista.Value(gexLista.Columns("Ser_OrdComp").Index), gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), vusu, txtCod_Grupo.Text & " - " & TxtDes_Grupo.Text, DevuelveCampo(strSQL, cConnect), vemp, Me.varCod_TipRequ, cConnect
                        Screen.MousePointer = vbNormal
                        oo.Visible = True
                        Set oo = Nothing
                        
                        
        Case "CAMBIOESTADO":
                        varCambioEstado = MsgBox("¿Esta usted seguro de cambiar el estado al registro seleccionado?", vbInformation + vbYesNo, "Ordenes de Compra")
                        If varCambioEstado = vbYes Then
                            Call CAMBIO_ESTADO
                            Call CARGA_GRID
                        End If
        Case "DETALLE":
                        strSQL = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
                        Load frmOrdCompItem
                        frmOrdCompItem.Caption = "Detalles de la Orden de Compra: " & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & " - " & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                        frmOrdCompItem.varTip_Presentacion = DevuelveCampo(strSQL, cConnect)
                        frmOrdCompItem.varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                        frmOrdCompItem.varCod_OrdComp = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                        frmOrdCompItem.varCod_ClaOrdComp = gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index)
                        frmOrdCompItem.varPorc_IGV = gexLista.Value(gexLista.Columns("I.G.V.").Index)
                        frmOrdCompItem.varCod_Descuento = gexLista.Value(gexLista.Columns("Cod_Descuento").Index)
                        frmOrdCompItem.varCod_Proveedor = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
                        frmOrdCompItem.varCod_StaOrdComp = gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index)
                        frmOrdCompItem.varCod_GrupoTex = gexLista.Value(gexLista.Columns("Cod.Grupo").Index)
                        frmOrdCompItem.varDes_Grupo = Trim(TxtDes_Grupo.Text)
                        frmOrdCompItem.varCod_TipRequ = varCod_TipRequ
                        frmOrdCompItem.CARGA_GRID
                        frmOrdCompItem.Show 1
         Case "HILREQ"
                        MUESTRA_HILOS
         Case "NUMIMPORT"
                        If gexLista.RowCount = 0 Then Exit Sub
                        If Trim(gexLista.Value(gexLista.Columns("Num_Importacion").Index)) <> "" Then
                            MsgBox "Número de Importacion ya generado", vbExclamation, "Orden Compra"
                            Exit Sub
                        End If
                        Load FrmImportaciones
                        FrmImportaciones.sSer_OrdComp = Me.gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                        FrmImportaciones.sCod_Ordcomp = Me.gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                        FrmImportaciones.TxtSer_OrdComp = Me.gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                        FrmImportaciones.txtCod_OrdComp = Me.gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                        FrmImportaciones.Show 1
                        Set FrmImportaciones = Nothing
                        Call CARGA_GRID
    End Select
    'gexLista.Row = vRow
    Call gexLista.Find(3, jgexEqual, vOrdCompBusq)
    Exit Sub
AceptaError:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
    Set oo = Nothing

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
CERRAR_ORDCOMP
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim eliminar As Integer
    Dim vRow As Long
    vRow = gexLista.Row
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            CboEstado.Enabled = True
            txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "MODIFICAR"
        
            If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
                If gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index) <> "P" Then
                    MsgBox "El estado del registro no permite la modificación. Sirvase verificar", vbInformation, "Ordenes de Compra"
                    Exit Sub
                End If
            End If
        
            sTipo = "U"
            HABILITA_DATOS
            txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "ELIMINAR"
        
            If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
                If gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index) <> "P" Then
                    MsgBox "El estado del registro no permite la eliminación. Sirvase verificar", vbInformation, "Ordenes de Compra"
                    Exit Sub
                End If
            End If
        
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    gexLista.Row = vRow - 1
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                gexLista.Enabled = True
                If sTipo = "I" Then
                    gexLista.MoveLast
                    strSQL = DevuelveCampo("select origen from lg_proveedor where cod_proveedor='" & txtCod_Proveedor.Text & "'", cConnect)
                    If strSQL = "E" Then
                        Pregunta = MsgBox("¿Desea generar Num. Importación?", vbYesNo)
                        If Pregunta = vbYes Then
                            Call FunctButt1_ActionClick(0, 0, "NUMIMPORT")
                        End If
                    End If
                Else
                    gexLista.Row = vRow
                End If
                sTipo = ""
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            gexLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub OpGrupo_Click()
    FraOrdComp.Visible = False
    FraProveedor.Visible = False
    FraEstado.Visible = False
    FraGrupo.Visible = True
    fraOP.Visible = False
    txtCod_OrdPro.Text = ""

    TxtCodGrupo.Text = ""
    TxtDesGrupo.Text = ""
    TxtCodGrupo.SetFocus
    Opcion = 4
End Sub

Private Sub OpLog_Click()
    OpTex.Value = False
End Sub

Private Sub optEstado_Click()
    FraOrdComp.Visible = False
    FraProveedor.Visible = False
    FraGrupo.Visible = False
    FraEstado.Visible = True
    fraOP.Visible = False
    txtCod_OrdPro.Text = ""
    
    txtCodStaOrdComp.Text = ""
    txtDesStaOrdComp.Text = ""
    txtCodStaOrdComp.SetFocus
    Opcion = 3
End Sub

Private Sub OpTex_Click()
    OpLog.Value = False
End Sub

Private Sub optOP_Click()
    fraOP.Visible = True
    FraProveedor.Visible = False
    FraOrdComp.Visible = False
    FraGrupo.Visible = False
    FraEstado.Visible = False
    txtCod_OrdPro.Text = ""
    txtCod_OrdPro.SetFocus
    Opcion = 5
End Sub

Private Sub optOrdCompra_Click()
    FraProveedor.Visible = False
    FraEstado.Visible = False
    FraGrupo.Visible = False
    FraOrdComp.Visible = True
    fraOP.Visible = False
    txtCod_OrdPro.Text = ""
    txtDes_OrdPro.Text = ""
    
    txtSerOrdComp.Text = ""
    txtCodOrdComp.Text = ""
    txtSerOrdComp.SetFocus
    Opcion = 1
End Sub

Private Sub optProveedor_Click()
    FraEstado.Visible = False
    FraOrdComp.Visible = False
    FraGrupo.Visible = False
    FraProveedor.Visible = True
    fraOP.Visible = False
    txtCod_OrdPro.Text = ""
    
    txtCodProveedor.Text = ""
    txtDesProveedor.Text = ""
    txtCodProveedor.SetFocus
     
    Opcion = 2
End Sub

Private Sub TxtCodGrupo_Change()
    If Trim(codigo) <> "" Or Trim(TxtCodGrupo.Text) = "" Then
        Exit Sub
    End If
        Load frmBuscaGrupo
        Set frmBuscaGrupo.oParent = Me
        If OpTex.Value = True Then
            frmBuscaGrupo.varTipo = "1"
        Else
            frmBuscaGrupo.varTipo = "2"
        End If
        frmBuscaGrupo.txtCod_GrupoTex = Me.TxtCodGrupo
        frmBuscaGrupo.CARGA_GRID
        frmBuscaGrupo.Show 1

        Set frmBuscaGrupo = Nothing

        If Trim(codigo) <> "" Then
            Me.TxtCodGrupo.Text = codigo
            Me.TxtDesGrupo.Text = Descripcion
            FunctBuscar.SetFocus
        End If
        codigo = ""
        Descripcion = ""

End Sub

Private Sub txtCodOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCodOrdComp.Text = Right("000000" & Trim(txtCodOrdComp.Text), 6)
        FunctBuscar.SetFocus
    End If
End Sub

Private Sub txtCodOrdComp_LostFocus()
    txtCodOrdComp.Text = Right("000000" & Trim(txtCodOrdComp.Text), 6)
    FunctBuscar.SetFocus
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCodProveedor.Text) <> "" Then
            txtCodProveedor.Text = Right("000000000000" & txtCodProveedor.Text, 12)
            Call BUSCA_PROVEEDOR(1, 1)
        End If
    End If
End Sub


Private Sub txtDesProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesProveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2, 1)
        End If
    End If
End Sub

Private Sub Txtcod_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Proveedor.Text) <> "" Then
            txtCod_Proveedor.Text = Right("000000000000" & txtCod_Proveedor.Text, 12)
            Call BUSCA_PROVEEDOR(1, 2)
        End If
    End If
End Sub

Private Sub TxtDes_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Proveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2, 2)
        End If
    End If
End Sub

Private Sub txtCodStaOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCodStaOrdComp.Text) <> "" Then
            Call BUSCA_ESTADO(1)
        End If
    End If
End Sub
Private Sub txtDesStaOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesStaOrdComp.Text) <> "" Then
            Call BUSCA_ESTADO(2)
        End If
    End If
End Sub

Private Sub cmdBusEstado_Click()
    Call BUSCA_ESTADO(3)
End Sub

Private Sub txtCod_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Grupo.Text) <> "" Then
            txtCod_Grupo.Text = Right("00000000" & txtCod_Grupo.Text, 8)
            Call BUSCA_GRUPO(1)
        End If
    End If

End Sub

Private Sub txtDes_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtDes_Grupo.Text) <> "" Then
            Call BUSCA_GRUPO(2)
        End If
    End If

End Sub

Private Sub cmdBuscaGrupo_Click()
    Call BUSCA_GRUPO(3)
End Sub

Private Sub txtSerOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSerOrdComp.Text = Right("000" & Trim(txtSerOrdComp.Text), 3)
        txtCodOrdComp.SetFocus
    End If
End Sub

Private Sub txtSerOrdComp_LostFocus()
    txtSerOrdComp.Text = Right("000" & Trim(txtSerOrdComp.Text), 3)
End Sub

Private Sub CERRAR_ORDCOMP()
    Dim Con As New ADODB.Connection
    Dim Message As Integer
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    Message = MsgBox("¿Esta usted seguro que desea Abrir/Cerrar la O/C seleccionada?", vbInformation + vbYesNo, "Orden de Compra")
    If Message = vbYes Then
        Con.BeginTrans

        strSQL = "EXEC UP_MAN_ORDCOMP_ABRIRCERRAR '" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        vusu & "'"
        
        Con.Execute strSQL

        Con.CommitTrans
        
        MsgBox "La Orden de Compra se Modificó satisfactoriamente", vbOKOnly, "Ordenes de Compra"
        CARGA_GRID
    End If
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "CERRAR_ORDCOMP"
End Sub

Sub MUESTRA_HILOS()
On Error GoTo Muestra_DatosErr
Dim Rs As New ADODB.Recordset

Rs.Open "select * from lg_claordcomp where cod_claOrdComp='" & gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index) & "'", cConnect, adOpenStatic
If Rs.RecordCount Then
    If Rs.Fields("Tip_Item").Value = "T" And Rs.Fields("Tip_Presentacion").Value = "C" Then
        frmHiladosRequeridos.varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
        frmHiladosRequeridos.varCod_OrdComp = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
        frmHiladosRequeridos.varCod_Proveedor = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
        frmHiladosRequeridos.Show 1
    End If
End If

Set Rs = Nothing
    Exit Sub
Muestra_DatosErr:
    Set Rs = Nothing
    ErrorHandler Err, "MUESTRA_HILOS"
End Sub




Private Sub VerificaFabrica(ByRef objFabrica As TextBox, ByRef objNombreFabrica As TextBox)
    Dim sSQl As String
    Dim iRet As String
    
    sSQl = "SELECT count(*) FROM TG_Fabrica "
    iRet = DevuelveCampo(sSQl, cConnect)
    If iRet = 1 Then
        sSQl = "SELECT Cod_Fabrica FROM TG_Fabrica "
        objFabrica.Text = DevuelveCampo(sSQl, cConnect)
        
        sSQl = "SELECT Nom_Fabrica FROM TG_Fabrica "
        objNombreFabrica.Text = DevuelveCampo(sSQl, cConnect)
        objFabrica.Enabled = False
        objNombreFabrica.Enabled = False
        
    End If
End Sub


Private Sub txtCod_OrdPro_GotFocus()
    SelectionText txtCod_OrdPro
End Sub

Private Sub txtCod_Ordpro_KeyPress(KeyAscii As Integer)
    Dim iLen As Integer
    Dim sSQl As String
        
    If KeyAscii = vbKeyReturn Then
        If RTrim(txtCod_OrdPro.Text) <> "" Then
            
            txtCod_OrdPro.Text = LPadr(txtCod_OrdPro, 5, "0")
        
            If BuscaPedido(txtCod_OrdPro.Text) Then
                FunctBuscar.SetFocus
            End If
        End If
    End If

End Sub


Private Function BuscaPedido(ByVal sCod_Pedido As String) As Boolean
    Dim sSQl As String
    Dim mRs As ADODB.Recordset
    
    sSQl = "SM_MUESTRA_Cod_OrdPro '" & txtCod_Fabrica.Text & "', '" & txtCod_OrdPro.Text & "'"
    Set mRs = GetRecordset(cConnect, sSQl)
    
    If mRs.EOF Then
        MsgBox RTrim(sTituliAbrOP) & " NO EXISTE", vbCritical
        txtCod_OrdPro.SetFocus
        mRs.Close
        Set mRs = Nothing
        Exit Function
    Else
        txtCod_OrdPro.Text = mRs!Cod_Ordpro
        txtDes_OrdPro.Text = mRs!Des_EstPro
    End If
    mRs.Close
    Set mRs = Nothing
    BuscaPedido = True
End Function


Sub GENERA_NUMIMPORTACION()
Dim amensaje As New clsMessages
    On Error GoTo Salvar_DatosErr
    
        strSQL = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "'"
        
        Call ExecuteSQL(cConnect, strSQL)
        
        amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        Informa "", amensaje
    Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "GENERA_NUMIMPORTACION"
End Sub



