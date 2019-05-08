VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantItemServicios 
   Caption         =   "Control de Bordados Estampados y Aplicaciones"
   ClientHeight    =   9312
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   12552
   LinkTopic       =   "Form1"
   ScaleHeight     =   9312
   ScaleWidth      =   12552
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Identificador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   59
      Top             =   9930
      Width           =   6615
      Begin VB.ComboBox CboIde_PO 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   225
         Width           =   585
      End
      Begin VB.ComboBox cboIde_Destino 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   225
         Width           =   585
      End
      Begin VB.ComboBox cboIde_Color 
         Height          =   315
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   225
         Width           =   585
      End
      Begin VB.ComboBox cboIde_EsCli 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   240
         Width           =   585
      End
      Begin VB.ComboBox cboIde_Talla 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "P.O. :"
         Height          =   195
         Left            =   5520
         TabIndex        =   69
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         Height          =   195
         Left            =   4080
         TabIndex        =   68
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label16 
         Caption         =   "Color Cliente :"
         Height          =   390
         Left            =   2760
         TabIndex        =   67
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label17 
         Caption         =   "Estilo Cliente :"
         Height          =   375
         Left            =   1380
         TabIndex        =   66
         Top             =   135
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Talla :"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   290
         Width           =   435
      End
   End
   Begin VB.ComboBox cboCod_GruItem 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   9510
      Width           =   1815
   End
   Begin VB.ComboBox cboCod_MotPrePro 
      Height          =   315
      Left            =   5430
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   9390
      Width           =   1815
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      Left            =   120
      TabIndex        =   12
      Tag             =   "Detail"
      Top             =   4770
      Width           =   12300
      Begin VB.TextBox txtPrecioComercial 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1800
         TabIndex        =   85
         Text            =   "0"
         Top             =   3000
         Width           =   1812
      End
      Begin VB.TextBox txtTecnicaEstampado 
         Height          =   288
         Left            =   5400
         TabIndex        =   84
         Top             =   3000
         Width           =   6132
      End
      Begin VB.ComboBox cboModoProceso 
         Height          =   288
         Left            =   5280
         TabIndex        =   73
         Top             =   1320
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpFechaUbicacion 
         Height          =   336
         Left            =   5280
         TabIndex        =   70
         Top             =   960
         Width           =   1836
         _ExtentX        =   3239
         _ExtentY        =   572
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   58720257
         CurrentDate     =   39139
      End
      Begin VB.TextBox txtUbicacion 
         Height          =   315
         Left            =   1440
         TabIndex        =   52
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   11376
         Begin VB.TextBox txtUMPro 
            Height          =   285
            Left            =   10080
            MaxLength       =   2
            TabIndex        =   53
            Top             =   195
            Width           =   1095
         End
         Begin VB.TextBox txtPrecio 
            Height          =   285
            Left            =   1350
            MaxLength       =   8
            TabIndex        =   46
            Top             =   525
            Width           =   1245
         End
         Begin VB.TextBox txtObservaciones_Proveedor 
            Height          =   285
            Left            =   5040
            TabIndex        =   45
            Top             =   525
            Width           =   6210
         End
         Begin VB.TextBox txtCodItemPro 
            Height          =   285
            Left            =   7920
            MaxLength       =   15
            TabIndex        =   43
            Top             =   195
            Width           =   1335
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   285
            Left            =   1320
            TabIndex        =   40
            Top             =   195
            Width           =   1245
         End
         Begin VB.TextBox txtNombreProveedor 
            Height          =   285
            Left            =   2685
            TabIndex        =   39
            Top             =   195
            Width           =   4200
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones del Proveedor"
            Height          =   195
            Left            =   2670
            TabIndex        =   48
            Top             =   555
            Width           =   2100
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Precio Cotizado$:"
            Height          =   195
            Left            =   135
            TabIndex        =   47
            Top             =   555
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "UniMed"
            Height          =   195
            Left            =   9315
            TabIndex        =   44
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Código del Prov:"
            Height          =   390
            Left            =   6975
            TabIndex        =   42
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código/Nombre :"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFC&
         Caption         =   "Imagen del Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   8040
         TabIndex        =   36
         Top             =   360
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   1935
            Left            =   600
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.TextBox txtcoditem 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   20
         Top             =   200
         Width           =   915
      End
      Begin VB.ComboBox cboCod_FamItem 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   570
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_ClaItem 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1290
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_Origen 
         Height          =   288
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_UniMed 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   920
         Width           =   1815
      End
      Begin VB.TextBox txtDesItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   100
         TabIndex        =   15
         Top             =   200
         Width           =   4680
      End
      Begin VB.ComboBox cboFlg_Status 
         Height          =   288
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtComentario 
         Height          =   495
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   2400
         Width           =   5205
      End
      Begin VB.Label Label25 
         Caption         =   "Precio Comercial($)"
         Height          =   252
         Left            =   120
         TabIndex        =   87
         Top             =   3000
         Width           =   1452
      End
      Begin VB.Label Label24 
         Caption         =   "Tecnica Estampado"
         Height          =   252
         Left            =   3840
         TabIndex        =   86
         Top             =   3000
         Width           =   1572
      End
      Begin VB.Label Label19 
         Caption         =   "Modo Proceso :"
         Height          =   288
         Left            =   3840
         TabIndex        =   72
         Top             =   1320
         Width           =   1128
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha :"
         Height          =   276
         Left            =   3840
         TabIndex        =   71
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label10 
         Caption         =   "Ubicación"
         Height          =   252
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Familia Item:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Tag             =   "Mat. Prima :"
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblCod_Item 
         AutoSize        =   -1  'True
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "Hilado :"
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Medida :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   120
         TabIndex        =   25
         Tag             =   "Porcentaje :"
         Top             =   960
         Width           =   1368
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Clase de Item :"
         Height          =   192
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1056
      End
      Begin VB.Label Label13 
         Caption         =   "Origen :"
         Height          =   252
         Left            =   3840
         TabIndex        =   23
         Top             =   1680
         Width           =   732
      End
      Begin VB.Label Label18 
         Caption         =   "Status :"
         Height          =   252
         Left            =   3840
         TabIndex        =   22
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Comentario :"
         Height          =   192
         Left            =   132
         TabIndex        =   21
         Top             =   2520
         Width           =   888
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   105
      TabIndex        =   11
      Top             =   1425
      Width           =   12330
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Height          =   1800
         Left            =   4440
         TabIndex        =   80
         Top             =   840
         Visible         =   0   'False
         Width           =   3630
         Begin VB.ComboBox cboEsta 
            Height          =   315
            Left            =   1065
            TabIndex        =   81
            Top             =   405
            Width           =   2385
         End
         Begin FunctionsButtons.FunctButt FunctButt3 
            Height          =   510
            Left            =   690
            TabIndex        =   83
            Top             =   1080
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   910
            Custom          =   $"frmMantItemServicios.frx":0000
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Estado"
            Height          =   345
            Left            =   375
            TabIndex        =   82
            Top             =   420
            Width           =   765
         End
      End
      Begin VB.Frame fraImprimir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Impresión"
         Height          =   1920
         Left            =   4522
         TabIndex        =   74
         Top             =   780
         Visible         =   0   'False
         Width           =   3510
         Begin VB.OptionButton opttodas 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Todas"
            Height          =   210
            Left            =   1125
            TabIndex        =   76
            Top             =   285
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton optPendientes 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pendientes"
            Height          =   210
            Left            =   1140
            TabIndex        =   75
            Top             =   600
            Width           =   1785
         End
         Begin FunctionsButtons.FunctButt FunctButt2 
            Height          =   510
            Left            =   540
            TabIndex        =   77
            Top             =   1080
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   910
            Custom          =   $"frmMantItemServicios.frx":008D
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
      End
      Begin GridEX20.GridEX DGridLista 
         Height          =   2505
         Left            =   135
         TabIndex        =   54
         Top             =   180
         Width           =   11430
         _ExtentX        =   20151
         _ExtentY        =   4424
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         ColumnHeaderHeight=   288
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantItemServicios.frx":0126
         Column(2)       =   "frmMantItemServicios.frx":01EE
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantItemServicios.frx":0292
         FormatStyle(2)  =   "frmMantItemServicios.frx":03CA
         FormatStyle(3)  =   "frmMantItemServicios.frx":047A
         FormatStyle(4)  =   "frmMantItemServicios.frx":052E
         FormatStyle(5)  =   "frmMantItemServicios.frx":0606
         FormatStyle(6)  =   "frmMantItemServicios.frx":06BE
         ImageCount      =   0
         PrinterProperties=   "frmMantItemServicios.frx":079E
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      TabIndex        =   9
      Top             =   -15
      Width           =   12255
      Begin VB.TextBox txtCodFamilia 
         Height          =   300
         Left            =   9810
         TabIndex        =   79
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtNombreProveedor2 
         Height          =   285
         Left            =   7680
         TabIndex        =   8
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   270
         Left            =   2115
         TabIndex        =   35
         Tag             =   "..."
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton cmdBusEstado 
         Caption         =   "..."
         Height          =   270
         Left            =   2130
         TabIndex        =   34
         Tag             =   "..."
         Top             =   975
         Width           =   285
      End
      Begin VB.TextBox txtDesStatus 
         Height          =   285
         Left            =   2445
         TabIndex        =   7
         Top             =   975
         Width           =   4200
      End
      Begin VB.TextBox txtCodStatus 
         Height          =   285
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   6
         Top             =   975
         Width           =   1005
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Estado"
         Height          =   300
         Left            =   150
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   300
         Left            =   150
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optcliente 
         Caption         =   "Cliente"
         Height          =   300
         Left            =   150
         TabIndex        =   31
         Top             =   255
         Width           =   855
      End
      Begin VB.TextBox txtcliente 
         Height          =   285
         Left            =   1110
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txttemporada 
         Height          =   285
         Left            =   5190
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdBusTemporada 
         Caption         =   "..."
         Height          =   285
         Left            =   5835
         TabIndex        =   29
         Top             =   240
         Width           =   360
      End
      Begin VB.TextBox txtNom_TemCli 
         Height          =   300
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   2880
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   2445
         TabIndex        =   1
         Top             =   225
         Width           =   1695
      End
      Begin VB.CommandButton cmdBusItem 
         Caption         =   "..."
         Height          =   270
         Left            =   2100
         TabIndex        =   28
         Tag             =   "..."
         Top             =   615
         Width           =   300
      End
      Begin VB.TextBox txtdes_item 
         Height          =   285
         Left            =   2430
         TabIndex        =   5
         Top             =   600
         Width           =   4200
      End
      Begin VB.TextBox txtcod_item 
         Height          =   285
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   1005
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   10965
         TabIndex        =   10
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   868
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox txtCodProveedor2 
         Height          =   285
         Left            =   8565
         TabIndex        =   49
         Top             =   660
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label22 
         Caption         =   "Familia"
         Height          =   225
         Left            =   9240
         TabIndex        =   78
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "Proveedor"
         Height          =   405
         Left            =   7680
         TabIndex        =   50
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   4350
         TabIndex        =   30
         Top             =   255
         Width           =   810
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   90
      TabIndex        =   37
      Top             =   4215
      Width           =   12345
      _ExtentX        =   21781
      _ExtentY        =   910
      Custom          =   $"frmMantItemServicios.frx":0976
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1100
      ControlHeigth   =   490
      ControlSeparator=   20
   End
   Begin VB.Label Label9 
      Caption         =   "Grupo de Item :"
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   9630
      Width           =   1095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Preproduc :"
      Height          =   195
      Left            =   3990
      TabIndex        =   56
      Top             =   9510
      Width           =   1350
   End
End
Attribute VB_Name = "frmMantItemServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim Opcion As Integer
Dim sTipo As String
Dim varCod_item As String
Dim vCod_hiltel As String, sConta As Integer
Dim item As String
Dim sStrSQL As String

Private Sub cbogrupo_Click()
    Call CargaLista
End Sub

Private Sub cboCod_FamItem_Click()
    Dim StrSQL As String
    'Combo Grupo Item
    cboCod_GruItem.Clear
    'StrSQL = "SELECT des_famgruite + space(100) + Cod_Gruitem FROM LG_FamGruIte WHERE Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
    'Call LlenaCombo(cboCod_GruItem, StrSQL, cCONNECT)
    
    'StrSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
    'If Trim(cboCod_FamItem.Text) <> "" Then
        'If DevuelveCampo(StrSQL, cCONNECT) = "M" And (sTipo = "I" Or sTipo = "U") Then
        '    sConta = DevuelveCampo("select count(*) from LG_Autorizacion_Campos where cod_usuario='" & vusu & "' and Tipo_Autorizacion ='1'", cCONNECT)
        '    If sConta > 0 Then
        '        HABILITA_CARACMXT True
        '    End If
        'Else
        '    HABILITA_CARACMXT False
        'End If
    'End If
End Sub



Private Sub Cmd_Aceptar_Click()

 Call CargaLista
 
End Sub








Private Sub cmdBusEstado_Click()
 Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT flg_status as Código, des_status as Descripción FROM TG_StaDes "
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCodStatus.Text = Codigo
        txtDesStatus.Text = Descripcion
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdBusItem_Click()
    Dim StrSQL As String
    If Trim(txtcod_item.Text) <> "" Then
        StrSQL = "SELECT Cod_Item as Código, Des_Item as Descripción FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
    Else
        If Len(Trim(txtdes_item.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbExclamation)
            Exit Sub
        Else
            StrSQL = "SELECT Cod_Item as Código, Des_Item as Descripción  FROM LG_ITEM WHERE Des_Item LIKE '" & Trim(txtdes_item.Text) & "%'"
        End If
    End If
    
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = StrSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcod_item.Text = Codigo
        txtdes_item.Text = Descripcion
        FunctBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub



 
 
Private Sub xDGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral3
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcliente.Text = Codigo
        txtNom_Cliente.Text = Descripcion
        Codigo = ""
        txttemporada.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub


Private Sub cmdBusTemporada_Click()
    Dim oTipo As New frmBusqGeneral3
    Dim Rs As New ADODB.Recordset
    Dim StrSQL As String
    Set oTipo.oParent = Me
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    oTipo.sQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "'  AND cod_temcli like '%" & txttemporada.Text & "%'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txttemporada.Text = Codigo
        txtNom_TemCli.Text = Descripcion
        Codigo = ""
        FunctBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub DGridLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If Me.DGridLista.RowCount > 0 And (Not Me.DGridLista.IsGroupItem(Me.DGridLista.Row)) Then
    item = Me.DGridLista.Value(Me.DGridLista.Columns("cod_itemx").Index)
    'Me.Tag = item
    Call CargaDatos
    End If
    
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    Call CargaCombos
    Opcion = 1
        
    dtpFechaUbicacion.Value = Date
    dtpFechaUbicacion.Value = Null
    
    INHABILITA_DATOS
   
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub



 
Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call CargaLista
    
    'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo AceptarErr

    Dim StrSQL As String
    Dim vericono As Integer
    Select Case ActionName
 
    Case "ADICIONAR"
                If optcliente.Value = False Then
                    MsgBox "Debe Ingresar Cliente Tempordada antes de acceder a esta opción !! ", vbInformation
                    Exit Sub
                End If
                
                If optcliente.Value = True And (RTrim(txtcliente.Text) = "" Or RTrim(txttemporada) = "") Then
                    MsgBox "Debe Ingresar Cliente /Tempordada ", vbInformation
                    Exit Sub
                End If
                
                
                Load frmAdicionarModificarItems
                Set frmAdicionarModificarItems.oParent = Me
                frmAdicionarModificarItems.Caption = "Adicionar Item"
                frmAdicionarModificarItems.Abr_Cliente = Trim(txtcliente.Text)
                frmAdicionarModificarItems.sTemporada = Trim(txttemporada.Text)
                frmAdicionarModificarItems.Opcion = Opcion
                frmAdicionarModificarItems.sTipo = "I"
                frmAdicionarModificarItems.txtCodUM = "UN"
                frmAdicionarModificarItems.txtCodClase = "P"
                frmAdicionarModificarItems.txtCodMotivo = "PD"
                frmAdicionarModificarItems.txtCodOrigen = "L"
                frmAdicionarModificarItems.Show 1
                Set frmAdicionarModificarItems = Nothing
                
     Case "CAMBIOESTADO"
     
     
        If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
               
                  
                   
                   If Trim(DGridLista.Value(DGridLista.Columns("FLG_STATUS_UBICACION").Index)) <> "P" Then
                     If MsgBox("Esta seguro de cambiar de estado", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
                        Frame4.Visible = False
                        StrSQL = " exec ES_Cambia_Status_Ubicacion '" & Trim(DGridLista.Value(DGridLista.Columns("Cod_Itemx").Index)) & "','P' "
                        Call ExecuteSQL(cCONNECT, StrSQL)
                        Call CargaLista
                        
                      End If
                   Else
                        DGridLista.Enabled = False
                        Frame4.Visible = True
                   End If
               
                
        Else
                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
        End If
     
    
     Case "MODIFICAR"
     
     
        If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
               
                Load frmAdicionarModificarItems
                Set frmAdicionarModificarItems.oParent = Me
                frmAdicionarModificarItems.Caption = "Modificar Item"
                frmAdicionarModificarItems.Abr_Cliente = txtcliente.Text
                frmAdicionarModificarItems.sTemporada = txttemporada.Text
                frmAdicionarModificarItems.Opcion = Opcion
                frmAdicionarModificarItems.sTipo = "U"
                frmAdicionarModificarItems.txtcoditem = RTrim(DGridLista.Value(DGridLista.Columns("Cod_Itemx").Index))
                frmAdicionarModificarItems.txtDesItem = RTrim(DGridLista.Value(DGridLista.Columns("Des_Item").Index))
                frmAdicionarModificarItems.txtCodFamilia = RTrim(DGridLista.Value(DGridLista.Columns("cod_FamItem").Index))
                frmAdicionarModificarItems.txtDesFamilia = RTrim(DGridLista.Value(DGridLista.Columns("des_famitem").Index))
                frmAdicionarModificarItems.txtCodUM = RTrim(DGridLista.Value(DGridLista.Columns("cod_UniMed").Index))
                frmAdicionarModificarItems.txtDesUM = RTrim(DGridLista.Value(DGridLista.Columns("Des_UniMed").Index))
                frmAdicionarModificarItems.txtCodClase = RTrim(DGridLista.Value(DGridLista.Columns("cod_ClaItem").Index))
                frmAdicionarModificarItems.txtDesClase = RTrim(DGridLista.Value(DGridLista.Columns("des_claitem").Index))
                frmAdicionarModificarItems.txtCodGrupo = RTrim(DGridLista.Value(DGridLista.Columns("cod_GruItem").Index))
                frmAdicionarModificarItems.txtDesGrupo = RTrim(DGridLista.Value(DGridLista.Columns("des_famgruite").Index))
                frmAdicionarModificarItems.txtCodStatus = RTrim(DGridLista.Value(DGridLista.Columns("Flg_Status").Index))
                frmAdicionarModificarItems.txtDesStatus = RTrim(DGridLista.Value(DGridLista.Columns("des_status").Index))
                frmAdicionarModificarItems.txtCodTipoVersion = RTrim(DGridLista.Value(DGridLista.Columns("Tip_version").Index))
                'frmAdicionarModificarItems.txtDesTipoVersion = RTrim(DGridLista.Value(DGridLista.Columns("Descripcion").Index))
                frmAdicionarModificarItems.TxtModo = RTrim(DGridLista.Value(DGridLista.Columns("Flg_ModoProceso").Index))
                frmAdicionarModificarItems.TxtDes_modo = RTrim(DGridLista.Value(DGridLista.Columns("Des_ModoProceso").Index))
                frmAdicionarModificarItems.txtCodMotivo = RTrim(DGridLista.Value(DGridLista.Columns("Cod_MotPrePro").Index))
                frmAdicionarModificarItems.txtDesMotivo = RTrim(DGridLista.Value(DGridLista.Columns("des_motprepro").Index))
                frmAdicionarModificarItems.txtCodOrigen = RTrim(DGridLista.Value(DGridLista.Columns("cod_origen").Index))
                frmAdicionarModificarItems.txtDesOrigen = RTrim(DGridLista.Value(DGridLista.Columns("des_origen").Index))
                frmAdicionarModificarItems.txtUbicacion = RTrim(DGridLista.Value(DGridLista.Columns("Ubicacion").Index))
                frmAdicionarModificarItems.txtComentario = RTrim(DGridLista.Value(DGridLista.Columns("Comentario").Index))
                frmAdicionarModificarItems.txtCodProveedor = RTrim(DGridLista.Value(DGridLista.Columns("proveedor").Index))
                frmAdicionarModificarItems.txtNombreProveedor = RTrim(DGridLista.Value(DGridLista.Columns("des_proveedor").Index))
                frmAdicionarModificarItems.txtPrecio = RTrim(DGridLista.Value(DGridLista.Columns("Pre_cotizado_proveedor").Index))
                frmAdicionarModificarItems.txtObservacionesProv = RTrim(DGridLista.Value(DGridLista.Columns("Observaciones_proveedor").Index))
                frmAdicionarModificarItems.txtCodItemProv = RTrim(DGridLista.Value(DGridLista.Columns("cod_itemProv").Index))
                frmAdicionarModificarItems.txtUniMedProv = RTrim(DGridLista.Value(DGridLista.Columns("cod_unimedprov").Index))
                frmAdicionarModificarItems.txtDirIcono = RTrim(DGridLista.Value(DGridLista.Columns("Dir_Icono").Index))
                frmAdicionarModificarItems.strImagenCambio = RTrim(DGridLista.Value(DGridLista.Columns("Dir_Icono").Index))
                
                
                
                frmAdicionarModificarItems.txtPrecioComercial = RTrim(DGridLista.Value(DGridLista.Columns("Precio_Cotizacion_Artes").Index))
                frmAdicionarModificarItems.txtTecnicaEstampado = RTrim(DGridLista.Value(DGridLista.Columns("Tecnica_Estampado").Index))
                
                'frmAdicionarModificarItems.TxtIdeTalla = RTrim(DGridLista.Value(DGridLista.Columns("Ide_Talla").Index))
                
                'Call BuscaCombo(RTrim(DGridLista.Value(DGridLista.Columns("Ide_Talla").Index)), 2, cboIde_TallaX)
                
                
                Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Talla").Index), 2, frmAdicionarModificarItems.cboIde_TallaX)
                Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Color").Index), 2, frmAdicionarModificarItems.cboIde_Color)
                Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_EsCli").Index), 2, frmAdicionarModificarItems.cboIde_EsCli)
                
                
                If Not IsNull(DGridLista.Value(DGridLista.Columns("Ide_Destino").Index)) Then
                    Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Destino").Index), 2, frmAdicionarModificarItems.cboIde_Destino)
                Else
                    frmAdicionarModificarItems.cboIde_Destino.ListIndex = -1
                End If
                
                If Not IsNull(DGridLista.Value(DGridLista.Columns("Ide_Po").Index)) Then
                    Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Po").Index), 2, frmAdicionarModificarItems.CboIde_PO)
                Else
                    frmAdicionarModificarItems.CboIde_PO.ListIndex = -1
                End If
                
                frmAdicionarModificarItems.Frame3.Enabled = False
                frmAdicionarModificarItems.Show 1
                
                Set frmAdicionarModificarItems = Nothing
             
            
                        
        Else
                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
        End If
    
       
    Case "ELIMINAR"
    
      If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
               
                If MsgBox("Esta seguro de eliminar el registro", vbInformation + vbYesNo, "AVISO") = vbYes Then
                    EliminarItem
                End If
        Else
                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
        End If
     
    Case "ELIMDETEMP"
      If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
               
                If MsgBox("Esta seguro de eliminar el item " & DGridLista.Value(DGridLista.Columns("COD_ITEM").Index) & " de esta tempodada", vbInformation + vbYesNo, "AVISO") = vbYes Then
                    EliminarItemTemporada
                End If
        Else
                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
        End If
    Case "SALIR"
    
        Me.Caption = item
        Me.Tag = item
        Unload Me
    
 
    
    'antes
        Case "TEMPORADA"
            If txtcliente <> "" And txttemporada <> "" Then
                Load frmAdItemTemCli
                frmAdItemTemCli.sCod_Cliente = DevuelveCampo("SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE = '" & txtcliente & "'", cCONNECT)
                Set frmAdItemTemCli.oParent = Me
                frmAdItemTemCli.sCod_Temcli = txttemporada
                frmAdItemTemCli.Show vbModal
                Set frmAdItemTemCli = Nothing
            Else
                MsgBox "Debe ingresar Cliente Temporada", vbExclamation, "Aviso"
            End If

        Case "IMPRESION"
            If optcliente.Value Then
                Me.fraImprimir.Visible = True
            Else
                MsgBox "Debe ingresar Cliente Temporada", vbExclamation, "Aviso"
            End If
        Case "COMBINACIONES"
            If cboIde_Color.Text = "S" Or CboIde_PO.Text = "S" Then Exit Sub
            
                If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
                    Load frmMantItemComb
                    frmMantItemComb.Caption = "COMBINACIONES DE ITEM:" & DGridLista.Value(DGridLista.Columns("Cod_Item").Index) & " " & DGridLista.Value(DGridLista.Columns("Des_Item").Index)
                    frmMantItemComb.Codigo_item = DGridLista.Value(DGridLista.Columns("Cod_Item").Index)
                    frmMantItemComb.txtdes_item = DGridLista.Value(DGridLista.Columns("Des_Item").Index)
                    frmMantItemComb.FunctDetalles.Visible = False
                    frmMantItemComb.CARGA_GRID
                    frmMantItemComb.Show 1
                Else
                    MsgBox ("Debe seleccionar un Item para acceder a esta opcion")
                End If
        Case "PROVEEDOR"
            If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
                Load frmManItemProvShort
               frmManItemProvShort.sUniMedDefault = DGridLista.Value(DGridLista.Columns("cod_unimed").Index)

                frmManItemProvShort.varCod_item = DGridLista.Value(DGridLista.Columns("cod_item").Index)
                frmManItemProvShort.varCod_Proveedor = DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index)
                frmManItemProvShort.Caption = "Item Proveedor  Item :" & DGridLista.Value(DGridLista.Columns("cod_item").Index)
                frmManItemProvShort.CARGA_GRID
                frmManItemProvShort.Show 1
                Set frmManItemProvShort = Nothing
                CargaLista
            Else
                MsgBox ("Debe seleccionar un Item para acceder a esta opcion")
            End If

        Case "MEDIDA"
            If cboIde_Talla.Text = "S" Then Exit Sub
            
            FrmMantMed.Cod_Item = txtcoditem
            FrmMantMed.Tipo_Item = "I"
            FrmMantMed.Show 1
        Case "VERESTILOSNP"
           If DGridLista.RowCount > 0 Then
              frmEstilosNPs.codCliente = DGridLista.Value(DGridLista.Columns("cod_cliente").Index)
              frmEstilosNPs.codTemporada = DGridLista.Value(DGridLista.Columns("cod_temcli").Index)
              frmEstilosNPs.codItem = DGridLista.Value(DGridLista.Columns("cod_item").Index)
              frmEstilosNPs.Show 1
           End If
    End Select
Exit Sub
AceptarErr:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "IMPRIMIR"
            If optPendientes Then
                ReporteControl "P"
            Else
                ReporteControl "T"
            End If
        
    Case "CANCELAR"
        Me.fraImprimir.Visible = False

End Select
End Sub

Sub Plin(ByVal Text)
If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
End Sub


Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Eliminar As Integer
Dim StrSQL As String

    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            FraBuscar.Enabled = False
            LIMPIAR_DATOS
            HABILITA_DATOS
            txtcoditem.Enabled = False
            txtDesItem.SetFocus
            
            DGridLista.Enabled = False
            varCod_item = ""
        Case "MODIFICAR"
            sTipo = "U"
            FraBuscar.Enabled = False
            HABILITA_DATOS
            txtcoditem.Enabled = False
            cboCod_FamItem.Enabled = False
            txtDesItem.SetFocus
            
            DGridLista.Enabled = False
            varCod_item = DGridLista.Value(DGridLista.Columns("Cod_item").Index)
                
            StrSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
            If Trim(cboCod_FamItem.Text) <> "" Then
            If DevuelveCampo(StrSQL, cCONNECT) = "M" Then
                sConta = DevuelveCampo("select count(*) from LG_Autorizacion_Campos where cod_usuario='" & vusu & "' and Tipo_Autorizacion ='1'", cCONNECT)
                If sConta > 0 Then
                    HABILITA_CARACMXT True
                End If
                'CboTipCar.Enabled = False
            Else
                HABILITA_CARACMXT False
            End If
    End If

        Case "ELIMINAR"
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?.", vbInformation + vbYesNo, "Items")
            If Eliminar = vbYes Then
                sTipo = "D"
                Call ELIMINAR_DATOS
                Call CargaLista
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                
                SALVAR_DATOS
                CargaDatos
                
                DGridLista.Enabled = True
                'fraoptions.Enabled = False
                FraBuscar.Enabled = True
                If optItem.Value Then
                    txtcod_item.Text = varCod_item
                End If
                Call CargaLista
                sTipo = ""
                

            End If
        Case "DESHACER"
            INHABILITA_DATOS
            sTipo = ""
            LIMPIAR_DATOS
            CargaDatos
            
            DGridLista.Enabled = True
            'fraoptions.Enabled = False
            FraBuscar.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
Case "ACEPT"

   If MsgBox("Esta seguro de cambiar de estado", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
        Dim sEstado As String
        Dim StrSQL As String
        If cboEsta.Text = "Aprobado" Then
        sEstado = "A"
        End If
        
        If cboEsta.Text = "Aprobado Parcial" Then
        sEstado = "B"
        End If
      StrSQL = " exec ES_Cambia_Status_Ubicacion '" & Trim(DGridLista.Value(DGridLista.Columns("Cod_Itemx").Index)) & "','" & sEstado & "' "
      Call ExecuteSQL(cCONNECT, StrSQL)
      DGridLista.Enabled = True
      Call CargaLista
      Frame4.Visible = False
      End If
      
Case "CANCE"
    DGridLista.Enabled = True
    Frame4.Visible = False
End Select

End Sub

Private Sub optcliente_Click()
                    
    'txtcliente.Text = ""
    'txtNom_Cliente.Text = ""
    'txttemporada.Text = ""
    'txtNom_TemCli.Text = ""

    Opcion = 1
    
    'HabilitaMant Me.FunctButt1, "TEMPORADA/IMPRIMIR/LISTADO"
    'Call CargaLista
    If Me.Visible = True Then
    txtcliente.SetFocus
    End If
End Sub

Private Sub optfamitem_Click()
    
    
    
    Opcion = 1
       
    'HabilitaMant Me.FunctButt1, "TEMPORADA/LISTADO"
    
    Call CargaLista

End Sub

 

Private Sub OptEstado_Click()
   Opcion = 3
      Call CargaLista
       txtCodStatus.SetFocus
End Sub

Private Sub Option1_Click()

End Sub

Private Sub optitem_Click()
    'txtcod_item.Text = ""
    'txtdes_item.Text = ""
    Opcion = 2
    'HabilitaMant Me.FunctButt1, "TEMPORADA/LISTADO"
    'Call CargaLista
    txtcod_item.SetFocus
End Sub






 

 
 
Private Sub txtcliente_Change()
    optcliente.Value = True
End Sub

Private Sub txtcod_item_Change()
    optItem.Value = True
End Sub

Private Sub txtCodStatus_Change()
    OptEstado.Value = True
End Sub

Private Sub txtdes_item_Change()
    optItem.Value = True
End Sub

Private Sub txtDesStatus_Change()
    OptEstado.Value = True
End Sub

 Private Sub txtDesStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaStatus 2
         
End If
End Sub

Private Sub txtCodStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        BuscaStatus 1
          
End If
End Sub
Private Sub BuscaStatus(Opcion As Integer)
   Dim sField As String, iRows As Long
   Dim rstAux As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select flg_status, des_status From LG_Status_Servicios WHERE "
    txtCodStatus = Trim(txtCodStatus)
    txtDesStatus = Trim(txtDesStatus)
    sField = txtCodStatus
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "flg_status like '%" & txtCodStatus & "%'"
    Case 2: StrSQL = StrSQL & "des_status like '%" & txtDesStatus & "%'"
    End Select
    
    txtCodStatus = ""
    txtDesStatus = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        
        Codigo = ".."
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodStatus = rstAux!flg_status
            txtDesStatus = rstAux!des_status
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
                'txtCodOrigen.Enabled = False
                'txtDesOrigen.Enabled = False
            End If
            FunctBuscar.SetFocus
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
   

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            txttemporada.SetFocus
        End If
    End If
End Sub


Private Sub txtdes_item_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Trim(txtdes_item.Text) = "" Then
             'Call MsgBox("Sirvase ingresar una Descripcion del Item", vbInformation)
             Call MUESTRA_ITEMS(3)
        Else
            'Esta consulta es para obtener el Codigo de Cliente
            Call MUESTRA_ITEMS(2)
            

        End If
        Call CargaLista
    End If
End Sub



Private Sub txtNom_Cliente_Change()
    optcliente.Value = True
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtcliente.Text = DevuelveCampo(StrSQL, cCONNECT)
        Else
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        End If
    End If
End Sub

Private Sub txtNom_TemCli_Change()
    optcliente.Value = True
End Sub

Private Sub txtNom_TemCli_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    'Esta consulta es para obtener el Codigo de Cliente
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    If KeyAscii = 13 Then
        If Len(txtNom_TemCli) > 4 Then
            'Esta consulta nos permite obtener el Matching entre Cliente y Temporada
            StrSQL = "SELECT Cod_TemCli FROM TG_TEMCLI WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "' AND Nom_TemCli LIKE '" & Trim(txtNom_TemCli.Text) & "%'"
            txttemporada.Text = DevuelveCampo(StrSQL, cCONNECT)
        Else
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        End If
    End If
End Sub


 Private Sub txtCodProveedor2_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
        If Trim(txtCodProveedor2.Text) = "" Then
            Call Me.BUSCA_PROVEEDOR(3)
        Else
            txtCodProveedor.Text = Right("0000000000000" & Trim(txtCodProveedor2.Text), 12)
            Call Me.BUSCA_PROVEEDOR(1)
        End If
    End If
 
End Sub


Private Sub txtNombreProveedor2_KeyPress(KeyAscii As Integer)
 
  If KeyAscii = 13 Then
        Call Me.BUSCA_PROVEEDOR(2)
  End If
 
End Sub


Public Sub BUSCA_PROVEEDOR(tipo As Integer)
Dim StrSQL As String
    Select Case tipo
        Case 1:
                    'Strsql = "SELECT Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Me.txtCod_Proveedor.Text) & "' AND Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                    StrSQL = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor2.Text & "','" & Me.txtNombreProveedor2.Text & "'"
                    txtNombreProveedor2.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                    'txtCod_TemCli.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        'oTipo.sQuery = "SELECT Cod_Proveedor AS  'Código', Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Des_Proveedor LIKE  '%" & Trim(Me.txtDes_Proveedor.Text) & "%' AND Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                        oTipo.sQuery = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor2.Text & "','" & Me.txtNombreProveedor2.Text & "'"
                    Else
                        'oTipo.sQuery = "SELECT Cod_Proveedor AS  'Código', Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                        oTipo.sQuery = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor2.Text & "','" & Me.txtNombreProveedor2.Text & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                        txtCodProveedor2.Text = Trim(Codigo)
                        txtNombreProveedor2.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    FunctButt1.SetFocus
End Sub



Private Sub txttemporada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     cmdBusTemporada_Click
     txtCodFamilia.SetFocus
    End If
End Sub

Private Sub txtCodFamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctBuscar.SetFocus
    End If
End Sub

Private Sub txtcod_item_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Trim(txtcod_item.Text) = "" Then
            Call MUESTRA_ITEMS(1)
            'Call MsgBox("Sirvase ingresar un codigo de Item", vbInformation)
        Else
            txtcod_item.Text = CompletaCodigo(Trim(txtcod_item.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            StrSQL = "SELECT Des_Item FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
            txtdes_item.Text = DevuelveCampo(StrSQL, cCONNECT)
        End If
        FunctBuscar.SetFocus
        
    End If
End Sub

Public Sub CargaLista()
    Dim StrSQL As String
    Dim xRow As Variant
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
     
    StrSQL = "EXEC ES_SM_ItemServicios_ClienteTemp '" & Opcion & "','" & DevuelveCampo(StrSQL, cCONNECT) & "','" & txttemporada.Text & "','" & txtcod_item.Text & "','" & txtCodStatus.Text & "', '" & txtCodProveedor2.Text & "','" & txtCodFamilia.Text & "' "
    
    xRow = DGridLista.Row
    Set DGridLista.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)

    DGridLista.Row = xRow
    DGridLista.Enabled = True
    SeteaGrid
End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    Contador = Len(CodOrigen) - PosfinalCod
    If Contador < 0 Then
        Contador = 0
    End If
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Contador)
End Function


Public Sub CargaCombos()
    Dim StrSQL As String
    
    'Combo Familia Item
    StrSQL = "SELECT des_famitem + space(100) + cod_famitem  FROM LG_FamIte"
    Call LlenaCombo(cboCod_FamItem, StrSQL, cCONNECT)
    
    'Combo Grupo Item
    'Strsql = "SELECT  Cod_Gruitem as Código, des_famgruite as Descripción FROM LG_FamGruIte WHERE Cod_Famitem='" & cboCod_FamItem.Text & "'"
    'Call LlenaCombo(cboCod_GruItem, Strsql, cCONNECT)
    
    'Combo Unidad de Medida
    StrSQL = "SELECT Des_UniMed + space(100) + Cod_UniMed  FROM TG_UniMed"
    Call LlenaCombo(cboCod_UniMed, StrSQL, cCONNECT)
    
    'Combo Clase de Item
    StrSQL = "SELECT des_claitem + space(100) + cod_claitem  FROM LG_Claitem"
    Call LlenaCombo(cboCod_ClaItem, StrSQL, cCONNECT)
    
    'Combo Flag Estatus
    StrSQL = "SELECT Des_Status + space(100) + Flg_Status  FROM LG_Status_Servicios"
    Call LlenaCombo(cboFlg_Status, StrSQL, cCONNECT)
    
    'Combo Origen
    StrSQL = "SELECT des_origen + space(100) + cod_origen  FROM LG_Origen"
    Call LlenaCombo(cboCod_Origen, StrSQL, cCONNECT)
    
    'Combo Motivo Preproduccion
    StrSQL = "SELECT des_motprepro + space(100) + cod_motprepro  FROM TG_MotPrePro"
    Call LlenaCombo(cboCod_MotPrePro, StrSQL, cCONNECT)
    
    StrSQL = " SELECT  Des_ModoProceso + space(100) +   Flg_ModoProceso  FROM ES_ModoProceso "
    Call LlenaCombo(cboModoProceso, StrSQL, cCONNECT)
    
'    Strsql = "SELECT nom_cliente + space(100) + Cod_Cliente FROM TG_Cliente"
'    Call LlenaCombo(cboCod_Cliente, Strsql, cCONNECT)
        
        
    'Combo Identificador Talla
    cboIde_Talla.Clear
    cboIde_Talla.AddItem ("N")
    cboIde_Talla.AddItem ("S")
    cboIde_Talla.ListIndex = 0
    'Combo Identificador Color
    cboIde_Color.Clear
    cboIde_Color.AddItem ("N")
    cboIde_Color.AddItem ("S")
    cboIde_Color.ListIndex = 0
    'Combo Identificador Estilo Cliente
    cboIde_EsCli.Clear
    cboIde_EsCli.AddItem ("N")
    cboIde_EsCli.AddItem ("S")
    cboIde_EsCli.ListIndex = 0
    'Combo Identificador de Destino
    cboIde_Destino.Clear
    cboIde_Destino.AddItem ("N")
    cboIde_Destino.AddItem ("S")
    cboIde_Destino.ListIndex = 0
    'Combo Identificador de p.o.
    CboIde_PO.Clear
    CboIde_PO.AddItem ("N")
    CboIde_PO.AddItem ("S")
    CboIde_PO.ListIndex = 0
    
    
    
    cboEsta.Clear
    cboEsta.AddItem ("Aprobado")
    cboEsta.AddItem ("Aprobado Parcial")
    cboEsta.ListIndex = 0
    
End Sub

Public Sub CargaDatos()
Dim StrSQL As String
    If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
    
        txtcoditem.Text = Trim(DGridLista.Value(DGridLista.Columns("Cod_Itemx").Index))
        txtDesItem.Text = Trim(DGridLista.Value(DGridLista.Columns("Des_Item").Index))
       
        
        If IsNull(DGridLista.Value(DGridLista.Columns("Comentario").Index)) Then
            txtComentario.Text = ""
        Else
            txtComentario.Text = Trim(DGridLista.Value(DGridLista.Columns("Comentario").Index))
        End If
        
        
        If IsNull(DGridLista.Value(DGridLista.Columns("Ubicacion").Index)) Then
            txtUbicacion.Text = ""
        Else
            txtUbicacion.Text = Trim(DGridLista.Value(DGridLista.Columns("Ubicacion").Index))
        End If
        
        
        
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_FamItem").Index), 2, cboCod_FamItem)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_GruItem").Index), 2, cboCod_GruItem)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_UniMed").Index), 2, cboCod_UniMed)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_ClaItem").Index), 2, cboCod_ClaItem)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Flg_Status_Ubicacion").Index), 2, cboFlg_Status)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_Origen").Index), 2, cboCod_Origen)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Cod_MotPrePro").Index), 2, cboCod_MotPrePro)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Talla").Index), 2, cboIde_Talla)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Color").Index), 2, cboIde_Color)
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_EsCli").Index), 2, cboIde_EsCli)
        
        
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Flg_ModoProceso").Index), 2, cboModoProceso)
        
        If DGridLista.Value(DGridLista.Columns("Fec_Ult_Aprob_Ubicacion").Index) = "" Or IsNull(DGridLista.Value(DGridLista.Columns("Fec_Ult_Aprob_Ubicacion").Index)) Then
        dtpFechaUbicacion.Value = Null
        Else
        dtpFechaUbicacion.Value = DGridLista.Value(DGridLista.Columns("Fec_Ult_Aprob_Ubicacion").Index)
        End If
        
        
        Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Flg_Status_Ubicacion").Index), 2, cboIde_EsCli)
        
        
        
        
        
        If Not IsNull(DGridLista.Value(DGridLista.Columns("Ide_Destino").Index)) Then
            Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Destino").Index), 2, cboIde_Destino)
        Else
            cboIde_Destino.ListIndex = -1
        End If
        
        If Not IsNull(DGridLista.Value(DGridLista.Columns("Ide_Po").Index)) Then
            Call BuscaCombo(DGridLista.Value(DGridLista.Columns("Ide_Po").Index), 2, CboIde_PO)
        Else
            CboIde_PO.ListIndex = -1
        End If
        
        
        
        txtCodProveedor.Text = Trim(DGridLista.Value(DGridLista.Columns("proveedor").Index))
        txtNombreProveedor.Text = Trim(DGridLista.Value(DGridLista.Columns("DES_proveedor").Index))
        txtPrecio.Text = Trim(DGridLista.Value(DGridLista.Columns("Pre_Cotizado_Proveedor").Index))
        txtObservaciones_Proveedor = Trim(DGridLista.Value(DGridLista.Columns("Observaciones_Proveedor").Index))
        
        txtCodItemPro.Text = Trim(DGridLista.Value(DGridLista.Columns("cod_itemProv").Index))
        txtUMPro.Text = Trim(DGridLista.Value(DGridLista.Columns("cod_unimedprov").Index))
                       
        Open_Imagen (Trim(DGridLista.Value(DGridLista.Columns("Dir_Icono").Index)))
                
                
        Me.txtPrecioComercial = Trim(DGridLista.Value(DGridLista.Columns("Precio_Cotizacion_Artes").Index))
        Me.txtTecnicaEstampado = Trim(DGridLista.Value(DGridLista.Columns("Tecnica_Estampado").Index))
        
        
        
    End If
End Sub

Public Sub LIMPIAR_DATOS()

    txtcoditem.Text = ""
    txtDesItem.Text = ""
  '  txtCan_LotPed.Text = "0"
    txtComentario.Text = ""
    
    'Limpiamos el Grupo
    cboCod_GruItem.Clear  '.ListIndex = -1
    
    
    cboCod_FamItem_Click
    
    
    Call BuscaCombo("L ", 2, cboCod_Origen)
    Call BuscaCombo("P", 2, cboFlg_Status)
    Call BuscaCombo("P ", 2, cboCod_ClaItem)
    
    cboCod_UniMed.ListIndex = -1
    'cboCod_ClaItem.ListIndex = -1
    'cboFlg_Status.ListIndex = -1
    'cboCod_Origen.ListIndex = -1
    cboCod_MotPrePro.ListIndex = -1
    cboIde_Talla.ListIndex = 0
    cboIde_Color.ListIndex = 0
    cboIde_EsCli.ListIndex = 0
    cboIde_Destino.ListIndex = 0
    CboIde_PO.ListIndex = 0
    

    'sTipo = ""
End Sub

Public Sub HABILITA_DATOS()

    txtcoditem.Enabled = True
    txtDesItem.Enabled = True

    txtComentario.Enabled = True
    
    
    If Opcion = 1 Then
        cboCod_FamItem.Enabled = False
    Else
        cboCod_FamItem.Enabled = True
    End If
    cboCod_GruItem.Enabled = True
    cboCod_UniMed.Enabled = True
    cboCod_ClaItem.Enabled = True
    cboFlg_Status.Enabled = True
    cboCod_Origen.Enabled = True
    cboCod_MotPrePro.Enabled = True
    cboIde_Talla.Enabled = True
    cboIde_Color.Enabled = True
    cboIde_EsCli.Enabled = True
    cboIde_Destino.Enabled = True
    CboIde_PO.Enabled = True

End Sub

Public Sub INHABILITA_DATOS()

    txtcoditem.Enabled = False
    txtDesItem.Enabled = False
    'txtCan_LotPed.Enabled = False
    txtComentario.Enabled = False
    txtUbicacion.Enabled = False
    
    cboCod_FamItem.Enabled = False
    cboCod_GruItem.Enabled = False
    cboCod_UniMed.Enabled = False
    cboCod_ClaItem.Enabled = False
    cboFlg_Status.Enabled = False
    cboCod_Origen.Enabled = False
    cboCod_MotPrePro.Enabled = False
    cboIde_Talla.Enabled = False
    cboIde_Color.Enabled = False
    cboIde_EsCli.Enabled = False
    cboIde_Destino.Enabled = False
    CboIde_PO.Enabled = False
    
    dtpFechaUbicacion.Enabled = False
    cboModoProceso.Enabled = False
    
    Me.txtPrecioComercial.Enabled = False
    Me.txtTecnicaEstampado.Enabled = False
    
    Me.txtCodProveedor.Enabled = False
    Me.txtNombreProveedor.Enabled = False
    Me.txtCodItemPro.Enabled = False
    Me.txtUMPro.Enabled = False
    Me.txtPrecio.Enabled = False
    Me.txtObservaciones_Proveedor.Enabled = False
    
    HABILITA_CARACMXT False
    
End Sub

Public Function VALIDA_DATOS() As Boolean
Dim StrSQL As String
    VALIDA_DATOS = True
    If sTipo = "I" Then
        If Trim(cboCod_FamItem.Text) = "" Then
            Call MsgBox("Sirvase seleccionar una familia", vbCritical)
            cboCod_FamItem.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    If sTipo <> "D" Then 'Es decir es "I" o "U"
        If Trim(txtDesItem.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbCritical)
            txtDesItem.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(cboCod_UniMed.Text) = "" Then
            Call MsgBox("La Unidad de Medida no puede estar vacia. Sirvase verificar", vbCritical)
            cboCod_UniMed.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(cboCod_ClaItem.Text) = "" Then
            Call MsgBox("La Clase de Item no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboCod_ClaItem.SetFocus
            Exit Function
        End If
        If Trim(cboFlg_Status.Text) = "" Then
            Call MsgBox("El Status no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboFlg_Status.SetFocus
            Exit Function
        End If
        If Trim(cboCod_Origen.Text) = "" Then
            Call MsgBox("El código de origen no puede estar vacio. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboCod_Origen.SetFocus
            Exit Function
        End If
        If Trim(cboCod_MotPrePro.Text) = "" Then
            Call MsgBox("El Motivo de Pre Producción no puede estar vacio. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
        StrSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) = "M" Then
            
            'Call HABILITA_CARACMXT(True)
        End If
    End If
        
    
End Function

Public Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Rs.ActiveConnection = cCONNECT
    Rs.CursorLocation = adUseClient
    Rs.CursorType = adOpenStatic
    
    Con.BeginTrans
       
        'Esta sentecia es para obtener el Codigo de Cliente
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
       
        'Esta es la sentencia que realizara el salvado de datos
        StrSQL = "UP_MAN_ITEMS " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        txtcoditem.Text & "','" & _
        Right(cboCod_FamItem.Text, 2) & "','" & _
        Right(cboCod_GruItem.Text, 4) & "','" & _
        Right(cboCod_UniMed.Text, 2) & "','" & _
        txtDesItem.Text & "','" & _
        Right(cboCod_ClaItem.Text, 2) & "'," & _
        "" & "," & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Right(cboCod_Origen.Text, 2) & "','" & _
        cboIde_Talla.Text & "','" & _
        cboIde_Color.Text & "','" & _
        cboIde_EsCli.Text & "','" & _
        cboIde_Destino.Text & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        DevuelveCampo(StrSQL, cCONNECT) & "','" & _
        txttemporada.Text & "','" & Trim(txtComentario.Text) & "','" & _
        CboIde_PO.Text & "','" & vusu & "'"
          
        If sTipo = "I" Then
            Set Rs = Con.Execute(StrSQL)
            If Rs.RecordCount Then
                varCod_item = Rs(0)
            End If
            Set Rs = Nothing
        Else
            Con.Execute StrSQL
        End If
        
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call INHABILITA_DATOS
    Call LIMPIAR_DATOS
    Call HABILITA_CARACMXT(False)
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    Dim StrSQL As String
    
    StrSQL = "SELECT COD_CLIENTE FROM LG_ITEMTEMCLI WHERE Cod_Item='" & txtcoditem.Text & "'"

    If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
           
        'Esta sentecia es para obtener el Codigo de Cliente
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    
        'Esta es la sentencia que realiza la eliminacion del Registro
        StrSQL = "UP_MAN_ITEMS " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        txtcoditem.Text & "','" & _
        Right(cboCod_FamItem.Text, 2) & "','" & _
        Right(cboCod_GruItem.Text, 4) & "','" & _
        Right(cboCod_UniMed.Text, 2) & "','" & _
        txtDesItem.Text & "','" & _
        Right(cboCod_ClaItem.Text, 2) & "'," & _
        "" & "," & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Right(cboCod_Origen.Text, 2) & "','" & _
        cboIde_Talla.Text & "','" & _
        cboIde_Color.Text & "','" & _
        cboIde_EsCli.Text & "','" & _
        cboIde_Destino.Text & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        DevuelveCampo(StrSQL, cCONNECT) & "','" & _
        txttemporada.Text & "','" & Trim(txtComentario.Text) & "','" & _
        CboIde_PO.Text & "','" & vusu & "'"
                
        Con.Execute StrSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIAR_DATOS
    'RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Private Sub HABILITA_CARACMXT(vEstado As Boolean)
End Sub

Sub MUESTRA_ITEMS(tipo As Integer)
    Dim oTipo As New frmBusqGeneral3
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    If tipo = 1 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY cod_Item"
    ElseIf tipo = 2 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item where des_item like '%" & Trim(Me.txtdes_item.Text) & "%' ORDER BY des_Item"
    ElseIf tipo = 3 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY Des_Item"
    End If
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcod_item.Text = Codigo
        txtdes_item.Text = Descripcion
        
        FunctBuscar.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub


'----------------------------------------
Sub Open_Imagen(StrRuta As String)

On Error Resume Next

Set Image1.Picture = LoadPicture(StrRuta)

Exit Sub

Resume

ErrHandler:
  ErrorHandler Err, "Carga Imagen"
End Sub

Private Sub DGridlista_DblClick()
    Dim i As Integer
    For i = 1 To DGridLista.Columns.Count
        Debug.Print DGridLista.Name & ".Columns(" & Chr(34) & DGridLista.Columns(i).Caption & Chr(34) & ").width = " & CStr(DGridLista.Columns(i).Width)
    Next
End Sub


Sub SeteaGrid()


DGridLista.Columns("cod_cliente").Visible = False
DGridLista.Columns("cod_temcli").Visible = False
DGridLista.Columns("cod_item").Visible = False
DGridLista.Columns("abr_cliente").Visible = False
DGridLista.Columns("nom_cliente").Width = 1215
DGridLista.Columns("nom_cliente").Caption = "Cliente"
DGridLista.Columns("nom_temcli").Width = 1230
DGridLista.Columns("nom_temcli").Caption = "Temporada"
DGridLista.Columns("Cod_Item").Width = 1065
DGridLista.Columns("Cod_Item").Caption = "Item"
DGridLista.Columns("Cod_FamItem").Width = 600
DGridLista.Columns("Cod_FamItem").Caption = "Familia"
DGridLista.Columns("Cod_GruItem").Visible = False
DGridLista.Columns("Cod_UniMed").Visible = False
DGridLista.Columns("Cod_CtaCont").Visible = False
DGridLista.Columns("Des_Item").Width = 5000
DGridLista.Columns("Des_Item").Caption = "Descrip.Item"
DGridLista.Columns("Fec_Creacion").Visible = False
DGridLista.Columns("Cod_ClaItem").Visible = False
DGridLista.Columns("Can_PtoReor").Visible = False
DGridLista.Columns("Can_LotPed").Visible = False
DGridLista.Columns("Rep_PreDol").Visible = False
DGridLista.Columns("Flg_Status").Visible = False
DGridLista.Columns("Cod_Origen").Visible = False

DGridLista.Columns("Ide_Color").Visible = False
DGridLista.Columns("Ide_EsCli").Visible = False
DGridLista.Columns("Cod_MotPrePro").Visible = False


DGridLista.Columns("Cod_Proveedor").Visible = False
DGridLista.Columns("Ser_OrdComp").Visible = False
DGridLista.Columns("Cod_OrdComp").Visible = False
DGridLista.Columns("Pre_UltComp").Visible = False
DGridLista.Columns("Fec_UltComp").Visible = False
DGridLista.Columns("Cod_MonUltComp").Visible = False
DGridLista.Columns("Dir_Icono").Width = 1500
DGridLista.Columns("Dir_Icono").Caption = "Grafico"
DGridLista.Columns("Dat_UltAct").Visible = False
DGridLista.Columns("Ide_Destino").Visible = False
DGridLista.Columns("Comentario").Width = 1500
DGridLista.Columns("Comentario").Caption = "Comentario"
DGridLista.Columns("Por_MerTin").Visible = False
DGridLista.Columns("cod_hiltel").Visible = False
DGridLista.Columns("cod_tipcar").Visible = False
DGridLista.Columns("fac_conversion").Visible = False
DGridLista.Columns("Ide_PO").Visible = False

DGridLista.Columns("Tip_Item_Costeo").Visible = False
DGridLista.Columns("rep_presol").Visible = False
DGridLista.Columns("Codigo_Barras").Visible = False
DGridLista.Columns("Codigo_Cocina").Visible = False
DGridLista.Columns("Rep_PreDol_Importado").Visible = False
DGridLista.Columns("Flg_Quimico_Controlado").Visible = False
DGridLista.Columns("Porc_Concentracion").Visible = False
DGridLista.Columns("Kilos_Quimico_por_Kilo_Tela").Visible = False
DGridLista.Columns("Flg_Quimico_Lavanderia").Visible = False
DGridLista.Columns("Flg_Producto_Caldero").Visible = False
DGridLista.Columns("Flg_Combustible").Visible = False
DGridLista.Columns("Flg_Productos_Quimicos").Visible = False
DGridLista.Columns("Cod_CtaConCIF").Visible = False
DGridLista.Columns("Cod_CtaConDerechosAduaneros").Visible = False
DGridLista.Columns("Cod_CtaConGastosDespacho").Visible = False
DGridLista.Columns("Cod_CtaConAbono").Visible = False
DGridLista.Columns("Cod_Concepto").Visible = False
DGridLista.Columns("Cod_UniMed_Cotizacion").Visible = False
DGridLista.Columns("Ubicacion").Width = 1500
DGridLista.Columns("Ubicacion").Caption = "Ubicacion"
DGridLista.Columns("Fec_Ult_Aprob").Visible = False
DGridLista.Columns("Flg_Status_Ubicacion").Visible = False
DGridLista.Columns("Fec_Ult_Aprob_Ubicacion").Visible = False
DGridLista.Columns("Flg_ModoProceso").Visible = False
DGridLista.Columns("Tip_Version").Visible = False
DGridLista.Columns("proveedor").Visible = False

DGridLista.Columns("Des_Proveedor").Width = 1500
DGridLista.Columns("Des_Proveedor").Caption = "Proveedor"
DGridLista.Columns("des_famitem").Width = 1500
DGridLista.Columns("des_famitem").Caption = "Familia"
DGridLista.Columns("des_famgruite").Visible = False
DGridLista.Columns("Des_UniMed").Width = 1500
DGridLista.Columns("Des_UniMed").Caption = "Uni.Med"
DGridLista.Columns("des_claitem").Width = 1500
DGridLista.Columns("des_claitem").Caption = "Des. Clase Item"
DGridLista.Columns("des_status").Width = 1500
DGridLista.Columns("des_status").Caption = "Des. Status"
DGridLista.Columns("des_origen").Width = 1500
DGridLista.Columns("des_origen").Caption = "Des. origen"
DGridLista.Columns("des_motprepro").Width = 1500
DGridLista.Columns("des_motprepro").Caption = "Des. Motivo"
DGridLista.Columns("des_tipcar").Visible = False
DGridLista.Columns("des_modoproceso").Width = 1500
DGridLista.Columns("des_modoproceso").Caption = "Des. Modo"
DGridLista.Columns("Precio").Visible = False
DGridLista.Columns("Moneda").Visible = False
DGridLista.Columns("Ult_Compra").Width = 1500
DGridLista.Columns("Ult_Compra").Caption = "Ult Compra"
DGridLista.Columns("O_Compra").Width = 1500
DGridLista.Columns("O_Compra").Caption = "Ord. Compra"
DGridLista.Columns("Des_TipVersion").Visible = False
DGridLista.Columns("cod_itemprov").Width = 1500
DGridLista.Columns("cod_itemprov").Caption = "Item Proveedor"
DGridLista.Columns("cod_unimedprov").Visible = False
DGridLista.Columns("Pre_Cotizado_Proveedor").Width = 1500
DGridLista.Columns("Pre_Cotizado_Proveedor").Caption = "Pre. Coti. Proveedor"
DGridLista.Columns("Observaciones_Proveedor").Width = 1500
DGridLista.Columns("Observaciones_Proveedor").Caption = "Obs. Proveedor"

DGridLista.Columns("Ide_Talla").Caption = "Ide Talla"
DGridLista.Columns("Ide_Talla").Width = 800


 
End Sub

Private Sub ReporteControl(sFlg_Opcion As String)

            Dim oo As Object
            Dim sCod_Cliente As String
            
            sCod_Cliente = DevuelveCampo("SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE = '" & txtcliente.Text & "'", cCONNECT)
            
                On Error GoTo AceptarErr
                Set oo = CreateObject("excel.application")
                oo.workbooks.Open vRuta & "\RptServiciosPendientes.xlt"
                oo.Visible = True
                oo.run "Reporte", sCod_Cliente, txttemporada.Text, txtNom_Cliente, txtNom_TemCli, cCONNECT, sFlg_Opcion
                Screen.MousePointer = vbNormal
                oo.Visible = True
                Set oo = Nothing
                Me.fraImprimir.Visible = False
                Exit Sub
            
Exit Sub
AceptarErr:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub

Sub EliminarItem()
On Error GoTo errx

Dim StrSQL As String

StrSQL = "UP_MAN_ITEMS2 " & _
Opcion & ",'D','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_itemx").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_FamItem").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_GruItem").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_UniMed").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Des_Item").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_ClaItem").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_origen").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ide_Talla").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ide_Color").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ide_EsCli").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ide_Destino").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Cod_MotPrePro").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_cliente").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_temcli").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Comentario").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ide_Po").Index)) & "','" & _
vusu & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Ubicacion").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Flg_Status").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Tip_version").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("Flg_ModoProceso").Index)) & "','','" & Trim(DGridLista.Value(DGridLista.Columns("proveedor").Index)) & "','" & Trim(DGridLista.Value(DGridLista.Columns("cod_itemprov").Index)) & "','" & Trim(DGridLista.Value(DGridLista.Columns("cod_unimedprov").Index)) & "','" & Trim(DGridLista.Value(DGridLista.Columns("Pre_cotizado_proveedor").Index)) & "','" & Trim(DGridLista.Value(DGridLista.Columns("Observaciones_proveedor").Index)) & "'"

Call ExecuteSQL(cCONNECT, StrSQL)
Call CargaLista

Exit Sub

errx:
    ErrorHandler Err, "EliminarItem"
End Sub

Sub EliminarItemTemporada()
On Error GoTo errx

Dim StrSQL As String

StrSQL = "LG_ELIMINA_ITEM_TEMPORADA '" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_itemx").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_cliente").Index)) & "','" & _
Trim(DGridLista.Value(DGridLista.Columns("cod_temcli").Index)) & "'"

Call ExecuteSQL(cCONNECT, StrSQL)
Call CargaLista

Exit Sub

errx:
    ErrorHandler Err, "EliminarItem"
End Sub

Public Sub BuscaDesCliente()
 If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            sStrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(sStrSQL, cCONNECT)
     
        End If
End Sub

Public Sub BuscaDesTemporada()
 If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            sStrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
            sStrSQL = "SELECT  Nom_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(sStrSQL, cCONNECT) & "'  AND cod_temcli like '%" & txttemporada.Text & "%'"
            txtNom_TemCli.Text = DevuelveCampo(sStrSQL, cCONNECT)
           
        End If
End Sub

Public Sub FindItem(sCod_Item As String)
On Error GoTo errx
Dim bFind As Boolean

bFind = DGridLista.Find(DGridLista.Columns("cod_item").Index, jgexGreaterThanOrEqualTo, sCod_Item)
CargaDatos

Exit Sub

CargaDatos
errx:
    ErrorHandler Err, "FindItem"
End Sub
