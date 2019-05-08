VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmShowCtaCte 
   Caption         =   "Consulta Cuenta Corriente Clientes"
   ClientHeight    =   8595
   ClientLeft      =   480
   ClientTop       =   480
   ClientWidth     =   16335
   Icon            =   "frmShowCtaCte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   16335
   Begin VB.Frame FraFecEsp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   2520
      TabIndex        =   71
      Top             =   5400
      Visible         =   0   'False
      Width           =   6795
      Begin MSComCtl2.DTPicker dtpFechaT 
         Height          =   300
         Left            =   4920
         TabIndex        =   77
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   40579
      End
      Begin VB.TextBox txtOrigenT 
         Height          =   300
         Left            =   900
         TabIndex        =   75
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox txtDescripcionT 
         Height          =   300
         Left            =   1245
         TabIndex        =   74
         Top             =   360
         Width           =   2805
      End
      Begin VB.CommandButton cmdFecEsp 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4110
         TabIndex        =   73
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanFecEsp 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5340
         TabIndex        =   72
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha :"
         Height          =   375
         Left            =   4200
         TabIndex        =   78
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   195
         Left            =   240
         TabIndex        =   76
         Top             =   413
         Width           =   465
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         Height          =   1515
         Left            =   120
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame fanANO_MES 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1493
      TabIndex        =   49
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
      Begin VB.CommandButton cmdCancelar_AnoMes 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7650
         TabIndex        =   55
         Top             =   195
         Width           =   1005
      End
      Begin VB.CommandButton cmdAceptar_AnoMes 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6690
         TabIndex        =   54
         Top             =   195
         Width           =   945
      End
      Begin VB.TextBox txtDesCV 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3630
         TabIndex        =   53
         Top             =   210
         Width           =   2895
      End
      Begin VB.TextBox txtCodCV 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3180
         TabIndex        =   52
         Top             =   210
         Width           =   435
      End
      Begin VB.OptionButton optOPCION_AnoMes 
         Caption         =   "Condicion de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1260
         TabIndex        =   51
         Top             =   255
         Width           =   1905
      End
      Begin VB.OptionButton optOPCION_AnoMes 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   255
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   645
         Left            =   30
         Top             =   30
         Width           =   8775
      End
   End
   Begin VB.Frame fanComisionista 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1493
      TabIndex        =   56
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
      Begin VB.TextBox txtCodComisionista 
         Height          =   300
         Left            =   1500
         TabIndex        =   60
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtDesComisionista 
         Height          =   300
         Left            =   1950
         TabIndex        =   59
         Top             =   210
         Width           =   4545
      End
      Begin VB.CommandButton cmdAceptar_Com 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6690
         TabIndex        =   58
         Top             =   195
         Width           =   945
      End
      Begin VB.CommandButton cmdCancelar_Com 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7650
         TabIndex        =   57
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comisionista"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Tag             =   "Document Type"
         Top             =   270
         Width           =   1245
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   645
         Left            =   30
         Top             =   30
         Width           =   8775
      End
   End
   Begin VB.Frame famClienteComercial 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1493
      TabIndex        =   62
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   2010
         TabIndex        =   70
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox txtDes_Cliente 
         Height          =   285
         Left            =   2895
         TabIndex        =   69
         Top             =   225
         Width           =   3105
      End
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   2655
         TabIndex        =   68
         Tag             =   "..."
         Top             =   225
         Width           =   300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7650
         TabIndex        =   64
         Top             =   195
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6690
         TabIndex        =   63
         Top             =   195
         Width           =   945
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   645
         Left            =   30
         Top             =   30
         Width           =   8775
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Comercial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   65
         Tag             =   "Document Type"
         Top             =   270
         Width           =   1710
      End
   End
   Begin VB.TextBox TxtDDolares 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   885
      Width           =   1455
   End
   Begin VB.Frame fra_origen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   1320
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame Frame2 
         Caption         =   "Ordenado"
         Height          =   615
         Left            =   180
         TabIndex        =   43
         Top             =   840
         Width           =   8775
         Begin VB.OptionButton optCLienteComercial 
            Caption         =   "Por &Cliente Comercial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6570
            TabIndex        =   67
            Top             =   270
            Width           =   2115
         End
         Begin VB.OptionButton optComisionista 
            Caption         =   "Por &Comisionista"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4695
            TabIndex        =   66
            Top             =   270
            Width           =   1725
         End
         Begin VB.OptionButton optAñoMesCliente 
            Caption         =   "&Año-Mes Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2775
            TabIndex        =   48
            Top             =   270
            Width           =   1725
         End
         Begin VB.OptionButton optCliente 
            Caption         =   "&Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   47
            Top             =   270
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton optDocumento 
            Caption         =   "&Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1275
            TabIndex        =   46
            Top             =   270
            Width           =   1305
         End
      End
      Begin VB.TextBox Txt_Destipo 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   6030
         MaxLength       =   30
         TabIndex        =   37
         Top             =   360
         Width           =   2910
      End
      Begin VB.TextBox Txt_Tipo 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   5670
         MaxLength       =   2
         TabIndex        =   36
         Text            =   "FA"
         Top             =   360
         Width           =   360
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7860
         TabIndex        =   28
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6630
         TabIndex        =   27
         Top             =   1530
         Width           =   1215
      End
      Begin VB.TextBox Txt_Descripcion 
         Height          =   300
         Left            =   1125
         TabIndex        =   26
         Top             =   360
         Width           =   2805
      End
      Begin VB.TextBox Txt_Origen 
         Height          =   300
         Left            =   780
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1965
         Left            =   600
         Top             =   360
         Width           =   9105
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Left            =   4410
         TabIndex        =   38
         Tag             =   "Document Type"
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   413
         Width           =   465
      End
   End
   Begin VB.Frame FrmRptLetrasStatus 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   16305
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "LINEA DE CREDITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   12240
         TabIndex        =   86
         Top             =   240
         Width           =   2535
         Begin VB.TextBox Txt_disponible 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1080
            TabIndex        =   92
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Txt_utilizado 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   91
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Txt_otorgado 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   90
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Utilizado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Otorgado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Txt_DesUsuario 
         Height          =   285
         Left            =   4200
         TabIndex        =   85
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox Txt_Cod_Usuario 
         Height          =   285
         Left            =   3360
         TabIndex        =   83
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton Opt_Vendedor 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1920
         Width           =   2370
      End
      Begin VB.CheckBox chkIncluirBoletas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Incluir Boletas"
         Height          =   255
         Left            =   6240
         TabIndex        =   45
         Top             =   1200
         Width           =   1755
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2655
         TabIndex        =   2
         Top             =   360
         Width           =   6825
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2970
         MaxLength       =   11
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   690
         MaxLength       =   11
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.CheckBox chkInLetras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Incluir Letras"
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CheckBox chkLetraTercero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Letras de Tercero"
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox TxtDOtros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox Txt_Importe 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtDsoles 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optDocRef 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documento Específico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   2370
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1485
         Width           =   1440
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6765
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1485
         Width           =   540
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   3735
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1470
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   4095
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1470
         Width           =   1980
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   14880
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton opPendiente 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton oprCanceladas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Canceladas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton opTodas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   6135
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   79626241
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   4080
            TabIndex        =   13
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   79626241
            CurrentDate     =   37543
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rango Fecha de Emisión:"
            Height          =   240
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   2235
         End
      End
      Begin MSComCtl2.DTPicker DFVenci1 
         Height          =   315
         Left            =   3360
         TabIndex        =   80
         Top             =   1920
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   79626241
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker DFVenci2 
         Height          =   315
         Left            =   5460
         TabIndex        =   81
         Top             =   1920
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   79626241
         CurrentDate     =   37543
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   2640
         TabIndex        =   84
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ruc :"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Tag             =   "Anexo Type"
         Top             =   405
         Width           =   435
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Otra Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9720
         TabIndex        =   35
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dolares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9720
         TabIndex        =   31
         Top             =   930
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9720
         TabIndex        =   30
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Deuda Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10860
         TabIndex        =   29
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9720
         TabIndex        =   23
         Top             =   1785
         Width           =   735
      End
      Begin VB.Label lblCod_TipOrdCom 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo Documento:"
         Height          =   390
         Left            =   2745
         TabIndex        =   19
         Tag             =   "Document Type"
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Serie "
         Height          =   195
         Left            =   6150
         TabIndex        =   18
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Número :"
         Height          =   225
         Left            =   7410
         TabIndex        =   17
         Tag             =   "Number"
         Top             =   1560
         Width           =   645
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5220
      Left            =   0
      TabIndex        =   9
      Top             =   2760
      Width           =   16320
      _ExtentX        =   28787
      _ExtentY        =   9208
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
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
      Column(1)       =   "frmShowCtaCte.frx":000C
      Column(2)       =   "frmShowCtaCte.frx":00D4
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowCtaCte.frx":0178
      FormatStyle(2)  =   "frmShowCtaCte.frx":02B0
      FormatStyle(3)  =   "frmShowCtaCte.frx":0360
      FormatStyle(4)  =   "frmShowCtaCte.frx":0414
      FormatStyle(5)  =   "frmShowCtaCte.frx":04EC
      FormatStyle(6)  =   "frmShowCtaCte.frx":05A4
      FormatStyle(7)  =   "frmShowCtaCte.frx":0684
      FormatStyle(8)  =   "frmShowCtaCte.frx":0730
      ImageCount      =   0
      PrinterProperties=   "frmShowCtaCte.frx":07E0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   525
      Left            =   7800
      TabIndex        =   44
      Top             =   8010
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   926
      Custom          =   $"frmShowCtaCte.frx":09B8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   500
      ControlSeparator=   10
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   300
      Top             =   7170
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrEstus As String, strSQL As String
Public codigo, Descripcion As String, TipoAdd As String, strCod_Anxo As String
Dim OP_Opcion As String, sSql As String
Public oGroup As GridEX20.JSGroup
Public oFormat As JSFormatStyle


Private Sub Cmd_Cancelar_Click()
  fra_origen.Visible = False
  optCliente.Value = True
End Sub

Sub Reporte_Masivo()
On Error GoTo ERROR
Dim sSql As String
Dim oo As Object
Dim Ruta As String
Dim Reg1 As ADODB.Recordset, sRutaLogo As String

sSql = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(sSql, cCONNECT)
    

'sSQL = "Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & Txt_Origen & "','" & Txt_Tipo & "','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"

 sSql = "Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & Txt_Origen & "','" & Txt_Tipo & "','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','" & IIf(optCliente, "C", "D") & "','','" & IIf(chkIncluirBoletas, "S", "N") & "'"

'sSQL = "Ventas_Muestra_Documentos_por_Cerrar '" & OP_Opcion & "','" & txtCod_TipAne & "','" & txtCod_Anexo & "','2','N','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & _
'"','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"


Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)

Set Reg1 = GetRecordset1(cCONNECT, sSql)

Configurar

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir Todos") = vbYes Then
    If optCliente Then
      Ruta = vRuta & "\RptCuentasCorrientes_Clientes.xlt"
    Else
      Ruta = vRuta & "\RptCuentasCorrientes_Clientes_Correlativo.xlt"
    End If
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", Reg1, sRutaLogo
    Set oo = Nothing
Else
    If optCliente Then
      Ruta = vRuta & "\RptCuentasCorrientes_Clientes.OTS"
    Else
      Ruta = vRuta & "\RptCuentasCorrientes_Clientes_Correlativo.OTS"
    End If
    
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run sSql, cCONNECT, sRutaLogo
    Set oo = Nothing
End If
Exit Sub
ERROR:
    errores err.Number
End Sub



Private Sub Cmd_Imprimir_Click()
    If optAñoMesCliente.Value = True Then
        fra_origen.Enabled = False
        fanANO_MES.Visible = True
        Exit Sub
    End If
        
    If optComisionista.Value = True Then
        fra_origen.Enabled = False
        fanComisionista.Visible = True
        txtCodComisionista.SetFocus
        Exit Sub
    End If
    
    If optCLienteComercial.Value = True Then
        fra_origen.Enabled = False
        famClienteComercial.Visible = True
        txtAbr_Cliente.SetFocus
        Exit Sub
    End If
    
    
    
    If optComisionista.Value = True Then
        If Txt_Origen <> "E" Then
            optComisionista.Value = False
            optComisionista.Enabled = False
            Exit Sub
        End If
        fra_origen.Enabled = False
        fanComisionista.Visible = True
        txtCodComisionista.SetFocus
        Exit Sub
    End If
    
    Reporte_Masivo
    Call Cmd_Cancelar_Click
End Sub

Sub Reporte()
On Error GoTo ERROR
'Dim ssql As String
Dim oo As Object
Dim Ruta As String, sRutaLogo As String

strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Imprimir reporte usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", GridEX1.ADORecordset, sRutaLogo
    Set oo = Nothing
Else
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes.OTS"
    
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run sSql, cCONNECT, sRutaLogo
    Set oo = Nothing
End If

Exit Sub
ERROR:
    errores err.Number
End Sub

Private Sub cmdAceptar_AnoMes_Click()
On Error GoTo SALTO_ERROR
Dim oRs As New Recordset
Dim sCondicionVTA As String, Ruta As String

    sCondicionVTA = "''"
    If optOPCION_AnoMes(1).Value = True Then sCondicionVTA = "'" & Trim(txtCodCV) & "'" '
    strSQL = "EXEC Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & _
             Txt_Origen & "','" & Txt_Tipo & "','','','" & dtpFecEmiIni & "','" & _
             dtpFecEmiFin & "','','','','" & vusu & "','" & IIf(chkInLetras, "S", "N") & _
             "','A','','" & IIf(chkIncluirBoletas, "S", "N") & "' " & ", " & sCondicionVTA
    
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    If oRs.RecordCount = 0 Then
        MsgBox "No se han encontrado datos para la impresión.....", vbExclamation
        Exit Sub
    End If
    
    Dim oo As Object
    Dim sRutaLogo As String, sTitulo As String
    
    Ruta = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(Ruta, cCONNECT)
    
    If MsgBox("Desea imprimir utilizando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        Set oo = CreateObject("excel.application")
        
        oo.Workbooks.Open vRuta & "\rptCtaCteClientes_AMC.XLT"
        oo.Visible = True
        oo.displayalerts = False
        oo.Run "reporte", sRutaLogo, oRs
    Else
        Ruta = vRuta & "\rptCtaCteClientes_AMC.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run sRutaLogo, strSQL, cCONNECT
    End If
    Set oo = Nothing
Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub cmdAceptar_Com_Click()
On Error GoTo SALTO_ERROR
Dim oRs As New Recordset, oo As Object
Dim sRutaLogo As String, sComisionista As String, Ruta As String
    
    If Trim(txtCodComisionista) = Empty Then
        MsgBox "Indique el comisionista para poder emitir la impresión.....", vbCritical
        Exit Sub
    End If
    
    strSQL = "EXEC cn_ventas_muestra_facturas_exportacion_pendientes_por_comisionista '" & txtCodComisionista & "'"
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    If oRs.RecordCount = 0 Then
        MsgBox "No se han encontrado datos para la impresión.....", vbExclamation
        Exit Sub
    End If
    
    sComisionista = UCase(Trim(txtDesComisionista)) & " [" & txtCodComisionista & "]"
    
    Ruta = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(Ruta, cCONNECT)
    
    If MsgBox("Desea imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir Historico") = vbYes Then
        Set oo = CreateObject("excel.application")
        
        oo.Workbooks.Open vRuta & "\rptFacturasPendientesPorComisionista.XLT"
        oo.Visible = True
        oo.displayalerts = False
        oo.Run "reporte", sRutaLogo, oRs, sComisionista
    Else
        Ruta = vRuta & "\rptFacturasPendientesPorComisionista.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run sRutaLogo, strSQL, sComisionista, cCONNECT
    End If
    Set oo = Nothing
    
Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub cmdCanFecEsp_Click()
    FraFecEsp.Visible = False
End Sub

Private Sub cmdFecEsp_Click()
ReporteFecha
End Sub

Private Sub dtpFecEmiFin_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Final no puede ser menor a la fecha Inicial", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  dtpFecEmiFin.Value = Date
End Sub

Private Sub dtpFecEmiIni_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Inicial no puede ser mayor a la fecha final", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub Form_Load()
  txtCod_TipAne = "C"
  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date
  
  OP_Opcion = "1"
  DFVenci1 = Date
  DFVenci2 = Date
End Sub

Private Sub cmdBuscar_Click()
  Buscar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Buscar()
On Error GoTo hand

If Trim(txtCod_TipAne) = "" Then

  If optDocRef Then
  
        If txtCod_TipDoc = "" Or txtNum_Docum = "" Then
            Aviso "Debe Ingresar un documento Específico", 1
            Exit Sub
        End If
  Else
  
      If strCod_Anxo = "" Or txtCod_TipAne = "" Then
        MsgBox "Ingrese un Cliente", vbInformation, "AVISO"
        Exit Sub
      End If

      If (IsNull(dtpFecEmiIni) Or IsNull(dtpFecEmiIni)) Then
        MsgBox "Ingrese un Rango de Fechas", vbInformation, "AVISO"
        Exit Sub
      End If
    
      If (dtpFecEmiFin - dtpFecEmiIni) > 60 Then
        MsgBox "No puede Ingresar un Rango Mayor a 60 Dias", vbInformation, "AVISO"
        Exit Sub
      End If
  End If
Else

  If (oprCanceladas Or opTodas) And (IsNull(dtpFecEmiIni) Or IsNull(dtpFecEmiIni)) Then
    MsgBox "Ingrese un Rango de Fechas", vbInformation, "AVISO"
    Exit Sub
  
  End If
    
End If

If OP_Opcion = "9" Then
    dtpFecEmiIni = DFVenci1
    dtpFecEmiFin = DFVenci2
End If

sSql = "Ventas_Muestra_Documentos_por_Cerrar '" & OP_Opcion & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','2','N','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & _
"','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','" & IIf(chkLetraTercero, "X", "") & "','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','C','','" & IIf(chkIncluirBoletas, "S", "N") & "','','" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)

If GridEX1.RowCount = 0 Then Exit Sub
Configurar

Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"


End Sub
Sub Configurar()
On Error GoTo hand

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cliente").Index, jgexSortAscending)

GridEX1.BackColorRowGroup = &H80000005

Txt_Importe = Format(GridEX1.Value(GridEX1.Columns("SALDO_TOTAl").Index), "##,##0.00")
TxtDDolares = Format(GridEX1.Value(GridEX1.Columns("SALDO_total_DOLARES").Index), "##,##0.00")
TxtDsoles = Format(GridEX1.Value(GridEX1.Columns("SALDO_total_SOLES").Index), "##,##0.00")
TxtDOtros = Format(GridEX1.Value(GridEX1.Columns("SALDO_total_OTROS").Index), "##,##0.00")

Txt_otorgado = Format(GridEX1.Value(GridEX1.Columns("lineacredito").Index), "##,##0.00")
Txt_utilizado = Format(GridEX1.Value(GridEX1.Columns("SALDO_TOTAl").Index), "##,##0.00")

'If Not IsNull(GridEX1.Value(GridEX1.Columns("lineacredito").Index)) Then
Txt_disponible = Format(Format(GridEX1.Value(GridEX1.Columns("lineacredito").Index), "##,##0.00") - Format(GridEX1.Value(GridEX1.Columns("SALDO_TOTAl").Index), "##,##0.00"), "##,##0.00")
'End If

GridEX1.Columns("Cod_Tipdoc").Caption = "Tipo"
GridEX1.Columns("Cod_Tipdoc").Width = 600

GridEX1.Columns("SALDO_TOTAl").Visible = False
GridEX1.Columns("Ruc").Visible = False


GridEX1.Columns("Cliente").Width = 0
GridEX1.Columns("Num_Corre").Width = 0
GridEX1.Columns("saldo_equivalente").Width = 0
GridEX1.Columns("Anexo_Contable").Width = 0

GridEX1.Columns("SALDO_TOTAl").Visible = False
GridEX1.Columns("SALDO_total_SOLES").Visible = False
GridEX1.Columns("SALDO_total_DOLARES").Visible = False
GridEX1.Columns("SALDO_total_otros").Visible = False

GridEX1.Columns("Imp_Total").Width = 1200
GridEX1.Columns("Imp_Total").Caption = "Imp Total"

GridEX1.Columns("Saldo_Dolares").Width = 1200
GridEX1.Columns("Saldo_Dolares").Caption = "Saldo Dolares"

GridEX1.Columns("Saldo_Soles").Width = 1200
GridEX1.Columns("Saldo_Soles").Caption = "Saldo Soles"

GridEX1.Columns("Saldo_Otros").Width = 1200
GridEX1.Columns("Saldo_Otros").Caption = "Saldo Otra Moneda"

GridEX1.Columns("Importe_Cancelado").Width = 1300
GridEX1.Columns("Importe_Cancelado").Caption = "Imp Cancelado"

GridEX1.Columns("saldo_equivalente").Width = 1300
GridEX1.Columns("saldo_equivalente").Caption = "saldo Equivalente"

GridEX1.Columns("Flg_Status_DrawBack").Visible = False
GridEX1.Columns("Des_Status").Visible = False

GridEX1.Columns("Fec_Emision").Width = 1125
'GridEX1.Columns("Fec_VenDoc").Width = 1080
GridEX1.Columns("Num_Registro").Width = 1155
GridEX1.Columns("Moneda").Width = 720

If txtCod_TipAne = "" Then GridEX1.DefaultGroupMode = jgexDGMCollapsed Else GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.ContinuousScroll = True

Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"
End Sub




Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName

Case Is = "VERDETALLE"
  Call GridEX1_DblClick
Case Is = "IMPTODOS"
  fra_origen.Visible = True
  Txt_Origen = "N"
  Txt_Descripcion = "Nacional"
  Txt_Tipo = "FA"
  Txt_Destipo = "Facturas"
  Txt_Origen.SetFocus
Case Is = "IMPRIMIR"
  Reporte
Case Is = "IMPTFEES"
  FraFecEsp.Visible = True
  txtOrigenT = "N"
  txtDescripcionT = "Nacional"
  dtpFechaT.Value = Date
  txtOrigenT.SetFocus
  
Case Is = "HISTORICOPAGOS"
  ReporteHistorico
End Select

End Sub

Sub ReporteFecha()
 
Dim oo As Object
Dim sRuta_Logo As String
Dim strSQL As String
 
Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\RptCuentasCorrientes_Clientes_Fecha_Especifica.XLT"
oo.Visible = True
oo.displayalerts = False

Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)


strSQL = "exec cn_encuentra_doc_pend_cobro_por_fecha  '" & Str(dtpFechaT.Value) & "','1','','','N','" & txtOrigenT.Text & "','N'"
oo.Visible = True
oo.displayalerts = False
oo.Run "reporte", strSQL, cCONNECT, rutaLogo, Str(dtpFechaT.Value)

Set oo = Nothing

Exit Sub
errReporte:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Sub GridEX1_DblClick()
  If GridEX1.RowCount = 0 Then Exit Sub
  Load frmShowCtaCteDet
  frmShowCtaCteDet.Caption = "Detalle Cliente " & GridEX1.Value(GridEX1.Columns("Cliente").Index) & " Documento : " & GridEX1.Value(GridEX1.Columns("Documento").Index)
  frmShowCtaCteDet.strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & GridEX1.Value(GridEX1.Columns("NUM_CORRE").Index) & "'"
  frmShowCtaCteDet.Buscar
  frmShowCtaCteDet.Show vbModal
End Sub

Private Sub opPendiente_Click()
StrEstus = "P"
OP_Opcion = "1"

End Sub

Private Sub oprCanceladas_Click()
StrEstus = "C"
OP_Opcion = "2"
End Sub

Private Sub optDocumEspecifico_Click()
    StrEstus = "E"
    txtCod_TipDoc.SetFocus
End Sub

Private Sub Opt_Vendedor_Click()
OP_Opcion = "10"
End Sub

Private Sub optDocRef_Click()
    StrEstus = "R"
    OP_Opcion = "4"
    txtCod_TipDoc.SetFocus
End Sub



Private Sub Option1_Click()
OP_Opcion = "9"
End Sub

Private Sub Option2_Click()

End Sub

Private Sub opTodas_Click()
StrEstus = "T"
OP_Opcion = "3"
End Sub

Sub LimpiaFr()
  GridEX1.ClearFields
  txtCod_TipAne = ""
  txtDes_TipAnex = ""
  txtCod_TipAne = ""


  txtNum_Ruc = ""
End Sub


Private Sub optOPCION_AnoMes_Click(Index As Integer)
    txtCodCV.Enabled = False: txtDesCV.Enabled = False
    txtCodCV = Empty: txtDesCV = Empty
    Select Case Index
        Case 0
        Case 1
            txtCodCV.Enabled = True: txtDesCV.Enabled = True
            txtCodCV.SetFocus
    End Select
End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_Descripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.Txt_Descripcion.Text) = "" Then
            Call Me.BUSCA_ORIGEN(3)
            
        Else
            Call Me.BUSCA_ORIGEN(1)
        End If
        Cmd_Imprimir.SetFocus
    End If
End Sub


Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_Origen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.Txt_Origen.Text) = "" Then
            Call Me.BUSCA_ORIGEN(3)
        Else
            Call Me.BUSCA_ORIGEN(1)
        End If
        optComisionista.Enabled = False
        If UCase(Trim(Txt_Origen)) = "E" Then optComisionista.Enabled = True
        Txt_Tipo.SetFocus
    End If
End Sub


Private Sub Txt_Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Set frmBusqGeneral.oParent = Me
       frmBusqGeneral.sQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(Txt_Tipo.Text) & "%'"
       frmBusqGeneral.Cargar_Datos
       frmBusqGeneral.Show 1
    If codigo <> "" Then
        Txt_Tipo.Text = codigo
        Txt_Destipo.Text = Descripcion
        Cmd_Imprimir.SetFocus
    Else
       Txt_Tipo.Text = ""
       Txt_Destipo.Text = ""
    End If
       codigo = ""
            Descripcion = ""

    End If

End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

            Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(txtCod_TipDoc.Text) & "%'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
            If codigo <> "" Then
                txtCod_TipDoc.Text = codigo
                txtDes_TipDoc.Text = Descripcion
                txtSer_Docum.SetFocus
                
            Else
                txtCod_TipDoc.Text = ""
                txtDes_TipDoc.Text = ""
            End If
            codigo = ""
            Descripcion = ""

    End If
End Sub


Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
End Sub



Private Sub txtDescripcionT_GotFocus()
SelectionText txtDescripcionT
End Sub

Private Sub txtDescripcionT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.txtDescripcionT.Text) = "" Then
            Call Me.BUSCA_ORIGEN_T(3)
            
        Else
            Call Me.BUSCA_ORIGEN_T(1)
        End If
        dtpFechaT.SetFocus
    End If
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    SendKeys "{TAB}"
  End If
End Sub



Private Sub txtOrigenT_GotFocus()
SelectionText txtOrigenT
End Sub

Private Sub txtOrigenT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(Me.txtOrigenT.Text) = "" Then
            Call Me.BUSCA_ORIGEN_T(3)
        Else
            Call Me.BUSCA_ORIGEN_T(1)
        End If
        dtpFechaT.SetFocus
    End If
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSer_Docum = Format(txtSer_Docum, "000")
        txtNum_Docum.SetFocus
    End If
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNum_Docum = Format(txtNum_Docum, "00000000")
        cmdBuscar.SetFocus
    End If
End Sub


Public Sub BUSCA_ORIGEN(Tipo As Integer)
On Error GoTo hand
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Des_Origen as 'Descripción' FROM  cn_origen WHERE Origen = '" & Trim(Me.Txt_Origen.Text) & "'"
                    Me.Txt_Descripcion.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As Object
                    Set rs = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Origen as 'Código', Des_Origen as 'Descripción' FROM cn_origen WHERE Des_Origen LIKE '%" & Trim(Me.Txt_Origen) & "%' ORDER BY Des_Origen"
                    Else
                        oTipo.sQuery = "SELECT ORIGEN as 'Código', Des_Origen AS 'Descripción' FROM Cn_Origen ORDER BY Des_Origen"
                    End If
                    
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.Txt_Origen = Trim(codigo)
                        Me.Txt_Descripcion = Trim(Descripcion)
                        
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    
Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"
End Sub



Public Sub BUSCA_ORIGEN_T(Tipo As Integer)
On Error GoTo hand
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Des_Origen as 'Descripción' FROM  cn_origen WHERE Origen = '" & Trim(Me.txtOrigenT.Text) & "'"
                    Me.txtDescripcionT.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As Object
                    Set rs = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Origen as 'Código', Des_Origen as 'Descripción' FROM cn_origen WHERE Des_Origen LIKE '%" & Trim(Me.txtOrigenT) & "%' ORDER BY Des_Origen"
                    Else
                        oTipo.sQuery = "SELECT ORIGEN as 'Código', Des_Origen AS 'Descripción' FROM Cn_Origen ORDER BY Des_Origen"
                    End If
                    
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.txtOrigenT = Trim(codigo)
                        Me.txtDescripcionT = Trim(Descripcion)
                        
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    
Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"
End Sub


Public Function GetRecordset1(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object ' CreateObject("ADODB.Recordset") '
  Dim objAdoCn As Object ' New ADODB.Connection '
  
 
  Set objADORs = CreateObject("ADODB.Recordset") 'CreateObject("ADODB.Recordset") '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  objADORs.Open SQL, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset1 = objADORs
  Set GetRecordset1.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 
Exit Function
ehGetRecordset:
  err.Raise err.Number, err.Source, err.Description
  MsgBox err.Description
End Function

Sub ReporteHistorico()
On Error GoTo ERROR
Dim sSql As String
Dim oo As Object
Dim Ruta As String
Dim Reg1 As ADODB.Recordset

sSql = "Ventas_Muestra_Documentos_por_Cerrar '8','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & "2" & "','N','','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','" & IIf(optCliente, "C", "D") & "','','" & IIf(chkIncluirBoletas, "S", "N") & "'"

Set Reg1 = GetRecordset1(cCONNECT, sSql)

If MsgBox("Desea Imprimir usando Microsoft Excel?", vbYesNo + vbQuestion, "") = vbYes Then
    Ruta = vRuta & "\RptHistoricoPagosCliente.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "Reporte", Reg1
'Else
'    Ruta = vRuta & "\RptHistoricoPagosCliente.OTS"
'    Set oo = CreateObject("ooBusiness.Calc")
'    oo.OfficeTemplateSheet = Ruta
'    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
'    oo.MacroLibraryName = "Library1"
'    oo.MacroModuleName = "Module1"
'    oo.MacroName = "Reporte"
    
'    oo.Run sSQL, cCONNECT
End If
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number
End Sub




'************************************************************************************************************************************************************************************************************************************************************************************************
'==> IMPRESION : AÑO-MES CLIENTE
'************************************************************************************************************************************************************************************************************************************************************************************************
Private Sub txtCodCV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodCV.Text) = "" Then Call BUSCA_CV(3) _
        Else Call BUSCA_CV(1)
        cmdAceptar_AnoMes.SetFocus
    End If
End Sub

Private Sub txtDesCV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesCV.Text) = "" Then Call BUSCA_CV(3) _
        Else Call BUSCA_CV(2)
        cmdAceptar_AnoMes.SetFocus
    End If
End Sub

Public Sub BUSCA_CV(Tipo As Integer)
    On Error GoTo hand
    Select Case Tipo
        Case 1:
            strSQL = "SELECT Des_CondVent as 'Descripción' FROM  lg_condvent WHERE Cod_CondVent = '" & Trim(txtCodCV) & "'"
            txtDesCV.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
            'SELECT Cod_CondVent AS CODIGO, Des_CondVent AS DESCRIPCION FROM lg_condvent ORDER BY Cod_CondVent
        Case 2, 3:
            Dim oTipo As New frmBusqGeneral
            Dim rs As Object
            Set rs = CreateObject("ADODB.Recordset")
            Set oTipo.oParent = Me
            
            If Tipo = 2 Then
                oTipo.sQuery = "SELECT Cod_CondVent AS CODIGO, Des_CondVent AS DESCRIPCION FROM lg_condvent WHERE Des_CondVent LIKE '%" & Trim(txtDesCV) & "%' ORDER BY Cod_CondVent"
            Else
                oTipo.sQuery = "SELECT Cod_CondVent AS CODIGO, Des_CondVent AS DESCRIPCION FROM lg_condvent ORDER BY Cod_CondVent"
            End If
            oTipo.Cargar_Datos
            oTipo.Show 1
            If codigo <> "" Then
                txtCodCV = Trim(codigo)
                txtDesCV = Trim(Descripcion)
            End If
            Set oTipo = Nothing
            Set rs = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    Exit Sub
hand:
ErrorHandler err, "BUSCA CONDICION VENTA"
End Sub

Private Sub cmdCancelar_AnoMes_Click()
    fra_origen.Enabled = True
    fanANO_MES.Visible = False
End Sub

'************************************************************************************************************************************************************************************************************************************************************************************************
'==> IMPRESION : COMISIONISTA
'************************************************************************************************************************************************************************************************************************************************************************************************
Private Sub txtCodComisionista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodComisionista.Text) = "" Then Call BUSCA_COMISIONISTA(3) _
        Else Call BUSCA_COMISIONISTA(1)
        cmdAceptar_Com.SetFocus
    End If
End Sub

Private Sub txtDesComisionista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesComisionista.Text) = "" Then Call BUSCA_COMISIONISTA(3) _
        Else Call BUSCA_COMISIONISTA(2)
        cmdAceptar_Com.SetFocus
    End If
End Sub

Public Sub BUSCA_COMISIONISTA(Tipo As Integer)
    On Error GoTo hand
    Select Case Tipo
        Case 1:
            strSQL = "SELECT nom_comisionista AS DESCRIPCION FROM TG_COMISIONISTA WHERE cod_comisionista = '" & Trim(txtCodComisionista) & "'"
            txtDesCV.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
        Case 2, 3:
            Dim oTipo As New frmBusqGeneral
            Dim rs As Object
            Set rs = CreateObject("ADODB.Recordset")
            Set oTipo.oParent = Me
            
            If Tipo = 2 Then
                oTipo.sQuery = "SELECT cod_comisionista AS CODIGO, nom_comisionista AS DESCRIPCION FROM TG_COMISIONISTA WHERE nom_comisionista LIKE '%" & Trim(txtDesComisionista) & "%' ORDER BY nom_comisionista"
            Else
                oTipo.sQuery = "SELECT cod_comisionista AS CODIGO, nom_comisionista AS DESCRIPCION FROM TG_COMISIONISTA ORDER BY nom_comisionista"
            End If
            oTipo.Cargar_Datos
            oTipo.Show 1
            If codigo <> "" Then
                txtCodComisionista = Trim(codigo)
                txtDesComisionista = Trim(Descripcion)
            End If
            Set oTipo = Nothing
            Set rs = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    Exit Sub
hand:
ErrorHandler err, "BUSCA CONDICION VENTA"
End Sub

Private Sub cmdCancelar_Com_Click()
    fanComisionista.Visible = False
    fra_origen.Enabled = True
End Sub

'************************************************************************************************************************************************************************************************************************************************************************************************
'==> IMPRESION : CLIENTE COMERCIAL
'************************************************************************************************************************************************************************************************************************************************************************************************
Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            Command1.SetFocus
        End If
    End If
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If codigo <> "" Then
        txtAbr_Cliente.Text = codigo
        txtDes_Cliente.Text = Descripcion
        Command1.SetFocus
        codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtDes_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtDes_Cliente.Text) & "%'"
            txtAbr_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            Command1.SetFocus
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtDes_Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo SALTO_ERROR
Dim oRs As New Recordset
Dim sCodCLIENTE As String, Ruta As String
Dim oo As Object
Dim sRutaLogo As String, sClienteComercial As String
    
    sCodCLIENTE = DevuelveCampo("SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'", cCONNECT)
    
    If Trim(sCodCLIENTE) = Empty Then
        MsgBox "Indique el cliente comercial para poder emitir la impresión.....", vbCritical
        Exit Sub
    End If
    
    strSQL = "EXEC cn_ventas_muestra_facturas_exportacion_pendientes_por_cliente_comercial '" & sCodCLIENTE & "'"
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    If oRs.RecordCount = 0 Then
        MsgBox "No se han encontrado datos para la impresión.....", vbExclamation
        Exit Sub
    End If
    
    sClienteComercial = UCase(Trim(txtDes_Cliente)) & " [" & txtAbr_Cliente & "]" & " [" & sCodCLIENTE & "]"
    
    Ruta = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(Ruta, cCONNECT)
    
    If MsgBox("Desea imprimir usando Microsoft Excel?", vbYesNo + vbQuestion, "Imprimir x Cliente") = vbYes Then
        Set oo = CreateObject("excel.application")
        
        oo.Workbooks.Open vRuta & "\rptFacturasPendientesPorClienteComercial.XLT"
        oo.Visible = True
        oo.displayalerts = False
        oo.Run "reporte", sRutaLogo, oRs, sClienteComercial
    Else
        Ruta = vRuta & "\rptFacturasPendientesPorClienteComercial.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run sRutaLogo, strSQL, sClienteComercial, cCONNECT
    End If
    Set oo = Nothing
Exit Sub

SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub Command2_Click()
    famClienteComercial.Visible = False
    fra_origen.Enabled = True
End Sub


Public Sub Busca_Trabajador()
On Error GoTo Fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
      
strSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!codigo)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!codigo), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!codigo), 4)
            'stip_Trabajador = Left(rstAux!codigo, 1)
            'scod_trabajador = Right(rstAux!codigo, 4)
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
    "Búsqueda de Color (" & Opcion & ")"
End Sub
