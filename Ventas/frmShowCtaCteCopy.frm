VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowCtaCteCopy 
   BackColor       =   &H80000013&
   Caption         =   "Consulta Cuenta Corriente Clientes"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBuscar 
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
      Height          =   2205
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   15465
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1920
         TabIndex        =   56
         Top             =   720
         Width           =   6135
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   1980
            TabIndex        =   57
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   37289985
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   4080
            TabIndex        =   58
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   37289985
            CurrentDate     =   37543
         End
         Begin VB.Label Label1 
            Caption         =   "Rango Fecha de Emisión:"
            Height          =   240
            Left            =   0
            TabIndex        =   59
            Top             =   120
            Width           =   2235
         End
      End
      Begin VB.OptionButton opTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton oprCanceladas 
         Caption         =   "Canceladas"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton opPendiente 
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   8880
         TabIndex        =   1
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   52
         Top             =   1710
         Width           =   1980
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   1575
         MaxLength       =   2
         TabIndex        =   51
         Top             =   1710
         Width           =   360
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4605
         MaxLength       =   3
         TabIndex        =   50
         Top             =   1725
         Width           =   540
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   49
         Top             =   1725
         Width           =   1440
      End
      Begin VB.OptionButton optDocRef 
         Caption         =   "Documento Específico"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   2010
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   480
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CheckBox chkLetraTercero 
         Alignment       =   1  'Right Justify
         Caption         =   "&Letras de Tercero"
         Height          =   255
         Left            =   3840
         TabIndex        =   44
         Top             =   1200
         Width           =   1995
      End
      Begin VB.CheckBox chkInLetras 
         Alignment       =   1  'Right Justify
         Caption         =   "&Incluir Letras"
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Top             =   1200
         Width           =   1755
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   690
         MaxLength       =   11
         TabIndex        =   5
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2970
         MaxLength       =   11
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   0
         Top             =   360
         Width           =   4785
      End
      Begin VB.CheckBox chkIncluirBoletas 
         Alignment       =   1  'Right Justify
         Caption         =   "&Incluir Boletas"
         Height          =   255
         Left            =   6240
         TabIndex        =   41
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label Label5 
         Caption         =   "Número :"
         Height          =   225
         Left            =   5250
         TabIndex        =   68
         Tag             =   "Number"
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Serie Docum.:"
         Height          =   375
         Left            =   3990
         TabIndex        =   67
         Top             =   1695
         Width           =   750
      End
      Begin VB.Label lblCod_TipOrdCom 
         Caption         =   "Tipo Documento:"
         Height          =   390
         Left            =   705
         TabIndex        =   66
         Tag             =   "Document Type"
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   11280
         TabIndex        =   65
         Top             =   1785
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Left            =   11820
         TabIndex        =   64
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   11280
         TabIndex        =   63
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   11280
         TabIndex        =   62
         Top             =   930
         Width           =   840
      End
      Begin VB.Label Label11 
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
         Left            =   11280
         TabIndex        =   61
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Ruc :"
         Height          =   180
         Left            =   240
         TabIndex        =   60
         Tag             =   "Anexo Type"
         Top             =   405
         Width           =   435
      End
   End
   Begin VB.Frame fra_origen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   1320
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   9150
      Begin VB.TextBox Txt_Origen 
         Height          =   300
         Left            =   780
         TabIndex        =   37
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox Txt_Descripcion 
         Height          =   300
         Left            =   1125
         TabIndex        =   36
         Top             =   360
         Width           =   2805
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
         TabIndex        =   35
         Top             =   1530
         Width           =   1215
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
         TabIndex        =   34
         Top             =   1530
         Width           =   1095
      End
      Begin VB.TextBox Txt_Tipo 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   5670
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "FA"
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox Txt_Destipo 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   6030
         MaxLength       =   30
         TabIndex        =   32
         Top             =   360
         Width           =   2910
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordenado"
         Height          =   615
         Left            =   180
         TabIndex        =   26
         Top             =   840
         Width           =   8775
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
            TabIndex        =   31
            Top             =   270
            Width           =   1305
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
            TabIndex        =   30
            Top             =   270
            Value           =   -1  'True
            Width           =   945
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
            TabIndex        =   29
            Top             =   270
            Width           =   1725
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
            TabIndex        =   28
            Top             =   270
            Width           =   1725
         End
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
            TabIndex        =   27
            Top             =   270
            Width           =   2115
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   413
         Width           =   465
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
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1965
         Left            =   0
         Top             =   30
         Width           =   9105
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
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   885
      Width           =   1455
   End
   Begin VB.Frame famClienteComercial 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1493
      TabIndex        =   17
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
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
         TabIndex        =   22
         Top             =   195
         Width           =   945
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
         TabIndex        =   21
         Top             =   195
         Width           =   1005
      End
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   2655
         TabIndex        =   20
         Tag             =   "..."
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox txtDes_Cliente 
         Height          =   285
         Left            =   2895
         TabIndex        =   19
         Top             =   225
         Width           =   3105
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   2010
         TabIndex        =   18
         Top             =   210
         Width           =   615
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
         TabIndex        =   23
         Tag             =   "Document Type"
         Top             =   270
         Width           =   1710
      End
      Begin VB.Shape Shape4 
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
      TabIndex        =   11
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
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
         TabIndex        =   15
         Top             =   195
         Width           =   1005
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
         TabIndex        =   14
         Top             =   195
         Width           =   945
      End
      Begin VB.TextBox txtDesComisionista 
         Height          =   300
         Left            =   1950
         TabIndex        =   13
         Top             =   210
         Width           =   4545
      End
      Begin VB.TextBox txtCodComisionista 
         Height          =   300
         Left            =   1500
         TabIndex        =   12
         Top             =   210
         Width           =   435
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   645
         Left            =   30
         Top             =   30
         Width           =   8775
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
         TabIndex        =   16
         Tag             =   "Document Type"
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.Frame fanANO_MES 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1493
      TabIndex        =   2
      Top             =   4110
      Visible         =   0   'False
      Width           =   8805
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
         TabIndex        =   10
         Top             =   255
         Value           =   -1  'True
         Width           =   945
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
         TabIndex        =   9
         Top             =   255
         Width           =   1905
      End
      Begin VB.TextBox txtCodCV 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3180
         TabIndex        =   8
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtDesCV 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3630
         TabIndex        =   7
         Top             =   210
         Width           =   2895
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
         TabIndex        =   6
         Top             =   195
         Width           =   945
      End
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
         TabIndex        =   3
         Top             =   195
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   645
         Left            =   30
         Top             =   30
         Width           =   8775
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4860
      Left            =   0
      TabIndex        =   69
      Top             =   2280
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   8573
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
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
      BackColorBkg    =   12648447
      ColumnHeaderHeight=   495
      ColumnsCount    =   2
      Column(1)       =   "frmShowCtaCteCopy.frx":0000
      Column(2)       =   "frmShowCtaCteCopy.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowCtaCteCopy.frx":016C
      FormatStyle(2)  =   "frmShowCtaCteCopy.frx":02A4
      FormatStyle(3)  =   "frmShowCtaCteCopy.frx":0354
      FormatStyle(4)  =   "frmShowCtaCteCopy.frx":0408
      FormatStyle(5)  =   "frmShowCtaCteCopy.frx":04E0
      FormatStyle(6)  =   "frmShowCtaCteCopy.frx":0598
      FormatStyle(7)  =   "frmShowCtaCteCopy.frx":0678
      FormatStyle(8)  =   "frmShowCtaCteCopy.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmShowCtaCteCopy.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   525
      Left            =   0
      TabIndex        =   70
      Top             =   7200
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   926
      Custom          =   $"frmShowCtaCteCopy.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   500
      ControlSeparator=   10
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   3855
      Left            =   0
      TabIndex        =   71
      Top             =   7920
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   6800
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   12648447
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowCtaCteCopy.frx":0B43
      Column(2)       =   "frmShowCtaCteCopy.frx":0C0B
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowCtaCteCopy.frx":0CAF
      FormatStyle(2)  =   "frmShowCtaCteCopy.frx":0DE7
      FormatStyle(3)  =   "frmShowCtaCteCopy.frx":0E97
      FormatStyle(4)  =   "frmShowCtaCteCopy.frx":0F4B
      FormatStyle(5)  =   "frmShowCtaCteCopy.frx":1023
      FormatStyle(6)  =   "frmShowCtaCteCopy.frx":10DB
      ImageCount      =   0
      PrinterProperties=   "frmShowCtaCteCopy.frx":11BB
   End
End
Attribute VB_Name = "frmShowCtaCteCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrEstus As String, strSQL As String
Public codigo, Descripcion As String, TipoAdd As String, strCod_Anxo As String
Dim OP_Opcion As String, sSQL As String
Public oGroup As GridEX20.JSGroup
Public oFormat As JSFormatStyle


Private Sub Cmd_Cancelar_Click()
  fra_origen.Visible = False
  optCliente.Value = True

End Sub

Sub Reporte_Masivo()
On Error GoTo ERROR
Dim sSQL As String
Dim oo As Object
Dim Ruta As String
Dim Reg1 As ADODB.Recordset, sRutaLogo As String

sSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(sSQL, cCONNECT)
    

'sSQL = "Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & Txt_Origen & "','" & Txt_Tipo & "','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"

 sSQL = "Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & Txt_Origen & "','" & Txt_Tipo & "','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','" & IIf(optCliente, "C", "D") & "','','" & IIf(chkIncluirBoletas, "S", "N") & "'"

'sSQL = "Ventas_Muestra_Documentos_por_Cerrar '" & OP_Opcion & "','" & txtCod_TipAne & "','" & txtCod_Anexo & "','2','N','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & _
'"','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"


Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

Set Reg1 = GetRecordset1(cCONNECT, sSQL)

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
    
    oo.Run sSQL, cCONNECT, sRutaLogo
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
    
    If optClienteComercial.Value = True Then
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
    
If gridex1.RowCount = 0 Then Exit Sub

If MsgBox("Imprimir reporte usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", gridex1.ADORecordset, sRutaLogo
    Set oo = Nothing
Else
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes.OTS"
    
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run sSQL, cCONNECT, sRutaLogo
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

Private Sub dtpFecEmiFin_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Final no puede ser menor a la fecha Inicial", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub dtpFecEmiIni_Change()
  gridex1.ClearFields
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

End Sub

Private Sub cmdBuscar_Click()
Dim vNum_Corre As String
  Buscar
  vNum_Corre = Trim$(gridex1.Value(gridex1.Columns("num_corre").Index))
  
  Call BUSCARDETFACTURA(vNum_Corre)
  
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Buscar()

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

sSQL = "Ventas_Muestra_Documentos_por_Cerrar '" & OP_Opcion & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','2','N','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & _
"','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','" & IIf(chkLetraTercero, "X", "") & "','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','C','','" & IIf(chkIncluirBoletas, "S", "N") & "'"

gridex1.ClearFields

gridex1.DefaultGroupMode = jgexDGMExpanded
Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)


Configurar

End Sub
Sub Configurar()

Set oGroup = gridex1.Groups.Add(gridex1.Columns("Cliente").Index, jgexSortAscending)

gridex1.BackColorRowGroup = &HC0FFFF

Txt_Importe = Format(gridex1.Value(gridex1.Columns("SALDO_TOTAl").Index), "##,##0.00")
TxtDDolares = Format(gridex1.Value(gridex1.Columns("SALDO_total_DOLARES").Index), "##,##0.00")
TxtDsoles = Format(gridex1.Value(gridex1.Columns("SALDO_total_SOLES").Index), "##,##0.00")
TxtDOtros = Format(gridex1.Value(gridex1.Columns("SALDO_total_OTROS").Index), "##,##0.00")

gridex1.Columns("Cod_Tipdoc").Caption = "Tipo"
gridex1.Columns("Cod_Tipdoc").Width = 600

gridex1.Columns("SALDO_TOTAl").Visible = False
gridex1.Columns("Ruc").Visible = False


gridex1.Columns("Cliente").Width = 0
gridex1.Columns("Num_Corre").Width = 0
gridex1.Columns("saldo_equivalente").Width = 0
gridex1.Columns("Anexo_Contable").Width = 0

gridex1.Columns("SALDO_TOTAl").Visible = False
gridex1.Columns("SALDO_total_SOLES").Visible = False
gridex1.Columns("SALDO_total_DOLARES").Visible = False
gridex1.Columns("SALDO_total_otros").Visible = False

gridex1.Columns("Imp_Total").Width = 900
gridex1.Columns("Imp_Total").Caption = "Imp Total"

gridex1.Columns("Saldo_Dolares").Width = 900
gridex1.Columns("Saldo_Dolares").Caption = "Saldo Dolares"
gridex1.Columns("Saldo_Dolares").Visible = True


gridex1.Columns("Saldo_Soles").Width = 900
gridex1.Columns("Saldo_Soles").Caption = "Saldo Soles"
gridex1.Columns("Saldo_Soles").Visible = True

gridex1.Columns("Saldo_Otros").Width = 900
gridex1.Columns("Saldo_Otros").Caption = "Saldo Otra Moneda"
gridex1.Columns("Saldo_Otros").Visible = True

gridex1.Columns("Importe_Cancelado").Width = 900
gridex1.Columns("Importe_Cancelado").Caption = "Imp Cancelado"

gridex1.Columns("saldo_equivalente").Width = 1300
gridex1.Columns("saldo_equivalente").Caption = "saldo Equivalente"

gridex1.Columns("Flg_Status_DrawBack").Visible = False
gridex1.Columns("Des_Status").Visible = False

gridex1.Columns("Fec_Emision").Width = 1125
'GridEX1.Columns("Fec_VenDoc").Width = 1080
gridex1.Columns("Num_Registro").Width = 1155
gridex1.Columns("Moneda").Width = 720

gridex1.Columns("Importe_Dolares").Width = 900
gridex1.Columns("Importe_Dolares").Caption = "Importe Dolares"
gridex1.Columns("Importe_Dolares").Visible = False


gridex1.Columns("Importe_Soles").Width = 900
gridex1.Columns("Importe_Soles").Caption = "Importe Soles"
gridex1.Columns("Importe_Soles").Visible = False


gridex1.Columns("Importe_Otros").Width = 900
gridex1.Columns("Importe_Otros").Caption = "Importe Otros"
gridex1.Columns("Importe_Otros").Visible = False


gridex1.Columns("Imp_En_Letras_Planeado").Width = 1100
gridex1.Columns("Imp_En_Letras_Planeado").Caption = "Imp en Letras Planeado"

gridex1.Columns("Fec_Ult_Pago").Width = 1100

gridex1.Columns("Fec_Registro").Width = 1100
gridex1.Columns("Fec_Venc").Width = 1100

gridex1.Columns("Tipo_Cambio").Width = 1100

gridex1.Columns("Tipo_Cambio_Otra_Moneda").Width = 1100
gridex1.Columns("Tipo_Cambio_Otra_Moneda").Caption = "Tip.Cambio Otra Moneda"

If txtCod_TipAne = "" Then gridex1.DefaultGroupMode = jgexDGMCollapsed Else gridex1.DefaultGroupMode = jgexDGMExpanded

gridex1.ContinuousScroll = True

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
  
Case Is = "IMPFACTURA"
Dim vNum_Corre As String
Dim vImp_Total As Double
Dim vCod_TipDoc As String
vCod_TipDoc = Trim$(gridex1.Value(gridex1.Columns("Cod_Tipdoc").Index))
vNum_Corre = Trim$(gridex1.Value(gridex1.Columns("num_corre").Index))
vImp_Total = Trim$(gridex1.Value(gridex1.Columns("imp_total").Index))
vImp_Total = CDbl(vImp_Total)
Call Imprimir(vNum_Corre, vImp_Total, False, vCod_TipDoc)

Case Is = "SALIR"
Unload Me

End Select
End Sub

Private Sub GridEX1_Click()
Call BUSCARDETALLEFACTURA
End Sub

Private Sub GridEX1_DblClick()
  If gridex1.RowCount = 0 Then Exit Sub
  Load frmShowCtaCteDet
  frmShowCtaCteDet.Caption = "Detalle Cliente " & gridex1.Value(gridex1.Columns("Cliente").Index) & " Documento : " & gridex1.Value(gridex1.Columns("Documento").Index)
  frmShowCtaCteDet.strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & gridex1.Value(gridex1.Columns("NUM_CORRE").Index) & "'"
  frmShowCtaCteDet.Buscar
  frmShowCtaCteDet.Show vbModal
End Sub




Private Sub GridEX2_DblClick()
Dim vNum_Corre As String
Dim vImp_Total As Double
Dim vCod_TipDoc As String
vCod_TipDoc = Trim$(gridex1.Value(gridex1.Columns("Cod_Tipdoc").Index))
vNum_Corre = Trim$(gridex1.Value(gridex1.Columns("num_corre").Index))
vImp_Total = Trim$(gridex1.Value(gridex1.Columns("imp_total").Index))
vImp_Total = CDbl(vImp_Total)
Call Imprimir(vNum_Corre, vImp_Total, False, vCod_TipDoc)
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

Private Sub optDocRef_Click()
    StrEstus = "R"
    OP_Opcion = "4"
    txtCod_TipDoc.SetFocus
End Sub



Private Sub Option1_Click()
Op: OP_Opcion = "1"
End Sub

Private Sub opTodas_Click()
StrEstus = "T"
OP_Opcion = "3"
End Sub

Sub LimpiaFr()
  gridex1.ClearFields
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
       frmBusqGeneral.SQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(Txt_Tipo.Text) & "%'"
       frmBusqGeneral.CARGAR_DATOS
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
                frmBusqGeneral.SQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(txtCod_TipDoc.Text) & "%'"
                frmBusqGeneral.CARGAR_DATOS
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
   'SendKeys "{TAB}"
 
 ' FunctButt1.SetFocus
End Sub


Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    SendKeys "{TAB}"
   
    
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
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Origen as 'Código', Des_Origen as 'Descripción' FROM cn_origen WHERE Des_Origen LIKE '%" & Trim(Me.Txt_Origen) & "%' ORDER BY Des_Origen"
                    Else
                        oTipo.SQuery = "SELECT ORIGEN as 'Código', Des_Origen AS 'Descripción' FROM Cn_Origen ORDER BY Des_Origen"
                    End If
                    
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.Txt_Origen = Trim(codigo)
                        Me.Txt_Descripcion = Trim(Descripcion)
                        
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
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
Dim sSQL As String
Dim oo As Object
Dim Ruta As String
Dim Reg1 As ADODB.Recordset

sSQL = "Ventas_Muestra_Documentos_por_Cerrar '8','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & "2" & "','N','','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','','','','" & vusu & "','" & IIf(chkInLetras, "S", "N") & "','" & IIf(optCliente, "C", "D") & "','','" & IIf(chkIncluirBoletas, "S", "N") & "'"

Set Reg1 = GetRecordset1(cCONNECT, sSQL)

If MsgBox("Desea Imprimir usando Microsoft Excel?", vbYesNo + vbQuestion, "") = vbYes Then
    Ruta = vRuta & "\RptHistoricoPagosCliente.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "Reporte", Reg1
Else
    Ruta = vRuta & "\RptHistoricoPagosCliente.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run sSQL, cCONNECT
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
            Dim RS As Object
            Set RS = CreateObject("ADODB.Recordset")
            Set oTipo.oParent = Me
            
            If Tipo = 2 Then
                oTipo.SQuery = "SELECT Cod_CondVent AS CODIGO, Des_CondVent AS DESCRIPCION FROM lg_condvent WHERE Des_CondVent LIKE '%" & Trim(txtDesCV) & "%' ORDER BY Cod_CondVent"
            Else
                oTipo.SQuery = "SELECT Cod_CondVent AS CODIGO, Des_CondVent AS DESCRIPCION FROM lg_condvent ORDER BY Cod_CondVent"
            End If
            oTipo.CARGAR_DATOS
            oTipo.Show 1
            If codigo <> "" Then
                txtCodCV = Trim(codigo)
                txtDesCV = Trim(Descripcion)
            End If
            Set oTipo = Nothing
            Set RS = Nothing
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
            Dim RS As Object
            Set RS = CreateObject("ADODB.Recordset")
            Set oTipo.oParent = Me
            
            If Tipo = 2 Then
                oTipo.SQuery = "SELECT cod_comisionista AS CODIGO, nom_comisionista AS DESCRIPCION FROM TG_COMISIONISTA WHERE nom_comisionista LIKE '%" & Trim(txtDesComisionista) & "%' ORDER BY nom_comisionista"
            Else
                oTipo.SQuery = "SELECT cod_comisionista AS CODIGO, nom_comisionista AS DESCRIPCION FROM TG_COMISIONISTA ORDER BY nom_comisionista"
            End If
            oTipo.CARGAR_DATOS
            oTipo.Show 1
            If codigo <> "" Then
                txtCodComisionista = Trim(codigo)
                txtDesComisionista = Trim(Descripcion)
            End If
            Set oTipo = Nothing
            Set RS = Nothing
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
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtAbr_Cliente.Text = codigo
        txtDes_Cliente.Text = Descripcion
        Command1.SetFocus
        codigo = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
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




Public Sub BUSCARDETALLEFACTURA()
On Error GoTo Err_Buscar
 Dim vNum_Corre As String
 vNum_Corre = Trim$(gridex1.Value(gridex1.Columns("num_corre").Index))


 
 
 strSQL = "exec Ventas_Muestra_Detalle_Factura_Items '" & vNum_Corre & "'"
 
Set GridEX2.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
'GridEX1.FrozenColumns = 4
GridEX2.Columns("num_corre").Visible = False
GridEX2.Columns("T").Visible = False
GridEX2.Columns("Secuencia").Visible = False
GridEX2.Columns("Porcentaje_Commision").Visible = False
GridEX2.Columns("Cantidad_Item_NC_ND").Visible = False
GridEX2.Columns("Origen").Visible = False
GridEX2.Columns("Articulo").Width = 4500
GridEX2.Columns("Cantidad").Format = "#,##0"
GridEX2.Columns("valor_unitario").Format = "#,##0.00"
GridEX2.Columns("valor_venta").Format = "#,##0.00"
Exit Sub
Err_Buscar:
  If IsNull(vNum_Corre) Then Exit Sub
  GridEX2.ClearFields
   ' MsgBox err.Description, vbCritical + vbOKOnly, "Ventas"
End Sub

Public Function BUSCARDETFACTURA(ByVal numCorre As String)
On Error GoTo Err_Buscar
   
strSQL = "exec Ventas_Muestra_Detalle_Factura_Items '" & numCorre & "'"
 
Set GridEX2.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
'GridEX1.FrozenColumns = 4
GridEX2.Columns("num_corre").Visible = False
GridEX2.Columns("T").Visible = False
GridEX2.Columns("Secuencia").Visible = False
GridEX2.Columns("Porcentaje_Commision").Visible = False
GridEX2.Columns("Cantidad_Item_NC_ND").Visible = False
GridEX2.Columns("Origen").Visible = False
GridEX2.Columns("Articulo").Width = 4500
GridEX2.Columns("Cantidad").Format = "#,##0"
GridEX2.Columns("valor_unitario").Format = "#,##0.00"
GridEX2.Columns("valor_venta").Format = "#,##0.00"
Exit Function
Err_Buscar:
    MsgBox err.Description, vbCritical + vbOKOnly, "Ventas"
End Function
