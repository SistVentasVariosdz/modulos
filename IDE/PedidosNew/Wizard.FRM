VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   Caption         =   "Generación de Purchase Order"
   ClientHeight    =   7980
   ClientLeft      =   1980
   ClientTop       =   1830
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9945
   StartUpPosition =   1  'CenterOwner
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Datos Generales"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7350
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Tag             =   "1000"
      Top             =   0
      Width           =   9945
      Begin VB.OptionButton optFlg_CartaDesaprobada 
         Caption         =   "Option1"
         Height          =   195
         Left            =   3720
         TabIndex        =   50
         Top             =   3975
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optFlg_CartaAprobada 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1710
         TabIndex        =   49
         Top             =   3990
         Width           =   1815
      End
      Begin VB.CommandButton cmdDes_Despacho 
         Caption         =   "Comentario para Despachos"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   47
         Top             =   6030
         Width           =   9720
      End
      Begin VB.CommandButton cmdDes_General 
         Caption         =   "Comentario General"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   46
         Top             =   4785
         Width           =   9735
      End
      Begin VB.TextBox txtDes_Despacho 
         Height          =   1050
         Left            =   150
         TabIndex        =   45
         Top             =   6285
         Width           =   9750
      End
      Begin VB.TextBox txtDes_General 
         Height          =   885
         Left            =   165
         TabIndex        =   44
         Top             =   5055
         Width           =   9750
      End
      Begin VB.TextBox txtPor_Slush 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8310
         TabIndex        =   43
         Top             =   465
         Width           =   750
      End
      Begin VB.TextBox txtCod_Banco 
         Height          =   285
         Left            =   1725
         MaxLength       =   3
         TabIndex        =   40
         Top             =   4365
         Width           =   600
      End
      Begin VB.TextBox txtNom_Banco 
         Height          =   285
         Left            =   2370
         MaxLength       =   30
         TabIndex        =   39
         Top             =   4365
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Moneda 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   37
         Top             =   3555
         Width           =   600
      End
      Begin VB.TextBox txtNom_Moneda 
         Height          =   285
         Left            =   2355
         MaxLength       =   30
         TabIndex        =   36
         Top             =   3555
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   285
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   34
         Top             =   3165
         Width           =   600
      End
      Begin VB.TextBox txtDes_Embarque 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   33
         Top             =   3165
         Width           =   4035
      End
      Begin VB.TextBox txtCod_PagEmb 
         Height          =   285
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   31
         Top             =   2775
         Width           =   600
      End
      Begin VB.TextBox txtDes_PagEmb 
         Height          =   285
         Left            =   2355
         MaxLength       =   30
         TabIndex        =   30
         Top             =   2775
         Width           =   4035
      End
      Begin VB.TextBox txtPor_Comision 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8310
         TabIndex        =   29
         Top             =   105
         Width           =   750
      End
      Begin VB.TextBox txtCod_TemCli 
         Height          =   285
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   26
         Top             =   2100
         Width           =   600
      End
      Begin VB.TextBox txtNom_TemCli 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   25
         Top             =   2100
         Width           =   4035
      End
      Begin VB.TextBox txtNom_DivCli 
         Height          =   285
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1725
         Width           =   4035
      End
      Begin VB.TextBox txtCod_DivCli 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   23
         Top             =   1710
         Width           =   615
      End
      Begin VB.TextBox txtDes_Destino 
         Height          =   285
         Left            =   2325
         MaxLength       =   30
         TabIndex        =   21
         Top             =   915
         Width           =   4050
      End
      Begin VB.TextBox txtCod_Destino 
         Height          =   285
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   20
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox txtNom_Fabrica 
         Height          =   285
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   19
         Top             =   525
         Width           =   4050
      End
      Begin VB.TextBox txtAbr_Fabrica 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   18
         Top             =   525
         Width           =   630
      End
      Begin MSComCtl2.DTPicker dtpFec_DespachoAct 
         Height          =   315
         Left            =   1665
         TabIndex        =   15
         Top             =   1320
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   37159
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   4590
         Left            =   6540
         Top             =   60
         Width           =   90
      End
      Begin VB.Label labels 
         Caption         =   "Estado de Carta Credito"
         Height          =   360
         Index           =   12
         Left            =   195
         TabIndex        =   48
         Top             =   3930
         Width           =   1365
      End
      Begin VB.Label labels 
         Caption         =   "Slush"
         Height          =   255
         Index           =   11
         Left            =   6780
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Banco"
         Height          =   255
         Index           =   10
         Left            =   210
         TabIndex        =   41
         Top             =   4395
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   9
         Left            =   195
         TabIndex        =   38
         Top             =   3585
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Tipo de Embarque"
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   35
         Top             =   3195
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Pago de  Embarque"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   32
         Top             =   2790
         Width           =   1440
      End
      Begin VB.Label labels 
         Caption         =   "Comisión"
         Height          =   255
         Index           =   6
         Left            =   6765
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Temporada"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   27
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "División del Cliente"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   1785
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Destino"
         Height          =   255
         Index           =   3
         Left            =   165
         TabIndex        =   17
         Top             =   930
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fabrica"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   16
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fecha Despacho"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   1380
         Width           =   1200
      End
      Begin MSForms.ComboBox cboCod_ClaPurOrd 
         Height          =   285
         Left            =   1665
         TabIndex        =   13
         Top             =   135
         Width           =   2100
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3704;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label labels 
         Caption         =   "Clase"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   150
         Width           =   1200
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Elegir Estilo Cliente"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7380
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Tag             =   "2000"
      Top             =   -15
      Width           =   9885
      Begin VB.Frame fraTallas 
         Caption         =   "Tallas"
         Height          =   4725
         Left            =   4980
         TabIndex        =   72
         Top             =   2610
         Width           =   4860
         Begin VB.CommandButton Command5 
            Caption         =   "Seleccionados"
            Height          =   210
            Index           =   1
            Left            =   2700
            TabIndex        =   82
            Top             =   270
            Width           =   1995
         End
         Begin VB.CommandButton cmdColores 
            Caption         =   "Para Seleccionar"
            Height          =   210
            Index           =   1
            Left            =   135
            TabIndex        =   81
            Top             =   270
            Width           =   1995
         End
         Begin VB.CommandButton Command4 
            Caption         =   ">"
            Height          =   315
            Left            =   2220
            TabIndex        =   76
            Top             =   1470
            Width           =   360
         End
         Begin VB.CommandButton Command3 
            Caption         =   "<"
            Height          =   315
            Left            =   2220
            TabIndex        =   75
            Top             =   1830
            Width           =   360
         End
         Begin VB.CommandButton Command2 
            Caption         =   ">>"
            Height          =   315
            Left            =   2235
            TabIndex        =   74
            Top             =   2370
            Width           =   360
         End
         Begin VB.CommandButton Command1 
            Caption         =   "<<"
            Height          =   315
            Left            =   2235
            TabIndex        =   73
            Top             =   2730
            Width           =   360
         End
         Begin MSForms.ListBox lstTallas 
            Height          =   4080
            Left            =   135
            TabIndex        =   78
            Top             =   495
            Width           =   1995
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3519;7197"
            MatchEntry      =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox lstTallasSELEC 
            Height          =   4080
            Left            =   2700
            TabIndex        =   77
            Top             =   495
            Width           =   1995
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3519;6932"
            MatchEntry      =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraColores 
         Caption         =   "Colores"
         Height          =   4725
         Left            =   75
         TabIndex        =   65
         Top             =   2610
         Width           =   4860
         Begin VB.CommandButton Command5 
            Caption         =   "Seleccionados"
            Height          =   210
            Index           =   0
            Left            =   2700
            TabIndex        =   80
            Top             =   285
            Width           =   1995
         End
         Begin VB.CommandButton cmdColores 
            Caption         =   "Para Seleccionar"
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   79
            Top             =   285
            Width           =   1995
         End
         Begin VB.CommandButton cmdAllToLeft 
            Caption         =   "<<"
            Height          =   315
            Left            =   2235
            TabIndex        =   69
            Top             =   2730
            Width           =   360
         End
         Begin VB.CommandButton cmdAllToRight 
            Caption         =   ">>"
            Height          =   315
            Left            =   2235
            TabIndex        =   68
            Top             =   2370
            Width           =   360
         End
         Begin VB.CommandButton cmdToLeft 
            Caption         =   "<"
            Height          =   315
            Left            =   2220
            TabIndex        =   67
            Top             =   1830
            Width           =   360
         End
         Begin VB.CommandButton cmdToRight 
            Caption         =   ">"
            Height          =   315
            Left            =   2220
            TabIndex        =   66
            Top             =   1470
            Width           =   360
         End
         Begin MSForms.ListBox lstColores 
            Height          =   4050
            Left            =   135
            TabIndex        =   70
            Top             =   510
            Width           =   1995
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3519;6932"
            MatchEntry      =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox lstColoresSELEC 
            Height          =   4050
            Left            =   2700
            TabIndex        =   71
            Top             =   510
            Width           =   1995
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3519;7144"
            MatchEntry      =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtPor_ComisionLOT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1620
         TabIndex        =   63
         Top             =   2160
         Width           =   750
      End
      Begin VB.TextBox txtCod_FabricaLOT 
         Height          =   300
         Left            =   1635
         MaxLength       =   3
         TabIndex        =   58
         Top             =   930
         Width           =   630
      End
      Begin VB.TextBox txtNom_FabricaLOT 
         Height          =   300
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   57
         Top             =   930
         Width           =   4050
      End
      Begin VB.TextBox txtCod_DestinoLOT 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   56
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDes_DestinoLOT 
         Height          =   285
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   55
         Top             =   1320
         Width           =   4050
      End
      Begin VB.TextBox txtPrecioLOT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1635
         TabIndex        =   53
         Top             =   540
         Width           =   750
      End
      Begin VB.TextBox txtCod_EstCliLOT 
         Height          =   285
         Left            =   1635
         TabIndex        =   51
         Top             =   165
         Width           =   2790
      End
      Begin MSComCtl2.DTPicker dtpFec_DespachoReaLOT 
         Height          =   315
         Left            =   1620
         TabIndex        =   59
         Top             =   1725
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   37159
      End
      Begin VB.Label labels 
         Caption         =   "Comisión"
         Height          =   255
         Index           =   18
         Left            =   135
         TabIndex        =   64
         Top             =   2175
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Fecha Despacho"
         Height          =   255
         Index           =   17
         Left            =   135
         TabIndex        =   62
         Top             =   1785
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fabrica"
         Height          =   240
         Index           =   16
         Left            =   120
         TabIndex        =   61
         Top             =   975
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Destino"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   60
         Top             =   1335
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Precio"
         Height          =   255
         Index           =   14
         Left            =   135
         TabIndex        =   54
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Estilo del Cliente"
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   52
         Top             =   195
         Width           =   1200
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   7410
      Width           =   9945
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   4545
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   11350
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   255
         X2              =   9735
         Y1              =   -30
         Y2              =   -45
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Matriz de Cantidades"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Index           =   2
      Left            =   -10000
      TabIndex        =   8
      Tag             =   "2002"
      Top             =   0
      Width           =   9855
      Begin SSDataWidgets_B.SSDBGrid ssgrdDatos 
         Height          =   7095
         Left            =   105
         TabIndex        =   83
         Top             =   135
         Width           =   9705
         _Version        =   196617
         DataMode        =   2
         HeadLines       =   2
         Col.Count       =   0
         BackColorOdd    =   10354687
         RowHeight       =   423
         ExtraHeight     =   185
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   17119
         _ExtentY        =   12515
         _StockProps     =   79
         Caption         =   "Matriz de Cantidades"
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Matriz de Cantidades y Precios"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Index           =   3
      Left            =   -10000
      TabIndex        =   9
      Tag             =   "2004"
      Top             =   0
      Width           =   9855
      Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
         Height          =   7095
         Left            =   105
         TabIndex        =   84
         Top             =   135
         Width           =   9705
         _Version        =   196617
         DataMode        =   2
         HeadLines       =   2
         Col.Count       =   0
         BackColorOdd    =   10354687
         RowHeight       =   423
         ExtraHeight     =   185
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   17119
         _ExtentY        =   12515
         _StockProps     =   79
         Caption         =   "Matriz de Cantidades y Precios"
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Generar Lote Estilo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Index           =   4
      Left            =   -10000
      TabIndex        =   10
      Tag             =   "3000"
      Top             =   0
      Width           =   9795
      Begin VB.Label lblStepFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "El Sistema generará información de acuerdo a los datos proporcionados por Ud."
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   3240
         TabIndex        =   11
         Tag             =   "3001"
         Top             =   2370
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1065
         Index           =   5
         Left            =   1485
         Picture         =   "Wizard.frx":0000
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent As Object
Const NUM_STEPS = 5

Public aCarga(NUM_STEPS) As Boolean

Const RES_ERROR_MSG = 30000

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Const HELP_BASE = 1000
Const HELP_FILE = "MYWIZARD.HLP"

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_FINISH = 4

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = ""
Const INTRO_KEY = ""
Const SHOW_INTRO = ""
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

'Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
            'wizard creation code goes here
      
            Unload Me
                    
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
        

    SetStep 0, DIR_NONE

End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO
            LoadDataGeneral
        Case STEP_1
            LoadDataColores
            LoadDataTallas
        Case STEP_2
        
        Case STEP_3
            mbFinishOK = False
      
        Case STEP_FINISH
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long

    If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

Private Sub LoadDataGeneral()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_PurOrd
    Dim i As Long
    
    If aCarga(STEP_INTRO) = False Then
        Set obj = New clsTG_PurOrd
        obj.Connect = cCONNECT
        vbuff = obj.ViewAllClaPurOrd
        Set obj = Nothing
        
        Me.cboCod_ClaPurOrd.Clear

        If Not IsEmpty(vbuff) Then
            For i = 0 To UBound(vbuff, 2)
                Me.cboCod_ClaPurOrd.AddItem
                Me.cboCod_ClaPurOrd.Column(0, i) = vbuff(0, i)
                Me.cboCod_ClaPurOrd.Column(1, i) = vbuff(1, i)
                Me.cboCod_ClaPurOrd.Column(2, i) = vbuff(2, i)
            Next
        End If
    End If
    
Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    errores Err.Number
End Sub

Public Sub LoadDataColores()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_PurOrd

    If aCarga(STEP_1) = False Then
        Set obj = New clsTG_PurOrd
        obj.Connect = cCONNECT
        vbuff = obj.ViewAllClaPurOrd
        Set obj = Nothing
        
        lstColores.Clear

        If Not IsEmpty(vbuff) Then
            For i = 0 To UBound(vbuff, 2)
                Me.lstColores.AddItem
                Me.lstColores.Column(0, i) = vbuff(0, i)
                Me.lstColores.Column(1, i) = vbuff(1, i)
                Me.lstColores.Column(2, i) = vbuff(2, i)
            Next
        End If
    End If

Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    errores Err.Number

End Sub

Public Sub LoadDataTallas()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_PurOrd
    
    If aCarga(STEP_2) = False Then
        Set obj = New clsTG_PurOrd
        obj.Connect = cCONNECT
        vbuff = obj.ViewAllClaPurOrd
        Set obj = Nothing
        
        lstTallasSELEC.Clear

        If Not IsEmpty(vbuff) Then
            For i = 0 To UBound(vbuff, 2)
                Me.lstColores.AddItem
                Me.lstColores.Column(0, i) = vbuff(0, i)
                Me.lstColores.Column(1, i) = vbuff(1, i)
                Me.lstColores.Column(2, i) = vbuff(2, i)
            Next
        End If
    End If
Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    errores Err.Number

End Sub

