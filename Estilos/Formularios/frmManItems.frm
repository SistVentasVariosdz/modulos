VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmManItems 
   Caption         =   "Items"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   1290
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11685
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlDirIcono 
      Left            =   11280
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.BMP, *.JPG, *.GIF"
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
      Height          =   1170
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      Begin VB.Frame fraoptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   1440
         TabIndex        =   68
         Top             =   120
         Width           =   6135
         Begin VB.OptionButton optcliente 
            Caption         =   "Cliente"
            Height          =   300
            Left            =   4560
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton optfamitem 
            Caption         =   "Familia de Item"
            Height          =   330
            Left            =   600
            TabIndex        =   11
            Top             =   0
            Width           =   1550
         End
         Begin VB.OptionButton optitem 
            Caption         =   "Item"
            Height          =   300
            Left            =   2760
            TabIndex        =   12
            Top             =   0
            Width           =   1425
         End
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   7800
         TabIndex        =   10
         Top             =   360
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
      Begin VB.Frame Frafamilia 
         Height          =   645
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtdes_famgruite 
            Height          =   285
            Left            =   5680
            MaxLength       =   50
            TabIndex        =   9
            Top             =   270
            Width           =   1695
         End
         Begin VB.TextBox txtdes_famitem 
            Height          =   285
            Left            =   2400
            TabIndex        =   5
            Top             =   270
            Width           =   1575
         End
         Begin VB.CommandButton cmdBusFamItem 
            Caption         =   "..."
            Height          =   330
            Left            =   2085
            TabIndex        =   4
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtfamilia 
            Height          =   285
            Left            =   1605
            MaxLength       =   2
            TabIndex        =   3
            Top             =   240
            Width           =   525
         End
         Begin VB.CommandButton cmdBusgrupo 
            Caption         =   "..."
            Height          =   330
            Left            =   5325
            TabIndex        =   8
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtgrupo 
            Height          =   285
            Left            =   4605
            TabIndex        =   7
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Familia de Item"
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Left            =   4125
            TabIndex        =   6
            Top             =   345
            Width           =   435
         End
      End
      Begin VB.Frame Fracliente 
         Height          =   640
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   330
            Left            =   1830
            TabIndex        =   79
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtcliente 
            Height          =   285
            Left            =   1110
            MaxLength       =   5
            TabIndex        =   78
            Top             =   270
            Width           =   765
         End
         Begin VB.TextBox txttemporada 
            Height          =   285
            Left            =   4830
            TabIndex        =   77
            Top             =   270
            Width           =   735
         End
         Begin VB.CommandButton cmdBusTemporada 
            Caption         =   "..."
            Height          =   330
            Left            =   5550
            TabIndex        =   76
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtNom_TemCli 
            Height          =   285
            Left            =   5910
            TabIndex        =   75
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   285
            Left            =   2190
            TabIndex        =   74
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Temporada"
            Height          =   195
            Left            =   3990
            TabIndex        =   80
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Frame Fraitem 
         Height          =   640
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton cmdBusItem 
            Caption         =   "..."
            Height          =   330
            Left            =   2520
            TabIndex        =   82
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtdes_item 
            Height          =   285
            Left            =   2880
            TabIndex        =   72
            Top             =   240
            Width           =   4200
         End
         Begin VB.TextBox txtcod_item 
            Height          =   285
            Left            =   1590
            MaxLength       =   8
            TabIndex        =   71
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "Item"
            Height          =   240
            Left            =   360
            TabIndex        =   73
            Top             =   330
            Width           =   690
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
      TabIndex        =   65
      Top             =   1200
      Width           =   11580
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   1515
         Left            =   165
         TabIndex        =   64
         Top             =   240
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   2672
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Cod_Item"
            Caption         =   "Item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Des_Item"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cod_UniMed"
            Caption         =   "U.Medida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Cod_FamItem"
            Caption         =   "Familia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Fec_Creacion"
            Caption         =   "Fec. Creación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Ser_OrdComp"
            Caption         =   "Ser. O.C."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Cod_OrdComp"
            Caption         =   "O.C."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Pre_UltComp"
            Caption         =   "Pre. Ult. Compra"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Fec_UltComp"
            Caption         =   "Fec. Ult. Compra"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdLast 
      Height          =   495
      Left            =   3405
      Picture         =   "frmManItems.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Ultimo"
      Top             =   8145
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   495
      Left            =   1920
      Picture         =   "frmManItems.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Primero"
      Top             =   8145
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Height          =   495
      Left            =   2910
      Picture         =   "frmManItems.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Siguiente"
      Top             =   8145
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   495
      Left            =   2400
      Picture         =   "frmManItems.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Anterior"
      Top             =   8145
      Width           =   495
   End
   Begin VB.Frame Fraopciones 
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
      Height          =   870
      Left            =   60
      TabIndex        =   14
      Top             =   3030
      Width           =   11550
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   240
         TabIndex        =   66
         Top             =   225
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   900
         Custom          =   $"frmManItems.frx":05C8
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   60
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4560
      TabIndex        =   31
      Top             =   8115
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManItems.frx":088C
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
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
      Height          =   4170
      Left            =   60
      TabIndex        =   67
      Tag             =   "Detail"
      Top             =   3900
      Width           =   11550
      Begin VB.Frame Fra_cambio 
         Caption         =   "Cambio Cuenta Contable"
         Height          =   1935
         Left            =   3780
         TabIndex        =   103
         Top             =   510
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Cmd_Cancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   2040
            TabIndex        =   107
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton Cmd_Aceptar 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   840
            TabIndex        =   106
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txt_Contable 
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
            Left            =   1320
            TabIndex        =   104
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Cont :"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   480
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker DTPUltCompra 
         Height          =   255
         Left            =   4920
         TabIndex        =   98
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   55640065
         CurrentDate     =   38551
      End
      Begin VB.TextBox txtComentario 
         Height          =   495
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2040
         Width           =   5205
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identificador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   55
         Top             =   2520
         Width           =   6615
         Begin VB.ComboBox CboIde_PO 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   225
            Width           =   585
         End
         Begin VB.ComboBox cboIde_Destino 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   225
            Width           =   585
         End
         Begin VB.ComboBox cboIde_Color 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   225
            Width           =   585
         End
         Begin VB.ComboBox cboIde_EsCli 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   240
            Width           =   585
         End
         Begin VB.ComboBox cboIde_Talla 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "P.O. :"
            Height          =   195
            Left            =   5520
            TabIndex        =   89
            Top             =   285
            Width           =   405
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Destino :"
            Height          =   195
            Left            =   4080
            TabIndex        =   59
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label16 
            Caption         =   "Color Cliente :"
            Height          =   390
            Left            =   2760
            TabIndex        =   58
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label17 
            Caption         =   "Estilo Cliente :"
            Height          =   375
            Left            =   1380
            TabIndex        =   57
            Top             =   135
            Width           =   885
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Talla :"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   290
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdDirIcono 
         Caption         =   "..."
         Height          =   285
         Left            =   11040
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDir_Icono 
         Height          =   285
         Left            =   9540
         TabIndex        =   24
         Top             =   1320
         Width           =   1515
      End
      Begin VB.ComboBox cboFlg_Status 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_MotPrePro 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtDesItem 
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
         MaxLength       =   100
         TabIndex        =   17
         Top             =   240
         Width           =   5640
      End
      Begin VB.ComboBox cboCod_UniMed 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_Origen 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_ClaItem 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_GruItem 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cboCod_FamItem 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtcoditem 
         BackColor       =   &H80000004&
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
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtcta_cont 
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
         Left            =   9480
         TabIndex        =   21
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txtCan_LotPed 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         TabIndex        =   33
         Text            =   "0"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCan_PtoReor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9540
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtRep_PreDol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9540
         TabIndex        =   23
         Text            =   "0"
         Top             =   960
         Width           =   1875
      End
      Begin VB.Frame FrameMixtas 
         Caption         =   "Caracteristicas mixtas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   6945
         TabIndex        =   83
         Top             =   1920
         Width           =   4560
         Begin VB.ComboBox cboCod_HilTel 
            Height          =   315
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   1470
            Width           =   4410
         End
         Begin VB.TextBox TxtMerma 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   28
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton CmdHilado 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   3120
            TabIndex        =   30
            Top             =   1800
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox TxtHilado 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1845
            Width           =   1635
         End
         Begin VB.ComboBox CboTipCar 
            Height          =   315
            ItemData        =   "frmManItems.frx":0A44
            Left            =   1320
            List            =   "frmManItems.frx":0A46
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   2835
         End
         Begin VB.TextBox TxtFacConv 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   27
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Hilo Nuevo"
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   1290
            Width           =   795
         End
         Begin VB.Label Label25 
            Caption         =   "Merma Ten. (%):"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1000
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Hil.Antiguo:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label Label23 
            Caption         =   "F.Conv.(Mts/Kg)"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   680
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "Caracteristica :"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   340
            Width           =   1215
         End
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Ult. Compra"
         Height          =   195
         Left            =   3600
         TabIndex        =   99
         Top             =   3300
         Width           =   1185
      End
      Begin VB.Label LblOrdCompra 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   97
         Top             =   3600
         Width           =   1905
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
         Height          =   195
         Left            =   3600
         TabIndex        =   96
         Top             =   3690
         Width           =   1020
      End
      Begin VB.Label LblMoneda 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2760
         TabIndex        =   95
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   2040
         TabIndex        =   94
         Top             =   3300
         Width           =   630
      End
      Begin VB.Label LblProveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   93
         Top             =   3600
         Width           =   2265
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   3690
         Width           =   780
      End
      Begin VB.Label LblPrecio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   91
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Left            =   240
         TabIndex        =   90
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Comentario :"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   2130
         Width           =   885
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Dir. Icono"
         Height          =   195
         Left            =   8190
         TabIndex        =   53
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label Label18 
         Caption         =   "Status :"
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Preproduc :"
         Height          =   195
         Left            =   3840
         TabIndex        =   52
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label13 
         Caption         =   "Origen :"
         Height          =   255
         Left            =   3870
         TabIndex        =   51
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Pre. Repos($) :"
         Height          =   195
         Left            =   8190
         TabIndex        =   49
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Clase de Item :"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Cont :"
         Height          =   195
         Left            =   8190
         TabIndex        =   45
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Grupo de Item :"
         Height          =   255
         Left            =   150
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Medida :"
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
         Left            =   120
         TabIndex        =   44
         Tag             =   "Porcentaje :"
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lblCod_Item 
         AutoSize        =   -1  'True
         Caption         =   "Item :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Tag             =   "Hilado :"
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Familia Item:"
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
         Left            =   120
         TabIndex        =   42
         Tag             =   "Mat. Prima :"
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pto de Reorden :"
         Height          =   195
         Left            =   8190
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lote Pedido :"
         Height          =   195
         Left            =   3840
         TabIndex        =   48
         Top             =   720
         Width           =   945
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   8520
      TabIndex        =   102
      Top             =   8160
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmManItems.frx":0A48
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin FunctionsButtons.FunctButt FunctButt3 
      Height          =   495
      Left            =   10200
      TabIndex        =   108
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~UBICACION~Verdadero~Verdadero~&Imprimir Ubicación~0~0~1~~0~Falso~Falso~&Imprimir Ubicación~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   300
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmManItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim Opcion As Integer
Dim sTipo As String
Dim rslista As ADODB.Recordset
Dim varCod_item As String
Dim vCod_hiltel As String, sConta As Integer
Dim strSQL As String

Private Sub cbogrupo_Click()
    Call CargaLista
End Sub



Private Sub cboCod_FamItem_Click()
    Dim strSQL As String
    'Combo Grupo Item
    cboCod_GruItem.Clear
    strSQL = "SELECT des_famgruite + space(100) + Cod_Gruitem FROM LG_FamGruIte WHERE Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
    Call LlenaCombo(cboCod_GruItem, strSQL, cCONNECT)
    
    strSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
    If Trim(cboCod_FamItem.Text) <> "" Then
        If DevuelveCampo(strSQL, cCONNECT) = "M" And (sTipo = "I" Or sTipo = "U") Then
            sConta = DevuelveCampo("select count(*) from LG_Autorizacion_Campos where cod_usuario='" & vusu & "' and Tipo_Autorizacion ='1'", cCONNECT)
            If sConta > 0 Then
                HABILITA_CARACMXT True
            End If
        Else
            HABILITA_CARACMXT False
        End If
    End If
End Sub

Private Sub cboCod_HilTel_Click()
vCod_hiltel = Trim(Right(cboCod_HilTel.Text, 10))
TxtHilado = Trim(Right(cboCod_HilTel.Text, 10))
End Sub




Private Sub CboTipCar_Click()
If Trim(Right(CboTipCar.Text, 2)) = "T" Or Trim(Right(CboTipCar.Text, 2)) = "C" Then
    cboCod_HilTel.Enabled = True
    'TxtHilado.Enabled = True
    'CmdHilado.Enabled = True
Else
    cboCod_HilTel.Enabled = False
    'TxtHilado.Enabled = False
    'CmdHilado.Enabled = False
End If
End Sub

Private Sub Cmd_Aceptar_Click()
CAMBIO
 Call CargaLista
 Fra_cambio.Visible = False
End Sub

Private Sub Cmd_Cancelar_Click()
Fra_cambio.Visible = False
End Sub

Private Sub cmdBusItem_Click()
    Dim strSQL As String
    If Trim(txtcod_item.Text) <> "" Then
        strSQL = "SELECT Cod_Item as Código, Des_Item as Descripción FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
    Else
        If Len(Trim(txtDes_Item.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbExclamation)
            Exit Sub
        Else
            strSQL = "SELECT Cod_Item as Código, Des_Item as Descripción  FROM LG_ITEM WHERE Des_Item LIKE '" & Trim(txtDes_Item.Text) & "%'"
        End If
    End If
    
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = strSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcod_item.Text = Codigo
        txtDes_Item.Text = Descripcion
        FunctBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdDirIcono_Click()
    cdlDirIcono.ShowOpen
    If cdlDirIcono.FileName <> "" Then
       txtDir_Icono.Text = cdlDirIcono.FileName
    End If
End Sub

Private Sub CmdHilado_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    'oTipo.sQuery = "select cod_hiltel as codigo,des_hiltel as descripcion from it_hilado ORDER BY des_hiltel"
    oTipo.sQuery = "SM_Muestra_It_Hilado"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        vCod_hiltel = Mid(Codigo, 1, 10)
        TxtHilado.Text = Mid(Codigo, 11) 'Descripcion
        Codigo = ""
        Descripcion = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rslista.State <> 1 Then
    Exit Sub
End If
If Not rslista.EOF And Not rslista.BOF Then
    Call CargaDatos
End If
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcliente.Text = Codigo
        txtNom_Cliente.Text = Descripcion
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdBusFamItem_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT cod_famitem as Codigo, des_famitem as Descripcion FROM LG_FamIte ORDER BY cod_famitem"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtfamilia.Text = Codigo
        txtdes_famitem.Text = Descripcion
        
        txtgrupo.Enabled = True
        cmdBusgrupo.Enabled = True
        txtgrupo.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdBusgrupo_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT  Cod_Gruitem as Código, des_famgruite as Descripción FROM LG_FamGruIte WHERE Cod_Famitem='" & Trim(txtfamilia.Text) & "'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtgrupo.Text = Codigo
        txtdes_famgruite.Text = Descripcion
        FunctBuscar.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdBusTemporada_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Set oTipo.oParent = Me
    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    oTipo.sQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(strSQL, cCONNECT) & "'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txttemporada.Text = Codigo
        txtNom_TemCli.Text = Descripcion
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdFirst_Click()
    If Not rslista.BOF Then
        rslista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not rslista.EOF Then
        rslista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not rslista.EOF Then
        rslista.MoveNext
        If rslista.EOF Then
            rslista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not rslista.BOF Then
        rslista.MovePrevious
        If rslista.BOF Then
            rslista.MoveFirst
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    'Combo Familia Item
    If Opcion = 1 And txtfamilia.Text <> "" Then
        strSQL = "SELECT des_famitem + space(100) + cod_famitem  FROM LG_FamIte"
        Call LlenaCombo(cboCod_FamItem, strSQL, cCONNECT)
        Call BuscaCombo(txtfamilia.Text, 2, cboCod_FamItem)
    End If
    
    strSQL = "SM_Muestra_It_Hilado"
    Call LlenaCombo(cboCod_HilTel, strSQL, cCONNECT)

End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    Call CargaCombos
    Opcion = 1
    Call CargaLista
    
    FormateaGrid Me.DGridLista
    INHABILITA_DATOS
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
   

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
    Dim strSQL As String
    Dim vericono As Integer
    Select Case ActionName
        Case "TEMPORADA"
            If Not rslista.EOF And Not rslista.BOF Then
                Load frmMantItemTemCli
                frmMantItemTemCli.Codigo_item = rslista("Cod_item")

                frmMantItemTemCli.Carga_Datos
                frmMantItemTemCli.Show 1
            Else
                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
            End If
         Case "IMPRIMIR"
            If Not rslista.EOF And Not rslista.BOF Then

                vericono = MsgBox("¿Desea visualizar la muestra en el reporte?", vbInformation + vbYesNo)
                If vericono = vbYes Then
                    vericono = 1
                Else
                    vericono = 0
                End If


                'Esta sentecia es para obtener el Codigo de Cliente
                strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"


                Dim oo As Object
                On Error GoTo AceptarErr
                Set oo = CreateObject("excel.application")
                oo.workbooks.Open vRuta & "\RptAvios.xlt"
                oo.Visible = True
                oo.run "Reporte", txtcoditem.Text, DevuelveCampo(strSQL, cCONNECT), txttemporada.Text, vericono, cCONNECT
                Screen.MousePointer = vbNormal
                oo.Visible = True
                Set oo = Nothing
                'MsgBox ("Aqui ira el reporte")
            Else
                Call MsgBox("Debe seleccionar un item para acceder a esta opcion", vbInformation)
            End If
         Case "GRAFICO"
            Load FrmGrafico
            FrmGrafico.diricono = txtDir_Icono.Text
            FrmGrafico.CARGA_ICONO
            FrmGrafico.Show 1

        Case "COMPOSICION"
            If Not rslista.EOF And Not rslista.BOF Then
                Load frmMantHilosItem
                frmMantHilosItem.Codigo_item = rslista("Cod_Item").Value
                frmMantHilosItem.txtDes_Item = rslista("Des_Item").Value
                frmMantHilosItem.CARGA_GRID
                frmMantHilosItem.Show 1
            Else
                MsgBox ("Debe seleccionar un Item para acceder a esta opcion")
            End If

        Case "COMBINACIONES"
            If cboIde_Color.Text = "S" Or CboIde_PO.Text = "S" Then Exit Sub
                If Not rslista.EOF And Not rslista.BOF Then
                    Load frmMantItemComb
                    frmMantItemComb.Caption = "COMBINACIONES DE ITEM:" & rslista("Cod_Item").Value & " " & rslista("Des_Item").Value
                    frmMantItemComb.Codigo_item = rslista("Cod_Item").Value
                    frmMantItemComb.txtDes_Item = rslista("Des_Item").Value
                    frmMantItemComb.CARGA_GRID
                    frmMantItemComb.Show 1
                Else
                    MsgBox ("Debe seleccionar un Item para acceder a esta opcion")
                End If
        Case "PROVEEDOR"
            If Not rslista.EOF And Not rslista.BOF Then
                Load frmManItemProv
                frmManItemProv.varCod_item = rslista("cod_item").Value
                frmManItemProv.varCod_Proveedor = rslista("Cod_Proveedor").Value
                frmManItemProv.Caption = "Item Proveedor  Item :" & rslista("cod_item").Value
                frmManItemProv.CARGA_GRID
                frmManItemProv.Show 1
            Else
                MsgBox ("Debe seleccionar un Item para acceder a esta opcion")
            End If

        Case "MEDIDA"
            If cboIde_Talla.Text = "S" Then Exit Sub
            
            FrmMantMed.Cod_Item = txtcoditem
            FrmMantMed.Tipo_Item = "I"
            FrmMantMed.Show 1
        Case "BITACORA"
           If Not rslista.EOF And Not rslista.BOF Then
            Load FrmBitacoraItems
            FrmBitacoraItems.Caption = "Bitacora Item: " & rslista("cod_item").Value
            FrmBitacoraItems.Cod_Item = rslista("cod_item").Value
            FrmBitacoraItems.CARGA_GRID
            FrmBitacoraItems.Show 1
            Set FrmBitacoraItems = Nothing
          End If
    End Select
Exit Sub
AceptarErr:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPETIQUETA"
On Error GoTo dprDepurar
    
Dim oPrint As clsPrintFile, lvEspacio As Integer, i As Integer
Dim lvLinea0 As String, lvLinea1 As String, lvLinea2 As String, lvLinea3 As String, _
    lvCodPro As String, lvDescripcion As String, lvUbicacion As String, lvLinea4 As String, lvCantidad As Integer, lvOpcional As String
    
If Trim(txtcoditem) <> "" Then
  lvLinea0 = ""
  lvLinea1 = ""
  lvLinea2 = ""
  lvLinea3 = ""
  lvLinea4 = ""
  
  lvCantidad = CInt(InputBox("Ingrese la Cantidad de Etiqueta a Imprimir :", "ETIQUETAS"))
  
  If lvCantidad = 0 Then Exit Sub
  
  lvOpcional = txtDesItem
  lvCodPro = "CODIGO  : " & txtcoditem
  lvUbicacion = "UBICACION: " & txtComentario
  
  If Len(lvOpcional) >= 25 Then
    lvDescripcion = Left(lvOpcional, 25)
    lvOpcional = Mid(lvOpcional, 26, Len(lvOpcional))
  Else
    lvDescripcion = lvOpcional
    lvOpcional = lvUbicacion
    lvUbicacion = ""
  End If
  
  lvEspacio = 34
  
  Open "c:\Etiquetas.txt" For Output As #1

  Plin lvLinea0
  
  For i = 1 To lvCantidad
    lvLinea1 = lvLinea1 & lvCodPro & Space(CInt(lvEspacio - Len(lvCodPro)))
    lvLinea2 = lvLinea2 & lvDescripcion & Space(CInt(lvEspacio - Len(lvDescripcion)))
    lvLinea3 = lvLinea3 & lvOpcional & Space(CInt(lvEspacio - Len(lvOpcional)))
    lvLinea4 = lvLinea4 & lvUbicacion & IIf(lvUbicacion <> "", Space(CInt(lvEspacio - Len(lvUbicacion))), "")

      Plin lvLinea1
      Plin lvLinea2
      Plin lvLinea3
      Plin lvLinea4
      Plin lvLinea0
      Plin lvLinea0
      Plin lvLinea0
      Plin lvLinea0
  
      lvLinea1 = ""
      lvLinea2 = ""
      lvLinea3 = ""
      lvLinea4 = ""
  
  Next i
  
  If lvLinea1 <> "" Then
      Plin lvLinea1
      Plin lvLinea2
      Plin lvLinea3
      Plin lvLinea4
      Plin lvLinea0
  End If
  
  Close #1
  Set oPrint = New clsPrintFile
  oPrint.SendPrint "c:\Etiquetas.txt"
  
  Set oPrint = Nothing
  
End If

Exit Sub
dprDepurar:
  ErrorHandler Err, "Imprime Etiqueta"
  Close #1
  Set oPrint = Nothing
Case "CAMBIO"
    Fra_cambio.Visible = True
    txt_Contable = ""
    txt_Contable.SetFocus

End Select
End Sub

Sub Plin(ByVal Text)
If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
End Sub




Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "UBICACION"
         ImprimirUbicacion
    End Select
End Sub


Private Sub ImprimirUbicacion()
Dim strCadena As String, rs As New ADODB.Recordset, strSQL As String
Dim oPrint As LibraryVB.clsPrintFile
Dim mRs As ADODB.Recordset
Dim iLinDet  As Integer
On Error GoTo ErrorHandler
Dim x As Integer
Dim Z As Double

Z = 0
Set oPrint = New clsPrintFile

Open "c:\workarea\FichaUbicacion.txt" For Output As #1


  Plin Chr(20) & "   "
  strCadena = "   " & "Codigo : " & Trim(txtcoditem.Text)
  Plin strCadena
  strCadena = "   " & "Descripcion : "
  Plin strCadena
  strCadena = "   " & Trim(txtDesItem.Text)
  Plin strCadena
  strCadena = "   " & "Ubicacion : " & Trim(txtComentario.Text)
  Plin strCadena
  
  Close #1
  
oPrint.SendPrint "c:\workarea\FichaUbicacion.txt"
Set oPrint = Nothing

'Unload Me
Exit Sub
Resume
ErrorHandler:
    Close #1
    ErrorHandler Err, "Impresion Factura"
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Eliminar As Integer
Dim strSQL As String

    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            FraBuscar.Enabled = False
            LIMPIAR_DATOS
            HABILITA_DATOS
            txtcoditem.Enabled = False
            txtDesItem.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            varCod_item = ""
        Case "MODIFICAR"
            sTipo = "U"
            FraBuscar.Enabled = False
            HABILITA_DATOS
            txtcoditem.Enabled = False
            cboCod_FamItem.Enabled = False
            txtDesItem.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            varCod_item = rslista("Cod_item").Value
                
            strSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
            If Trim(cboCod_FamItem.Text) <> "" Then
            If DevuelveCampo(strSQL, cCONNECT) = "M" Then
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
                'varCod_item = rslista("Cod_item").Value
                SALVAR_DATOS
                CargaDatos
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                'fraoptions.Enabled = False
                FraBuscar.Enabled = True
                If optItem.Value Then
                    txtcod_item.Text = varCod_item
                End If
                Call CargaLista
                'If sTipo = "U" Then
                    Call BuscaCampo(rslista, "Cod_Item", varCod_item)
                'End If
                sTipo = ""
                'Call BuscaCampo(rslista, "Cod_Item", Trim(txtcoditem.Text))

            End If
        Case "DESHACER"
            INHABILITA_DATOS
            sTipo = ""
            LIMPIAR_DATOS
            CargaDatos
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            'fraoptions.Enabled = False
            FraBuscar.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Private Sub optcliente_Click()
                    
    txtcliente.Text = ""
    txtNom_Cliente.Text = ""
    txttemporada.Text = ""
    txtNom_TemCli.Text = ""
    Frafamilia.Visible = False
    Fraitem.Visible = False
    Fracliente.Visible = True

    Opcion = 3
    
    'HabilitaMant Me.FunctButt1, "TEMPORADA/IMPRIMIR/LISTADO"
    Call CargaLista
    'txtcliente.SetFocus
End Sub

Private Sub optfamitem_Click()
    txtfamilia.Text = ""
    txtdes_famitem.Text = ""
    txtgrupo.Text = ""
    txtdes_famgruite.Text = ""
    
    txtgrupo.Enabled = False
    cmdBusgrupo.Enabled = False
    
    Frafamilia.Visible = True
    Fraitem.Visible = False
    Fracliente.Visible = False
    
    Opcion = 1
       
    'HabilitaMant Me.FunctButt1, "TEMPORADA/LISTADO"
    
    Call CargaLista
    txtfamilia.SetFocus
End Sub

Private Sub optitem_Click()
    txtcod_item.Text = ""
    txtDes_Item.Text = ""

    Frafamilia.Visible = False
    Fraitem.Visible = True
    Fracliente.Visible = False

    Opcion = 2
    
    'HabilitaMant Me.FunctButt1, "TEMPORADA/LISTADO"
    Call CargaLista
    txtcod_item.SetFocus
End Sub

Private Sub txt_Contable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmd_Aceptar.SetFocus
End If
End Sub

Private Sub txtCan_LotPed_KeyPress(KeyAscii As Integer)
    SoloNumeros txtCan_LotPed, KeyAscii, False, 0, 4
End Sub

Private Sub txtCan_LotPed_LostFocus()
    If Trim(txtCan_LotPed.Text) = "" Then
        txtCan_LotPed.Text = 0
    End If
End Sub

Private Sub txtCan_PtoReor_KeyPress(KeyAscii As Integer)
    SoloNumeros txtCan_PtoReor, KeyAscii, False, 0, 4
End Sub

Private Sub txtCan_PtoReor_LostFocus()
    If Trim(txtCan_PtoReor.Text) = "" Then
        txtCan_PtoReor.Text = 0
    End If
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            txttemporada.Enabled = True
            txtNom_TemCli.Enabled = True
            txttemporada.SetFocus
        End If
    End If
End Sub

'Private Sub txtcliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Trim(txtcliente.Text) = "" Then
'            cmdBusCliente_Click
'        Else
'            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
'            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
'            txttemporada.Enabled = True
'            txtNom_TemCli.Enabled = True
'            txttemporada.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtdes_famgruite_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        If Len(Trim(txtdes_famgruite.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        Else
            strSQL = "SELECT Cod_Gruitem FROM LG_FamGruIte WHERE Cod_Famitem='" & Trim(txtfamilia.Text) & "' AND des_famgruite LIKE '" & Trim(txtgrupo.Text) & "%'"
            txtgrupo.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
    End If

End Sub

Private Sub txtdes_famitem_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        txtgrupo.Text = ""
        If Len(Trim(txtdes_famitem.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        Else
            strSQL = "SELECT Cod_famitem FROM LG_FamIte WHERE Des_famitem LIKE '" & txtdes_famitem.Text & "%'"
            txtfamilia.Text = DevuelveCampo(strSQL, cCONNECT)
            
            txtgrupo.Enabled = True
            cmdBusgrupo.Enabled = True
            txtgrupo.SetFocus
        End If
    End If
End Sub

Private Sub txtdes_item_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        If Trim(txtDes_Item.Text) = "" Then
             'Call MsgBox("Sirvase ingresar una Descripcion del Item", vbInformation)
             Call MUESTRA_ITEMS(3)
        Else
            'Esta consulta es para obtener el Codigo de Cliente
            Call MUESTRA_ITEMS(2)
            
            'strSQL = "SELECT Cod_Item FROM LG_ITEM WHERE Des_Item LIKE '" & Trim(txtdes_item.Text) & "%'"
            'txtcod_item.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
        Call CargaLista
    End If
End Sub

Private Sub TxtFacConv_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtFacConv, KeyAscii, True, 3, 6)
End Sub

Private Sub txtMerma_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtMerma, KeyAscii, True, 2, 3)
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtcliente.Text = DevuelveCampo(strSQL, cCONNECT)
        Else
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        End If
    End If
End Sub

Private Sub txtNom_TemCli_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    'Esta consulta es para obtener el Codigo de Cliente
    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    If KeyAscii = 13 Then
        If Len(txtNom_TemCli) > 4 Then
            'Esta consulta nos permite obtener el Matching entre Cliente y Temporada
            strSQL = "SELECT Cod_TemCli FROM TG_TEMCLI WHERE Cod_Cliente='" & DevuelveCampo(strSQL, cCONNECT) & "' AND Nom_TemCli LIKE '" & Trim(txtNom_TemCli.Text) & "%'"
            txttemporada.Text = DevuelveCampo(strSQL, cCONNECT)
        Else
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        End If
    End If
End Sub

Private Sub txtRep_PreDol_KeyPress(KeyAscii As Integer)
    SoloNumeros txtRep_PreDol, KeyAscii, True, 6, 5
End Sub

Private Sub txtRep_PreDol_LostFocus()
    If Trim(txtRep_PreDol.Text) = "" Then
        txtRep_PreDol.Text = 0
    End If
End Sub

Private Sub txttemporada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txttemporada.Text) = "" Then
            cmdBusTemporada_Click
        Else
            strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
            strSQL = "SELECT Nom_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(strSQL, cCONNECT) & "' AND Cod_TemCli='" & txttemporada.Text & "'"
            txtNom_TemCli.Text = DevuelveCampo(strSQL, cCONNECT)
                       
            FunctBuscar.SetFocus
        End If
    End If
End Sub

Private Sub txtcod_item_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(txtcod_item.Text) = "" Then
            Call MUESTRA_ITEMS(1)
            'Call MsgBox("Sirvase ingresar un codigo de Item", vbInformation)
        Else
            txtcod_item.Text = CompletaCodigo(Trim(txtcod_item.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            strSQL = "SELECT Des_Item FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
            txtDes_Item.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
        Call CargaLista
    End If
End Sub

Private Sub txtfamilia_KeyPress(KeyAscii As Integer)
    Dim Opcion As Integer
    Dim strSQL As String
    If KeyAscii = 13 Then
        txtgrupo.Text = ""
        If Trim(txtfamilia.Text) = "" Then
            cmdBusFamItem_Click
        Else
            If ValidaFamilia = False Then
                 Exit Sub
            Else
                strSQL = "SELECT Des_famitem FROM LG_FamIte WHERE Cod_famitem='" & txtfamilia.Text & "'"
                txtdes_famitem.Text = DevuelveCampo(strSQL, cCONNECT)
                
                txtgrupo.Enabled = True
                cmdBusgrupo.Enabled = True
                txtgrupo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtgrupo_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    If KeyAscii = 13 Then
        If Trim(txtgrupo.Text) = "" Then
            cmdBusgrupo_Click
        Else
            If ValidaGrupo = False Then
                 Exit Sub
            Else
                strSQL = "SELECT  Cod_Gruitem as Código, des_famgruite as Descripción FROM LG_FamGruIte WHERE Cod_Famitem='" & Trim(txtfamilia.Text) & "' AND Cod_Gruitem='" & Trim(txtgrupo.Text) & "'"
                txtdes_famgruite = DevuelveCampo(strSQL, cCONNECT)
            
                FunctBuscar.SetFocus
            End If
        End If
    End If
End Sub


Private Sub CargaLista()
    Dim strSQL As String
    Set rslista = New ADODB.Recordset
    rslista.ActiveConnection = cCONNECT
    rslista.CursorType = adOpenStatic
    rslista.CursorLocation = adUseClient
    rslista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    strSQL = "EXEC UP_SEL_ITEMS " & Opcion & ",'" & txtfamilia.Text & "','" & Right(txtgrupo.Text, 4) & "','" & txtcod_item.Text & "','" & DevuelveCampo(strSQL, cCONNECT) & "','" & txttemporada.Text & "'"
    'strSQL = "EXEC UP_SEL_ITEMS_PRUEBA " & Opcion & ",'" & txtfamilia.Text & "','" & Right(txtgrupo.Text, 4) & "','" & txtcod_item.Text & "','" & DevuelveCampo(strSQL, cCONNECT) & "','" & txttemporada.Text & "'"
    
    rslista.Open strSQL
    Set DGridLista.DataSource = rslista

    If rslista.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
    Else
        DGridLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
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
Public Function ValidaFamilia() As Boolean
    Dim rs As New ADODB.Recordset
    Dim opcmessage As Integer
    rs.ActiveConnection = cCONNECT
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockReadOnly
    rs.Open "SELECT cod_famitem as Codigo, des_famitem as Descripcion FROM LG_FamIte WHERE cod_famitem='" & Trim(txtfamilia.Text) & "'"
    If rs.EOF Then
        opcmessage = MsgBox("La familia ingresada no existe, Desea Crearla?", vbInformation + vbYesNo)
        If opcmessage = vbYes Then
            'MsgBox ("Supuestamente llamamos al otro form")
            Load frmMantFamItem
            'frmMantFamItem.txtCod_FamItem.Text = txtfamilia.Text
            frmMantFamItem.Show 1
        Else
        ValidaFamilia = False
        End If
    Else
        ValidaFamilia = True
    End If
    Set rs = Nothing
End Function

Public Function ValidaGrupo() As Boolean
    Dim rs As New ADODB.Recordset
    Dim opcmessage As Integer
    rs.ActiveConnection = cCONNECT
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockReadOnly
    rs.Open "SELECT des_famgruite, Cod_Gruitem FROM LG_FamGruIte WHERE Cod_Famitem='" & Trim(txtfamilia.Text) & "' AND Cod_GruItem='" & Trim(txtgrupo.Text) & "'"
    If rs.EOF Then
        opcmessage = MsgBox("El Grupo ingresado no existe, Desea Crearlo?", vbInformation + vbYesNo)
        If opcmessage = vbYes Then
            Load frmMantFamGruItem
            frmMantFamGruItem.txtCod_FamItem = txtfamilia.Text
            frmMantFamGruItem.Cargar_Datos
            frmMantFamGruItem.Show 1
            'MsgBox ("Supuestamente llamamos al otro form")
        Else
        ValidaGrupo = False
        End If
    Else
        ValidaGrupo = True
    End If
    Set rs = Nothing
End Function

Public Sub CargaCombos()
    Dim strSQL As String
    
    'Combo Familia Item
    strSQL = "SELECT des_famitem + space(100) + cod_famitem  FROM LG_FamIte"
    Call LlenaCombo(cboCod_FamItem, strSQL, cCONNECT)
    
    'Combo Grupo Item
    'Strsql = "SELECT  Cod_Gruitem as Código, des_famgruite as Descripción FROM LG_FamGruIte WHERE Cod_Famitem='" & cboCod_FamItem.Text & "'"
    'Call LlenaCombo(cboCod_GruItem, Strsql, cCONNECT)
    
    'Combo Unidad de Medida
    strSQL = "SELECT Des_UniMed + space(100) + Cod_UniMed  FROM TG_UniMed"
    Call LlenaCombo(cboCod_UniMed, strSQL, cCONNECT)
    
    'Combo Clase de Item
    strSQL = "SELECT des_claitem + space(100) + cod_claitem  FROM LG_Claitem"
    Call LlenaCombo(cboCod_ClaItem, strSQL, cCONNECT)
    
    'Combo Flag Estatus
    strSQL = "SELECT des_status + space(100) + flg_status  FROM TG_StaDes"
    'Strsql = "SELECT cod_famitem as Codigo, des_famitem as Descripcion FROM LG_FamIte"
    Call LlenaCombo(cboFlg_Status, strSQL, cCONNECT)
    
    'Combo Origen
    strSQL = "SELECT des_origen + space(100) + cod_origen  FROM LG_Origen"
    Call LlenaCombo(cboCod_Origen, strSQL, cCONNECT)
    
    'Combo Motivo Preproduccion
    strSQL = "SELECT des_motprepro + space(100) + cod_motprepro  FROM TG_MotPrePro"
    Call LlenaCombo(cboCod_MotPrePro, strSQL, cCONNECT)
    
    strSQL = "select des_tipcar + space(100) + cod_tipcar from lg_tipcar"
    Call LlenaCombo(CboTipCar, strSQL, cCONNECT)
    
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
    
End Sub

Public Sub CargaDatos()
Dim strSQL As String
    If Not rslista.EOF Then
    
        txtcoditem.Text = Trim(rslista("Cod_Item"))
        txtDesItem.Text = Trim(rslista("Des_Item"))
        txtcta_cont.Text = Trim(rslista("Cod_CtaCont"))
        txtCan_PtoReor.Text = Trim(rslista("Can_PtoReor"))
        txtCan_LotPed.Text = Trim(rslista("Can_LotPed"))
        txtRep_PreDol.Text = Trim(rslista("Rep_PreDol"))
        txtDir_Icono.Text = Trim(rslista("Dir_Icono"))
        
        If IsNull(rslista("Comentario")) Then
            txtComentario.Text = ""
        Else
            txtComentario.Text = Trim(rslista("Comentario"))
        End If
        
        Call BuscaCombo(rslista("Cod_FamItem"), 2, cboCod_FamItem)
        Call BuscaCombo(rslista("Cod_GruItem"), 2, cboCod_GruItem)
        Call BuscaCombo(rslista("Cod_UniMed"), 2, cboCod_UniMed)
        Call BuscaCombo(rslista("Cod_ClaItem"), 2, cboCod_ClaItem)
        Call BuscaCombo(rslista("Flg_Status"), 2, cboFlg_Status)
        Call BuscaCombo(rslista("Cod_Origen"), 2, cboCod_Origen)
        Call BuscaCombo(rslista("Cod_MotPrePro"), 2, cboCod_MotPrePro)
        Call BuscaCombo(rslista("Ide_Talla"), 2, cboIde_Talla)
        
        Call BuscaCombo(rslista("Ide_Color"), 2, cboIde_Color)
        Call BuscaCombo(rslista("Ide_EsCli"), 2, cboIde_EsCli)
        
        
        If Not IsNull(rslista("Ide_Destino")) Then
            Call BuscaCombo(rslista("Ide_Destino"), 2, cboIde_Destino)
        Else
            cboIde_Destino.ListIndex = -1
        End If
        
        If Not IsNull(rslista("Ide_Po")) Then
            Call BuscaCombo(rslista("Ide_Po"), 2, CboIde_PO)
        Else
            CboIde_PO.ListIndex = -1
        End If
        
        If Not IsNull(rslista("Cod_tipcar")) Then
            Call BuscaCombo(rslista("Cod_tipcar"), 2, CboTipCar)
        Else
            CboTipCar.ListIndex = -1
        End If
        TxtFacConv.Text = IIf(IsNull(rslista("Fac_Conversion")), 0, rslista("Fac_conversion"))
        txtMerma.Text = IIf(IsNull(rslista("Por_mertin")), 0, rslista("Por_mertin"))
        
        If Not IsNull(rslista("cod_hiltel")) And Trim(rslista("cod_hiltel")) <> "" Then
            vCod_hiltel = Trim(rslista("cod_hiltel"))
            'strSQL = "select des_hiltel from it_hilado where cod_hiltel='" & rslista("cod_hiltel") & "'"
            'TxtHilado.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "select cod_hilado_estructurado from it_hilado where cod_hiltel='" & rslista("cod_hiltel") & "'"
            TxtHilado.Text = Trim(rslista("cod_hiltel"))
            Call BuscaCombo(DevuelveCampo(strSQL, cCONNECT), 1, cboCod_HilTel)
        Else
            TxtHilado.Text = ""
            cboCod_HilTel.ListIndex = -1
        End If
        
        
        
        LblPrecio = Trim(rslista("precio"))
        LblProveedor = Trim(rslista("proveedor"))
        LblMoneda = Trim(rslista("Moneda"))
        
        If Not IsNull(rslista("Ult_compra")) Then
            DTPUltCompra.Value = rslista("Ult_compra")
        End If
        
        LblOrdCompra = Trim(rslista("O_Compra"))
    End If
End Sub

Public Sub LIMPIAR_DATOS()

    txtcoditem.Text = ""
    txtDesItem.Text = ""
    txtcta_cont.Text = ""
    txtCan_PtoReor.Text = "0"
    txtCan_LotPed.Text = "0"
    txtRep_PreDol.Text = "0"
    txtDir_Icono.Text = ""
    txtComentario.Text = ""
    
    'Limpiamos el Grupo
    cboCod_GruItem.Clear  '.ListIndex = -1
    
    If Opcion = 1 And Trim(txtfamilia.Text) <> "" And sTipo = "I" Then
        Call BuscaCombo(txtfamilia.Text, 2, cboCod_FamItem)
    Else
        cboCod_FamItem.ListIndex = -1
    End If
    
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
    
    CboTipCar.ListIndex = -1
    TxtFacConv.Text = "0"
    txtMerma.Text = "0"
    TxtHilado.Text = ""
    'sTipo = ""
End Sub

Public Sub HABILITA_DATOS()

    txtcoditem.Enabled = True
    txtDesItem.Enabled = True
    txtcta_cont.Enabled = False
    txtCan_PtoReor.Enabled = True
    txtCan_LotPed.Enabled = True
    txtRep_PreDol.Enabled = True
    txtDir_Icono.Enabled = True
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
    cmdDirIcono.Enabled = True

End Sub

Public Sub INHABILITA_DATOS()

    txtcoditem.Enabled = False
    txtDesItem.Enabled = False
    txtcta_cont.Enabled = False
    txtCan_PtoReor.Enabled = False
    txtCan_LotPed.Enabled = False
    txtRep_PreDol.Enabled = False
    txtDir_Icono.Enabled = False
    txtComentario.Enabled = False
    
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
    cmdDirIcono.Enabled = False
    HABILITA_CARACMXT False
    
End Sub

Public Function VALIDA_DATOS() As Boolean
Dim strSQL As String
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
        strSQL = "select cod_tipfam from LG_FamIte where Cod_Famitem='" & Right(cboCod_FamItem.Text, 2) & "'"
        If DevuelveCampo(strSQL, cCONNECT) = "M" Then
            
            'Call HABILITA_CARACMXT(True)
            If FrameMixtas.Enabled = True Then
                If Trim(CboTipCar.Text) = "" Then
                    Call MsgBox("La característica no puede estar vacio. Sirvase verificar", vbCritical)
                    CboTipCar.SetFocus
                    VALIDA_DATOS = False
                    Exit Function
                End If
                If Trim(TxtFacConv.Text) = "" Or Trim(TxtFacConv.Text) = "0" Then
                    Call MsgBox("El Factor de Conversion no puede ser cero o estar vacio. Sirvase verificar", vbCritical)
                    TxtFacConv.SetFocus
                    VALIDA_DATOS = False
                    Exit Function
                End If
                If Trim(txtMerma.Text) = "" Or Trim(txtMerma.Text) = "0" Then
                    Call MsgBox("La Merma no puede ser cero o estar vacio. Sirvase verificar", vbCritical)
                    txtMerma.SetFocus
                    VALIDA_DATOS = False
                    Exit Function
                End If
                If Trim(Right(CboTipCar.Text, 2)) = "T" Or Trim(Right(CboTipCar.Text, 2)) = "C" Then
                    If vCod_hiltel = "" Then
                        Call MsgBox("El Hilado no puede estar vacio. Sirvase verificar", vbCritical)
                        txtMerma.SetFocus
                        VALIDA_DATOS = False
                        Exit Function
                    End If
                Else
                    vCod_hiltel = ""
                End If
            End If
        End If
    End If
        
    
End Function

Public Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    Con.ConnectionString = cCONNECT
    Con.Open
    
    rs.ActiveConnection = cCONNECT
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Con.BeginTrans
       
        'Esta sentecia es para obtener el Codigo de Cliente
        strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
       
        'Esta es la sentencia que realizara el salvado de datos
        strSQL = "UP_MAN_ITEMS " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        txtcoditem.Text & "','" & _
        Right(cboCod_FamItem.Text, 2) & "','" & _
        Right(cboCod_GruItem.Text, 4) & "','" & _
        Right(cboCod_UniMed.Text, 2) & "','" & _
        txtcta_cont.Text & "','" & _
        txtDesItem.Text & "','" & _
        Right(cboCod_ClaItem.Text, 2) & "'," & _
        txtCan_PtoReor.Text & "," & _
        txtCan_LotPed.Text & "," & _
        txtRep_PreDol.Text & ",'" & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Right(cboCod_Origen.Text, 2) & "','" & _
        cboIde_Talla.Text & "','" & _
        cboIde_Color.Text & "','" & _
        cboIde_EsCli.Text & "','" & _
        cboIde_Destino.Text & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        DevuelveCampo(strSQL, cCONNECT) & "','" & _
        txttemporada.Text & "','" & Trim(txtComentario.Text) & "','" & _
        txtDir_Icono.Text & "','" & _
        txtMerma.Text & "','" & vCod_hiltel & "','" & _
        Trim(Right(CboTipCar.Text, 2)) & "','" & TxtFacConv.Text & "','" & CboIde_PO.Text & "','" & vusu & "'"
          
        If sTipo = "I" Then
            Set rs = Con.Execute(strSQL)
            If rs.RecordCount Then
                varCod_item = rs(0)
            End If
            Set rs = Nothing
        Else
            Con.Execute strSQL
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
    Dim strSQL As String
    
    strSQL = "SELECT COD_CLIENTE FROM LG_ITEMTEMCLI WHERE Cod_Item='" & txtcoditem.Text & "'"

    If DevuelveCampo(strSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
           
        'Esta sentecia es para obtener el Codigo de Cliente
        strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    
        'Esta es la sentencia que realiza la eliminacion del Registro
        strSQL = "UP_MAN_ITEMS " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        txtcoditem.Text & "','" & _
        Right(cboCod_FamItem.Text, 2) & "','" & _
        Right(cboCod_GruItem.Text, 4) & "','" & _
        Right(cboCod_UniMed.Text, 2) & "','" & _
        txtcta_cont.Text & "','" & _
        txtDesItem.Text & "','" & _
        Right(cboCod_ClaItem.Text, 2) & "'," & _
        txtCan_PtoReor.Text & "," & _
        txtCan_LotPed.Text & "," & _
        txtRep_PreDol.Text & ",'" & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Right(cboCod_Origen.Text, 2) & "','" & _
        cboIde_Talla.Text & "','" & _
        cboIde_Color.Text & "','" & _
        cboIde_EsCli.Text & "','" & _
        cboIde_Destino.Text & "','" & _
        Right(cboCod_MotPrePro.Text, 2) & "','" & _
        DevuelveCampo(strSQL, cCONNECT) & "','" & _
        txttemporada.Text & "','" & Trim(txtComentario.Text) & "','" & _
        txtDir_Icono.Text & "','" & _
        txtMerma.Text & "','" & vCod_hiltel & "','" & _
        Trim(Right(CboTipCar.Text, 2)) & "','" & TxtFacConv.Text & "','" & CboIde_PO.Text & "','" & vusu & "'"
                
        Con.Execute strSQL
    
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
    FrameMixtas.Enabled = vEstado
    CboTipCar.Enabled = vEstado
    TxtFacConv.Enabled = vEstado
    txtMerma.Enabled = vEstado
End Sub

Sub MUESTRA_ITEMS(tipo As Integer)
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    If tipo = 1 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY cod_Item"
    ElseIf tipo = 2 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item where des_item like '%" & Trim(Me.txtDes_Item.Text) & "%' ORDER BY des_Item"
    ElseIf tipo = 3 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY Des_Item"
    End If
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcod_item.Text = Codigo
        txtDes_Item.Text = Descripcion
        
        txtgrupo.Enabled = True
        cmdBusgrupo.Enabled = True
        FunctBuscar.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Sub CAMBIO()
Dim Sql As String
On Error GoTo hand
   
   Sql = "up_lg_item_cambio_cuenta '" & txtcoditem & "','" & txt_Contable & "'"
   
   Call ExecuteSQL(cCONNECT, Sql)
   
    txt_Contable = ""
Exit Sub
hand:
    
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
    'ErrorHandler Err, "ASIGNAR_MAQUINA"
End Sub



