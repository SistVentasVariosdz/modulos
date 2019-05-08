VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmMantTelaComb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combinaciones"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   2055
      Left            =   7920
      TabIndex        =   41
      Top             =   4560
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   3625
      Custom          =   $"frmMantTelaComb.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   470
      ControlSeparator=   50
   End
   Begin FunctionsButtons.FunctButt FunctDetalles 
      Height          =   1020
      Left            =   7920
      TabIndex        =   19
      Top             =   3480
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1799
      Custom          =   $"frmMantTelaComb.frx":0166
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   470
      ControlSeparator=   50
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   870
      TabIndex        =   14
      Top             =   8625
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1410
         Picture         =   "frmMantTelaComb.frx":01FF
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   -15
         Picture         =   "frmMantTelaComb.frx":0371
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   930
         Picture         =   "frmMantTelaComb.frx":04E3
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   435
         Picture         =   "frmMantTelaComb.frx":0655
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
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
      Height          =   3270
      Left            =   90
      TabIndex        =   12
      Tag             =   "List"
      Top             =   0
      Width           =   9135
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   75
         TabIndex        =   13
         Top             =   225
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   5159
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   18
         BeginProperty Column00 
            DataField       =   "Cod_Comb"
            Caption         =   "Código"
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
            DataField       =   "Des_Comb"
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
            DataField       =   "Cod_Tipo_Desarrollo"
            Caption         =   "Cod_Tipo_Desarrollo"
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
         BeginProperty Column03 
            DataField       =   "Des_Tipo_Desarrollo"
            Caption         =   "Tipo Desarrollo"
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
         BeginProperty Column04 
            DataField       =   "Gramaje"
            Caption         =   "Gramaje"
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
            DataField       =   "Ancho"
            Caption         =   "Ancho"
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
            DataField       =   "Cod_Tela_Produccion"
            Caption         =   "Cod_Tela_Produccion"
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
            DataField       =   "Des_Tela_Produccion"
            Caption         =   "Des_Tela_Produccion"
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
            DataField       =   "Long_Malla"
            Caption         =   "Long_Malla"
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
         BeginProperty Column09 
            DataField       =   "Diametro"
            Caption         =   "Diametro"
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
         BeginProperty Column10 
            DataField       =   "Cod_Galga"
            Caption         =   "Cod_Galga"
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
         BeginProperty Column11 
            DataField       =   "Des_Galga"
            Caption         =   "Des_Galga"
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
         BeginProperty Column12 
            DataField       =   "Articulo"
            Caption         =   "Articulo Creado"
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
         BeginProperty Column13 
            DataField       =   "Rapport_Number"
            Caption         =   "Rapport_Number"
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
         BeginProperty Column14 
            DataField       =   "Rapport_Comb"
            Caption         =   "Rapport_Comb"
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
         BeginProperty Column15 
            DataField       =   "Tratamiento_en_Humedo"
            Caption         =   "Tratamiento en Humedo"
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
         BeginProperty Column16 
            DataField       =   "Tratamiento_en_Acabados"
            Caption         =   "Tratamiento en Acabados"
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
         BeginProperty Column17 
            DataField       =   "Tratamiento_Previo"
            Caption         =   "Tratamiento Previo"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4034.835
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   4034.835
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   4034.835
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   4034.835
            EndProperty
         EndProperty
      End
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
      Height          =   5400
      Left            =   120
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3240
      Width           =   7695
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   2240
         MaxLength       =   100
         TabIndex        =   60
         Top             =   4560
         Width           =   5145
      End
      Begin VB.TextBox txtRapport_Comb 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4960
         Width           =   1125
      End
      Begin VB.TextBox txtRapport_Number 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4960
         Width           =   1125
      End
      Begin VB.Frame fraGnrl 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   105
         TabIndex        =   28
         Top             =   200
         Width           =   7350
         Begin VB.TextBox txtDes_Tela 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4515
            TabIndex        =   31
            Top             =   100
            Width           =   2055
         End
         Begin VB.TextBox txtDes_Comb 
            Height          =   285
            Left            =   1260
            TabIndex        =   30
            Top             =   480
            Width           =   6015
         End
         Begin VB.TextBox txtCod_Comb 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   29
            Top             =   100
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   15
            TabIndex        =   34
            Top             =   195
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tela"
            Height          =   195
            Left            =   3675
            TabIndex        =   33
            Top             =   160
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   0
            TabIndex        =   32
            Top             =   525
            Width           =   930
         End
      End
      Begin VB.Frame fraDE 
         BorderStyle     =   0  'None
         Height          =   3675
         Left            =   75
         TabIndex        =   20
         Top             =   900
         Width           =   7410
         Begin VB.OptionButton OptPulgadas 
            Caption         =   "Pulg."
            Height          =   195
            Left            =   1245
            TabIndex        =   55
            Top             =   1800
            Width           =   735
         End
         Begin VB.OptionButton OptCentimetros 
            Caption         =   "Ctms."
            Height          =   195
            Left            =   1965
            TabIndex        =   54
            Top             =   1800
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtColumnas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4080
            TabIndex        =   52
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtCunSas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   53
            Text            =   "0"
            Top             =   1800
            Width           =   825
         End
         Begin VB.TextBox txtTrataPrevio 
            Height          =   300
            Left            =   2160
            TabIndex        =   51
            Top             =   2640
            Width           =   5145
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   285
            Left            =   3420
            TabIndex        =   48
            Top             =   2280
            Visible         =   0   'False
            Width           =   3885
         End
         Begin VB.TextBox txtcliente 
            Height          =   285
            Left            =   2175
            MaxLength       =   5
            TabIndex        =   47
            Top             =   2280
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   330
            Left            =   3045
            TabIndex        =   46
            Tag             =   "..."
            Top             =   2280
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox txtTrataAcabado 
            Height          =   285
            Left            =   2160
            TabIndex        =   45
            Top             =   3360
            Width           =   5145
         End
         Begin VB.TextBox txtTrataHumedo 
            Height          =   285
            Left            =   2160
            TabIndex        =   44
            Top             =   3000
            Width           =   5145
         End
         Begin VB.TextBox TxtGram_Despues 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   39
            Text            =   "0"
            Top             =   1440
            Width           =   825
         End
         Begin VB.TextBox txtDiametro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   10
            Text            =   "0"
            Top             =   1080
            Width           =   810
         End
         Begin VB.TextBox txtMalla 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3555
            TabIndex        =   9
            Text            =   "0"
            Top             =   1095
            Width           =   795
         End
         Begin VB.TextBox txtDes_Galga 
            Height          =   285
            Left            =   2040
            TabIndex        =   6
            Top             =   780
            Width           =   2310
         End
         Begin VB.TextBox txtCod_Galga 
            Height          =   285
            Left            =   1275
            TabIndex        =   5
            Top             =   780
            Width           =   750
         End
         Begin VB.TextBox txtDes_TelaProd 
            Height          =   285
            Left            =   2580
            TabIndex        =   4
            Top             =   465
            Width           =   4710
         End
         Begin VB.TextBox txtCod_TelaProd 
            Height          =   285
            Left            =   1275
            TabIndex        =   3
            Top             =   465
            Width           =   1290
         End
         Begin VB.TextBox txtCod_TipoDesarrollo 
            Height          =   285
            Left            =   1275
            TabIndex        =   1
            Top             =   120
            Width           =   750
         End
         Begin VB.TextBox txtDes_TipoDesarrollo 
            Height          =   285
            Left            =   2040
            TabIndex        =   2
            Top             =   120
            Width           =   5250
         End
         Begin VB.TextBox txtGramaje 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   7
            Text            =   "0"
            Top             =   780
            Width           =   825
         End
         Begin VB.TextBox txtAncho 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1275
            TabIndex        =   8
            Text            =   "0"
            Top             =   1095
            Width           =   810
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Medida:"
            Height          =   195
            Left            =   0
            TabIndex        =   58
            Top             =   1800
            Width           =   930
         End
         Begin VB.Label Label58 
            Caption         =   "Columnas :"
            Height          =   255
            Left            =   3240
            TabIndex        =   57
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label TxtxCunSas 
            Caption         =   "CunSas :"
            Height          =   255
            Left            =   5520
            TabIndex        =   56
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Tratamiento Previo"
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   2640
            Width           =   1995
         End
         Begin VB.Label Label16 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2280
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "Tratamiento en Acabados"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   3360
            Width           =   1845
         End
         Begin VB.Label Label14 
            Caption         =   "Tratamiento en Humedo"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   3000
            Width           =   1860
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Gram. Despues Lavado:"
            Height          =   195
            Left            =   4440
            TabIndex        =   40
            Top             =   1500
            Width           =   1725
         End
         Begin VB.Label Label10 
            Caption         =   "Diametro :"
            Height          =   165
            Left            =   4980
            TabIndex        =   27
            Top             =   1170
            Width           =   720
         End
         Begin VB.Label Label9 
            Caption         =   "Long.Malla :"
            Height          =   195
            Left            =   2595
            TabIndex        =   26
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label Label8 
            Caption         =   "Galga :"
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label7 
            Caption         =   "Tela Produccion :"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   525
            Width           =   1260
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Desarrollo :"
            Height          =   210
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Gramaje Antes Lavado:"
            Height          =   195
            Left            =   4440
            TabIndex        =   22
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label Label6 
            Caption         =   "Ancho :"
            Height          =   180
            Left            =   0
            TabIndex        =   21
            Top             =   1140
            Width           =   615
         End
      End
      Begin VB.Label Label18 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Rapport Comb:"
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   5040
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Rapport Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   5040
         Width           =   1215
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3120
      TabIndex        =   11
      Top             =   8760
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantTelaComb.frx":07C7
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantTelaComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Grid As New ADODB.Recordset
Dim strSQL As String
Public Codigo_tela As String, sCod_FamTela As String, sCod_Cliente As String, sCod_Temcli As String

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente order by 1"
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


Private Sub cmdFirst_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MovePrevious
    End If
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    HabilitaMant Me.MantFunc1, ""
    DESHABILITA_DATOS
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    FunctDetalles.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    OptPulgadas.Value = True
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Grid.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Grid.EOF And Not Rs_Grid.BOF Then
        Call Carga_Datos
        DESHABILITA_DATOS
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Grid = Nothing
End Sub





Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "POSTENIDO"
    If Not Rs_Grid.EOF Then
        If UCase(sCod_FamTela) <> "DE" Then
            MsgBox "Ingreso solo permitido para desarrollos", vbCritical
            Exit Sub
        End If
        Load FrmManTelasDatTec
        FrmManTelasDatTec.sCod_Tela = Codigo_tela
        FrmManTelasDatTec.sDes_tela = Trim(txtDes_Tela.Text)
        FrmManTelasDatTec.sFamite = sCod_FamTela
        FrmManTelasDatTec.sCod_Comb = Rs_Grid("Cod_Comb")
        FrmManTelasDatTec.sDes_Comb = Trim(Rs_Grid("Des_Comb"))
        FrmManTelasDatTec.Carga_Datos
        FrmManTelasDatTec.Show 1
        Set FrmManTelasDatTec = Nothing
    End If
Case "PRUEBA"
    If Not Rs_Grid.EOF Then
        If UCase(sCod_FamTela) <> "DE" Then
            MsgBox "Ingreso solo permitido para desarrollos", vbCritical
            Exit Sub
        End If
        Load FrmManTelasDatTecAdd
        FrmManTelasDatTecAdd.sCod_Tela = Codigo_tela
        FrmManTelasDatTecAdd.sFamite = sCod_FamTela
        FrmManTelasDatTecAdd.sCod_Comb = Rs_Grid("Cod_Comb")
        FrmManTelasDatTecAdd.Carga_Datos
        FrmManTelasDatTecAdd.Show 1
        Set FrmManTelasDatTecAdd = Nothing
    End If
Case "HOJA"
    If Not Rs_Grid.EOF Then
        Call Hoja_Ruta
    End If
Case "COPIAR"
    If Not Rs_Grid.EOF Then
        Load FrmCopiarComb
        FrmCopiarComb.Codigo_tela = Codigo_tela
        FrmCopiarComb.vCodCombD = Rs_Grid("Cod_Comb")
        FrmCopiarComb.Show 1
        Set FrmCopiarComb = Nothing
    End If
End Select
End Sub

Private Sub FunctDetalles_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "DETALLES"
            If Not Rs_Grid.EOF Then
                Load frmMantTelaCombDet
                frmMantTelaCombDet.Caption = "DETALLE DE COMBINACION:" & Rs_Grid("Cod_Comb") & " " & Rs_Grid("Des_Comb")
                frmMantTelaCombDet.Codigo_tela = Rs_Grid("Cod_Tela")
                frmMantTelaCombDet.Codigo_Comb = Rs_Grid("Cod_Comb")
                frmMantTelaCombDet.sCod_Cliente = sCod_Cliente
                frmMantTelaCombDet.sCod_Temcli = sCod_Temcli
                'frmMantTelaCombDet.txtDes_Tela.Text = txtDes_Tela.Text
                'frmMantTelaCombDet.txtDes_Comb.Text = Rs_Grid("Des_Comb")
                frmMantTelaCombDet.CargaCombos
                frmMantTelaCombDet.rapport_number = DGridLista.Columns("RAPPORT_NUMBER").Value
                frmMantTelaCombDet.Rapport_Comb = DGridLista.Columns("RAPPORT_COMB").Value
                frmMantTelaCombDet.Cod_Famtela = Me.sCod_FamTela
                If UCase(Me.sCod_FamTela) = "DE" Then
                    frmMantTelaCombDet.FraDesarrollo.Visible = True
                Else
                    frmMantTelaCombDet.FraDesarrollo.Visible = False
                End If
                
                frmMantTelaCombDet.CARGA_GRID
                frmMantTelaCombDet.Show 1
            Else
                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
            End If
        Case "PROCESOS"
            If Not Rs_Grid.EOF Then
                Load FrmManteProcesosTelasComb
                FrmManteProcesosTelasComb.Caption = "PROCESOS COMBINACION:" & Rs_Grid("Cod_Comb") & " " & Rs_Grid("Des_Comb")
                FrmManteProcesosTelasComb.Codigo_tela = Rs_Grid("Cod_Tela")
                FrmManteProcesosTelasComb.Codigo_Comb = Rs_Grid("Cod_Comb")
                FrmManteProcesosTelasComb.LblTela = Rs_Grid("Cod_Tela") & " - " & txtDes_Tela.Text
                FrmManteProcesosTelasComb.LblComb = Rs_Grid("Cod_Comb") & " - " & Rs_Grid("Des_Comb")
                FrmManteProcesosTelasComb.Carga_Datos
                FrmManteProcesosTelasComb.Show 1
            Else
                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
            End If
    End Select
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            If txtCod_Comb.Enabled Then
                txtCod_Comb.SetFocus
            Else
                txtDes_Comb.SetFocus
            End If
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            If FixNulos(DGridLista.Columns("Rapport_Number"), vbLong) <> 0 Then
                MsgBox "Combinación sólo puede ser modificada desde panatlla R.N.", vbCritical, "No se permite modificación"
                Exit Sub
            End If
            
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            If FixNulos(DGridLista.Columns("Rapport_Number"), vbLong) <> 0 Then
                MsgBox "Combinación sólo puede ser eliminada desde panatlla R.N.", vbCritical, "No se permite Eliminación"
                Exit Sub
            End If
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("Esta seguro de eliminar el registro", vbInformation + vbYesNo)
                If Eliminar = vbYes Then
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                sTipo = ""
            End If
        Case "DESHACER"
        
            DESHABILITA_DATOS
            LIMPIAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Sub LIMPIAR_DATOS()
    txtCod_Comb.Text = ""
    txtDes_Comb.Text = ""
    txtCod_TipoDesarrollo.Text = ""
    txtDes_TipoDesarrollo.Text = ""
    txtGramaje.Text = 0
    txtAncho.Text = 0
    txtCod_TelaProd.Text = ""
    txtDes_TelaProd.Text = ""
    txtCod_Galga.Text = ""
    txtDes_Galga.Text = ""
    txtMalla.Text = 0
    txtDiametro.Text = 0
    txtRapport_Comb.Text = ""
    txtRapport_Number = ""
    TxtGram_Despues.Text = 0
    
    Me.txtColumnas.Text = "0"
    Me.txtCunSas = "0"
    Me.txtObservaciones = ""
    
    End Sub


Sub Carga_Datos()

    If Not Rs_Grid.EOF Then
        txtCod_Comb.Text = Trim(Rs_Grid("Cod_Comb").Value)
        txtDes_Comb.Text = Trim(Rs_Grid("Des_Comb").Value)
        txtCod_TipoDesarrollo.Text = Trim(Rs_Grid("Cod_Tipo_Desarrollo").Value)
        txtDes_TipoDesarrollo.Text = Trim(Rs_Grid("Des_Tipo_Desarrollo").Value)
        txtGramaje.Text = Rs_Grid("gramaje").Value
        txtAncho.Text = Rs_Grid("ancho").Value
        txtCod_TelaProd.Text = Rs_Grid("cod_tela_produccion").Value
        txtDes_TelaProd.Text = Rs_Grid("des_tela_produccion").Value
        txtCod_Galga.Text = Rs_Grid("cod_galga").Value
        txtDes_Galga.Text = Rs_Grid("des_galga").Value
        txtMalla.Text = Rs_Grid("long_malla").Value
        txtDiametro.Text = Rs_Grid("diametro").Value
        txtRapport_Number = FixNulos(Rs_Grid("Rapport_Number").Value, vbLong)
        txtRapport_Comb = FixNulos(Rs_Grid("Rapport_Comb").Value, vbString)
        TxtGram_Despues.Text = Rs_Grid("gramaje_despues_lavado").Value
        txtTrataHumedo.Text = Trim(Rs_Grid("Tratamiento_en_Humedo").Value)
        txtTrataAcabado.Text = Trim(Rs_Grid("Tratamiento_en_Acabados").Value)
        txtTrataPrevio.Text = Trim(Rs_Grid("Tratamiento_previo").Value)
        txtObservaciones.Text = Trim(Rs_Grid("Observaciones").Value)
        
        If Rs_Grid("Tipo_Medida") = "P" Then
            Me.OptPulgadas = True
        Else
            Me.OptCentimetros = True
        End If
        
        Me.txtColumnas = Rs_Grid("Num_Columnas")
        Me.txtCunSas = Rs_Grid("Num_CunSas")
    
    End If

End Sub

Sub RECARGAR_DATOS()
    
    Rs_Grid.Close
    CARGA_GRID
    
End Sub

Public Sub CARGA_GRID()
    Dim strSQL As String
    Set Rs_Grid = New ADODB.Recordset
    Rs_Grid.ActiveConnection = cCONNECT
    Rs_Grid.CursorType = adOpenStatic
    Rs_Grid.CursorLocation = adUseClient
    Rs_Grid.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC UP_SEL_TELACOMB '" & Codigo_tela & "'"
    
    Rs_Grid.Open strSQL
    Set DGridLista.DataSource = Rs_Grid
    DGridLista.Refresh

    If Rs_Grid.RecordCount > 0 Then
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
        If Trim(txtDes_Comb.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    Else
        strSQL = "SELECT COUNT(Num_Secuencia) FROM TX_TELACOMBDET WHERE Cod_Tela='" & Rs_Grid("Cod_Tela").Value & "' AND Cod_Comb='" & Rs_Grid("Cod_Comb").Value & "'"
        If DevuelveCampo(strSQL, cCONNECT) > 0 Then
            Call MsgBox("No se puede eliminar este Regitro por que posee registros relacionados", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    txtCod_Comb = Trim(txtCod_Comb)
    If txtCod_Comb <> "" Then
        'txtCod_Comb = Format(txtCod_Comb, "000")
        txtCod_Comb = UCase(txtCod_Comb)
'        If Not IsNumeric(txtCod_Comb) Then
'            MsgBox "El Codigo de Combinacion debe ser un número", _
'            vbExclamation + vbOKOnly, "Datos de Combinación"
'            VALIDA_DATOS = False
'            Exit Function
'        End If
        If Len(txtCod_Comb) <> txtCod_Comb.MaxLength Then
            MsgBox "El Codigo de Combinacion debe tener tres dígitos", _
            vbExclamation + vbOKOnly, "Datos de Combinación"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub HABILITA_DATOS()
    txtCod_Comb.Enabled = (Left(Codigo_tela, 2) = "DE" And sTipo = "I")
    txtDes_Comb.Enabled = True
    txtCod_TipoDesarrollo.Enabled = True
    txtDes_TipoDesarrollo.Enabled = True
    txtGramaje.Enabled = True
    txtAncho.Enabled = True
    txtCod_TelaProd.Enabled = True
    txtDes_TelaProd.Enabled = True
    txtCod_Galga.Enabled = True
    txtDes_Galga.Enabled = True
    txtMalla.Enabled = True
    txtDiametro.Enabled = True
    TxtGram_Despues.Enabled = True
    txtTrataHumedo.Enabled = True
    txtTrataAcabado.Enabled = True
      
    txtTrataPrevio.Enabled = True
    
    txtcliente.Enabled = True
    txtNom_Cliente.Enabled = True
    cmdBusCliente.Enabled = True
    
    
    Me.OptPulgadas.Enabled = True
    Me.OptCentimetros.Enabled = True
    Me.txtColumnas.Enabled = True
    Me.txtCunSas.Enabled = True
    Me.txtObservaciones.Enabled = True
    
End Sub

Sub DESHABILITA_DATOS()
    txtCod_Comb.Enabled = False
    txtDes_Comb.Enabled = False
    txtCod_TipoDesarrollo.Enabled = False
    txtDes_TipoDesarrollo.Enabled = False
    txtGramaje.Enabled = False
    txtAncho.Enabled = False
    txtCod_TelaProd.Enabled = False
    txtDes_TelaProd.Enabled = False
    txtCod_Galga.Enabled = False
    txtDes_Galga.Enabled = False
    txtMalla.Enabled = False
    txtDiametro.Enabled = False
    TxtGram_Despues.Enabled = False
    txtTrataHumedo.Enabled = False
    txtTrataAcabado.Enabled = False
    txtTrataPrevio.Enabled = False
    
    txtcliente.Enabled = False
    txtNom_Cliente.Enabled = False
    cmdBusCliente.Enabled = False
    
    Me.txtColumnas.Enabled = False
    Me.txtCunSas.Enabled = False
    
    Me.OptPulgadas.Enabled = False
    Me.OptCentimetros.Enabled = False
    Me.txtObservaciones.Enabled = False
    
End Sub

Sub SALVAR_DATOS()
'Dim Con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Dim strSQL As String
    
'    Con.ConnectionString = cCONNECT
'    Con.Open
'
'    Con.BeginTrans
    
    strSQL = "EXEC UP_MAN_TELACOMB '" & _
    sTipo & "','" & _
    Codigo_tela & "','" & _
    txtCod_Comb.Text & "','" & _
    txtDes_Comb.Text & "','" & _
    txtCod_TipoDesarrollo.Text & "'," & _
    txtGramaje.Text & "," & _
    txtAncho.Text & ",'" & _
    txtCod_TelaProd.Text & "','" & _
    txtCod_Galga.Text & "'," & _
    txtMalla.Text & "," & _
    txtDiametro & ",'" & _
    vusu & "','S','" & TxtGram_Despues.Text & "','" & _
    Trim(txtTrataHumedo.Text) & "','" & Trim(txtTrataAcabado.Text) & "','" & Trim(txtTrataPrevio.Text) & "','" & _
    IIf(OptCentimetros, "C", "P") & "'," & _
    Me.txtColumnas & "," & _
    Me.txtCunSas & ",'" & _
    Me.txtObservaciones & "'"
    
'    Con.Execute StrSQL
'
'    Con.CommitTrans
    Call ExecuteSQL(cCONNECT, strSQL)
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
'    Con.RollbackTrans
'    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
       
        strSQL = "EXEC UP_MAN_TELACOMB '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        txtCod_Comb.Text & "','" & _
        txtDes_Comb.Text & "','" & _
        txtCod_TipoDesarrollo.Text & "'," & _
        txtGramaje.Text & "," & _
        txtAncho.Text & ",'" & _
        txtCod_TelaProd.Text & "','" & _
        txtCod_Galga.Text & "'," & _
        txtMalla.Text & "," & _
        txtDiametro & ",'" & _
        vusu & "'"
        
        Con.Execute strSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub TxtAncho_GotFocus()
    SelectionText txtAncho
End Sub

Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMalla.SetFocus
    Else
        Call SoloNumeros(txtAncho, KeyAscii, True, 2)
    End If
End Sub

Private Sub txtAncho_LostFocus()
    If Trim(txtAncho.Text) = "" Then txtAncho.Text = 0
End Sub

Private Sub txtCod_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'txtCod_Comb = Format(txtCod_Comb, "000")
        txtCod_Comb = UCase(txtCod_Comb)
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Galga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Galga.Text) = "" Then
            BUSCA_GALGA (3)
        Else
            BUSCA_GALGA (1)
        End If
    End If
End Sub

Private Sub txtCod_TelaProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtCod_TelaProd.Text)) < 3 Then
            If Trim(txtCod_TelaProd.Text) = "" Then
                txtDes_TelaProd.Text = ""
                txtDes_TelaProd.SetFocus
            Else
                BUSCA_TELA (3)
            End If
        Else
            BUSCA_TELA (1)
        End If
    End If
End Sub

Private Sub txtCod_TipoDesarrollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_TipoDesarrollo.Text) = "" Then
            BUSCA_TIPO (3)
        Else
            BUSCA_TIPO (1)
        End If
    End If
End Sub



Private Sub txtColumnas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(Me.txtColumnas, KeyAscii, False, 2)
End If
End Sub

Private Sub txtCunSas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.MantFunc1.SetFocus
Else
    Call SoloNumeros(Me.txtCunSas, KeyAscii, False, 2)
End If
End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fraDE.Visible Then
            txtCod_TipoDesarrollo.SetFocus
        Else
            'Me.MantFunc1.SetFocus
            Me.txtObservaciones.SetFocus
        End If
    End If
End Sub

Private Sub txtDes_Galga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Galga.Text) = "" Then
            BUSCA_GALGA (3)
        Else
            BUSCA_GALGA (2)
        End If
    End If
End Sub

Private Sub txtDes_TelaProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_TelaProd.Text) = "" Then
            txtCod_Galga.SetFocus
        Else
            BUSCA_TELA (2)
        End If
    End If
End Sub

Private Sub txtDes_TipoDesarrollo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_TipoDesarrollo.Text) = "" Then
            BUSCA_TIPO (3)
        Else
            BUSCA_TIPO (2)
        End If
    End If
End Sub

Private Sub txtDiametro_GotFocus()
    SelectionText txtDiametro
End Sub

Private Sub txtdiametro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtGram_Despues.SetFocus
    Else
        Call SoloNumeros(txtDiametro, KeyAscii, False, 0)
    End If
End Sub

Private Sub txtdiametro_LostFocus()
    If Trim(txtDiametro.Text) = "" Then txtDiametro.Text = 0
End Sub

Private Sub TxtGram_Despues_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtColumnas.SetFocus
Else
    Call SoloNumeros(TxtGram_Despues, KeyAscii, False, 0)
End If
End Sub

Private Sub TxtGramaje_GotFocus()
    SelectionText txtGramaje
End Sub

Private Sub TxtGramaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAncho.SetFocus
    Else
        Call SoloNumeros(txtGramaje, KeyAscii, False, 0)
    End If
End Sub

Private Sub txtGramaje_LostFocus()
    If Trim(txtGramaje.TabIndex) = "" Then txtGramaje.Text = 0
End Sub


Public Sub BUSCA_TIPO(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "SELECT Des_Tipo_Desarrollo FROM TX_TIPOS_DESARROLLO WHERE Cod_Tipo_Desarrollo = '" & Trim(Me.txtCod_TipoDesarrollo.Text) & "'"
                    Me.txtDes_TipoDesarrollo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    Me.txtCod_TelaProd.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_Tipo_Desarrollo AS 'Código', Des_Tipo_Desarrollo AS 'Descripción' FROM TX_TIPOS_DESARROLLO where Des_Tipo_Desarrollo like '%" & Trim(txtDes_TipoDesarrollo.Text) & "%' order by Cod_Tipo_Desarrollo"
                    Else
                        oTipo.sQuery = "SELECT Cod_Tipo_Desarrollo AS 'Código', Des_Tipo_Desarrollo AS 'Descripción' FROM TX_TIPOS_DESARROLLO order by Cod_Tipo_Desarrollo"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_TipoDesarrollo.Text = Trim(Codigo)
                         Me.txtDes_TipoDesarrollo.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.txtCod_TelaProd.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Public Sub BUSCA_TELA(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(txtCod_TelaProd.Text) = "", 0, Mid(txtCod_TelaProd.Text, 3)) & ")", cCONNECT))
                    txtCod_TelaProd.Text = Left(txtCod_TelaProd, 2) & strSQL
                    strSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela = '" & Trim(Me.txtCod_TelaProd.Text) & "'"
                    Me.txtDes_TelaProd.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    Me.txtCod_Galga.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_Tela as Codigo,Des_Tela as Descripcion FROM TX_TELA WHERE Des_Tela like '%" & Trim(Me.txtDes_TelaProd.Text) & "%' order by cod_tela"
                    Else
                        oTipo.sQuery = "SELECT Cod_Tela as Codigo,Des_Tela as Descripcion FROM TX_TELA WHERE cod_tela like '%" & txtCod_TelaProd.Text & "%' Order by cod_tela"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_TelaProd.Text = Trim(Codigo)
                         Me.txtDes_TelaProd.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.txtCod_Galga.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Public Sub BUSCA_GALGA(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "SELECT Des_Galga FROM TX_GALGA WHERE Cod_Galga = '" & Trim(Me.txtCod_Galga.Text) & "'"
                    Me.txtDes_Galga.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    Me.txtGramaje.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_Galga as Codigo,Des_Galga as Descripcion FROM TX_GALGA WHERE Des_Galga like '%" & Trim(Me.txtDes_Galga.Text) & "%' order by cod_galga"
                    Else
                        oTipo.sQuery = "SELECT Cod_Galga as Codigo,Des_Galga as Descripcion FROM TX_GALGA WHERE cod_galga like '%" & txtCod_Galga.Text & "%' Order by cod_galga"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_Galga.Text = Trim(Codigo)
                         Me.txtDes_Galga.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.txtGramaje.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Private Sub txtMalla_GotFocus()
    SelectionText txtMalla
End Sub

Private Sub txtMalla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDiametro.SetFocus
    End If
End Sub

Private Sub txtMalla_LostFocus()
    If Trim(txtMalla.Text) = "" Then txtMalla.Text = 0
End Sub

Sub Hoja_Ruta()
On Error GoTo hand
    Dim oo As Object
    Dim strSQL As String
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\Hoja_TecnicaComb.xlt"
    oo.Visible = True
    oo.run "Reporte", Codigo_tela, cCONNECT, Trim(Rs_Grid("Cod_Comb").Value)
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "Telas No Operativas"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub


Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
        txtTrataPrevio.SetFocus
    End If
End Sub


Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtcliente.Text = DevuelveCampo(strSQL, cCONNECT)
            txtTrataPrevio.SetFocus
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
         
    End If
End Sub

 
Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Me.MantFunc1.SetFocus
 End If
End Sub

Private Sub txtTrataAcabado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Me.txtObservaciones.SetFocus
 End If
End Sub

Private Sub txtTrataHumedo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtTrataAcabado.SetFocus
    End If
 
End Sub

Private Sub txtTrataPrevio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtTrataHumedo.SetFocus
    End If
End Sub
