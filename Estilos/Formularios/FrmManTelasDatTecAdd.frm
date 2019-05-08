VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmManTelasDatTecAdd 
   Caption         =   "Datos Tecnicos - Pruebas"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   104
      Top             =   6960
      Width           =   8175
      Begin VB.TextBox TxtObservaciones_Relevantes 
         Height          =   405
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   240
         Width           =   5655
      End
      Begin VB.TextBox TxtObservaciones_Considerables 
         Height          =   405
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones Relevantes"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   360
         Width           =   1920
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones Considerables"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   840
         Width           =   2100
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2760
      TabIndex        =   53
      Top             =   8280
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmManTelasDatTecAdd.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame6 
      Caption         =   "Estandar Requerido Lavado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   93
      Top             =   4560
      Width           =   3615
      Begin VB.TextBox TxtLav_Revirado 
         Height          =   285
         Left            =   2160
         TabIndex        =   50
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox TxtLav_Ancho_Tela 
         Height          =   285
         Left            =   2160
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtLav_PesoBW 
         Height          =   285
         Left            =   2160
         TabIndex        =   46
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtLav_PesoAW 
         Height          =   285
         Left            =   2160
         TabIndex        =   47
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtLav_Encog_Ancho 
         Height          =   285
         Left            =   2160
         TabIndex        =   48
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtLav_Encog_Largo 
         Height          =   285
         Left            =   2160
         TabIndex        =   49
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   480
         TabIndex        =   99
         Top             =   2085
         Width           =   1170
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   480
         TabIndex        =   98
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   97
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   96
         Top             =   1185
         Width           =   1230
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   480
         TabIndex        =   95
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   480
         TabIndex        =   94
         Top             =   1785
         Width           =   1125
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Pruebas de %E Residuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   86
      Top             =   4560
      Width           =   4215
      Begin VB.TextBox TxtResi3_Revirado 
         Height          =   285
         Left            =   2880
         TabIndex        =   44
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox TxtResi3_Ancho_Tela 
         Height          =   285
         Left            =   2880
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtResi3_PesoBW 
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtResi3_PesoAW 
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtResi3_Encog_Ancho 
         Height          =   285
         Left            =   2880
         TabIndex        =   40
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtResi3_Encog_Largo 
         Height          =   285
         Left            =   2880
         TabIndex        =   42
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_Revirado 
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_Ancho_Tela 
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_PesoBW 
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_PesoAW 
         Height          =   285
         Left            =   1920
         TabIndex        =   37
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_Encog_Ancho 
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtResi1_Encog_Largo 
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "%E Tercera"
         Height          =   195
         Left            =   2880
         TabIndex        =   103
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "%E Primera"
         Height          =   195
         Left            =   1920
         TabIndex        =   102
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   360
         TabIndex        =   92
         Top             =   2085
         Width           =   1170
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   360
         TabIndex        =   91
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   360
         TabIndex        =   90
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   360
         TabIndex        =   89
         Top             =   1185
         Width           =   1230
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   360
         TabIndex        =   88
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   360
         TabIndex        =   87
         Top             =   1785
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Estandar Requerido Acabado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   79
      Top             =   2160
      Width           =   3615
      Begin VB.TextBox TxtReqAca_Encog_Largo 
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox TxtReqAca_Encog_Ancho 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtReqAca_PesoAW 
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtReqAca_Peso_BW 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtReqAca_Ancho_Tela 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtReqAca_Revirado 
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   480
         TabIndex        =   85
         Top             =   1785
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   480
         TabIndex        =   84
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   83
         Top             =   1185
         Width           =   1230
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   82
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   480
         TabIndex        =   81
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   480
         TabIndex        =   80
         Top             =   2085
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pruebas de %E Acabado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   72
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox TxtAca1_Encog_Largo 
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox TxtAca1_Encog_Ancho 
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtAca1_PesoAW 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtAca1_PesoBW 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtAca1_Ancho_Tela 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtAca1_Revirado 
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_Encog_Largo 
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   1695
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_Encog_Ancho 
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   1395
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_PesoAW 
         Height          =   285
         Left            =   2880
         TabIndex        =   20
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_PesoBW 
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_Ancho_Tela 
         Height          =   285
         Left            =   2880
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtAca3_Revirado 
         Height          =   285
         Left            =   2880
         TabIndex        =   26
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "%E Tercera"
         Height          =   195
         Left            =   2880
         TabIndex        =   101
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "%E Primera"
         Height          =   195
         Left            =   1920
         TabIndex        =   100
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   360
         TabIndex        =   78
         Top             =   1785
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   360
         TabIndex        =   77
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   360
         TabIndex        =   76
         Top             =   1185
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   360
         TabIndex        =   75
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   360
         TabIndex        =   74
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   360
         TabIndex        =   73
         Top             =   2085
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobación Desarrollo Textil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5160
      TabIndex        =   65
      Top             =   0
      Width           =   3015
      Begin VB.TextBox TxtPruBoil_Encog_Largo 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1455
         Width           =   855
      End
      Begin VB.TextBox TxtPruBoil_Encog_Ancho 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox TxtPruBoil_PesoAW 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   855
         Width           =   855
      End
      Begin VB.TextBox TxtPruBoil_PesoBW 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtPruBoil_Ancho_Tela 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPruBoil_Revirado 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   480
         TabIndex        =   71
         Top             =   1845
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   480
         TabIndex        =   70
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   69
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   68
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   480
         TabIndex        =   67
         Top             =   1245
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   480
         TabIndex        =   66
         Top             =   1545
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pruebas de %E Secado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2160
      TabIndex        =   58
      Top             =   0
      Width           =   3015
      Begin VB.TextBox TxtPruSeca_Encog_Largo 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1455
         Width           =   855
      End
      Begin VB.TextBox TxtPruSeca_Encog_Ancho 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox TxtPruSeca_PesoAW 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   855
         Width           =   855
      End
      Begin VB.TextBox TxtPruSeca_PesoBW 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtPruSeca_Ancho_Tela 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPruSeca_Revirado 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Largo %"
         Height          =   195
         Left            =   480
         TabIndex        =   64
         Top             =   1545
         Width           =   1125
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "Encog. Ancho %"
         Height          =   195
         Left            =   480
         TabIndex        =   63
         Top             =   1245
         Width           =   1185
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Peso AW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   62
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Peso BW (gr/m2)"
         Height          =   195
         Left            =   480
         TabIndex        =   61
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tela"
         Height          =   195
         Left            =   480
         TabIndex        =   60
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Revirado Tela %"
         Height          =   195
         Left            =   480
         TabIndex        =   59
         Top             =   1845
         Width           =   1170
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Test Tambler (Teñido)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   2175
      Begin VB.TextBox TxtTambler_Ancho_Lavado 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtTambler_Densidad 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   915
         Width           =   735
      End
      Begin VB.TextBox TxtTambler_Ancho_Proyectado 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Lavado"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Densidad"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   1005
         Width           =   675
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Proyect."
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1335
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmManTelasDatTecAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Public sCod_Tela As String, sCod_Comb As String, sFamite As String
Public sCod_Ruta As String
Public Codigo As String, Descripcion As String, TipoAdd As String

Private Sub Form_Load()
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "GRABAR"
    Call Grabar
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Carga_Datos()
Dim Rs As New ADODB.Recordset
On Error GoTo errDatos

If UCase(sFamite) <> "DE" Then
    strSQL = "Tx_Muestra_DATOS_TECNICOS_TELA_Nuevo '" & sCod_Tela & "','" & sCod_Ruta & "'"
Else
    strSQL = "Tx_Muestra_DATOS_TECNICOS_TELAcomb_Nuevo '" & sCod_Tela & "','" & sCod_Comb & "'"
End If

Set Rs = Nothing
Rs.CursorLocation = adUseClient

Rs.Open strSQL, cCONNECT, 3
If Rs.RecordCount Then
    Me.TxtPruSeca_Ancho_Tela = Rs!Prue_Seca_Ancho_Tela
    Me.TxtPruSeca_PesoBW = Rs!Prue_Seca_Peso_BW
    Me.TxtPruSeca_PesoAW = Rs!Prue_Seca_Peso_AW
    Me.TxtPruSeca_Encog_Ancho = Rs!Prue_Seca_Encog_Ancho
    Me.TxtPruSeca_Encog_Largo = Rs!Prue_Seca_Encog_Largo
    Me.TxtPruSeca_Revirado = Rs!Prue_Seca_Revirado
    Me.TxtPruBoil_Ancho_Tela = Rs!Prue_BoilOff_Ancho_Tela
    Me.TxtPruBoil_PesoBW = Rs!Prue_BoilOff_Peso_BW
    Me.TxtPruBoil_PesoAW = Rs!Prue_BoilOff_Peso_AW
    Me.TxtPruBoil_Encog_Ancho = Rs!Prue_BoilOff_Encog_Ancho
    Me.TxtPruBoil_Encog_Largo = Rs!Prue_BoilOff_Encog_Largo
    Me.TxtPruBoil_Revirado = Rs!Prue_BoilOff_Revirado
    Me.TxtObservaciones_Relevantes = Rs!Observaciones_Relevantes
    Me.TxtObservaciones_Considerables = Rs!Observaciones_Considerables
    Me.TxtTambler_Ancho_Lavado = Rs!Test_Tam_Ancho_Lavado
    Me.TxtTambler_Densidad = Rs!Test_Tam_Densidad
    Me.TxtTambler_Ancho_Proyectado = Rs!Test_Tam_Ancho_Proyecta
    
    Me.TxtAca1_Ancho_Tela = Rs!Prue_Aca1_Ancho_Tela
    Me.TxtAca1_PesoBW = FixNulos(Rs!Prue_Aca1_Peso_BW, vbDouble)
    Me.TxtAca1_PesoAW = FixNulos(Rs!Prue_Aca1_Peso_AW, vbDouble)
    Me.TxtAca1_Encog_Ancho = FixNulos(Rs!Prue_Aca1_Encog_Ancho, vbDouble)
    Me.TxtAca1_Encog_Largo = FixNulos(Rs!Prue_Aca1_Encog_Largo, vbDouble)
    Me.TxtAca1_Revirado = FixNulos(Rs!Prue_Aca1_Revirado, vbDouble)
    Me.TxtAca3_Ancho_Tela = FixNulos(Rs!Prue_Aca3_Ancho_Tela, vbDouble)
    Me.TxtAca3_PesoBW = FixNulos(Rs!Prue_Aca3_Peso_BW, vbDouble)
    Me.TxtAca3_PesoAW = FixNulos(Rs!Prue_Aca3_Peso_AW, vbDouble)
    Me.TxtAca3_Encog_Ancho = FixNulos(Rs!Prue_Aca3_Encog_Ancho, vbDouble)
    Me.TxtAca3_Encog_Largo = FixNulos(Rs!Prue_Aca3_Encog_Largo, vbDouble)
    Me.TxtAca3_Revirado = FixNulos(Rs!Prue_Aca3_Revirado, vbDouble)
    
    Me.TxtReqAca_Ancho_Tela = FixNulos(Rs!Prue_Esta_Acab_Ancho_Tela, vbDouble)
    Me.TxtReqAca_Peso_BW = FixNulos(Rs!Prue_Esta_Acab_Peso_BW, vbDouble)
    Me.TxtReqAca_PesoAW = FixNulos(Rs!Prue_Esta_Acab_Peso_AW, vbDouble)
    Me.TxtReqAca_Encog_Ancho = FixNulos(Rs!Prue_Esta_Acab_Encog_Ancho, vbDouble)
    Me.TxtReqAca_Encog_Largo = FixNulos(Rs!Prue_Esta_Acab_Encog_largo, vbDouble)
    Me.TxtReqAca_Revirado = FixNulos(Rs!Prue_Esta_Acab_Revirado, vbDouble)
    
    Me.TxtResi1_Ancho_Tela = Rs!Prue_Resi1_Ancho_Tela
    Me.TxtResi1_PesoBW = FixNulos(Rs!Prue_Resi1_Peso_BW, vbDouble)
    Me.TxtResi1_PesoAW = FixNulos(Rs!Prue_Resi1_Peso_AW, vbDouble)
    Me.TxtResi1_Encog_Ancho = FixNulos(Rs!Prue_Resi1_Encog_Ancho, vbDouble)
    Me.TxtResi1_Encog_Largo = FixNulos(Rs!Prue_Resi1_Encog_Largo, vbDouble)
    Me.TxtResi1_Revirado = FixNulos(Rs!Prue_Resi1_Revirado, vbDouble)
    Me.TxtResi3_Ancho_Tela = FixNulos(Rs!Prue_Resi3_Ancho_Tela, vbDouble)
    Me.TxtResi3_PesoBW = FixNulos(Rs!Prue_Resi3_Peso_BW, vbDouble)
    Me.TxtResi3_PesoAW = FixNulos(Rs!Prue_Resi3_Peso_AW, vbDouble)
    Me.TxtResi3_Encog_Ancho = FixNulos(Rs!Prue_Resi3_Encog_Ancho, vbDouble)
    Me.TxtResi3_Encog_Largo = FixNulos(Rs!Prue_Resi3_Encog_Largo, vbDouble)
    Me.TxtResi3_Revirado = FixNulos(Rs!Prue_Resi3_Revirado, vbDouble)
    
    Me.TxtLav_Ancho_Tela = FixNulos(Rs!Prue_Esta_Lav_Ancho_Tela, vbDouble)
    Me.TxtLav_PesoBW = FixNulos(Rs!Prue_Esta_Lav_Peso_BW, vbDouble)
    Me.TxtLav_PesoAW = FixNulos(Rs!Prue_Esta_Lav_Peso_AW, vbDouble)
    Me.TxtLav_Encog_Ancho = FixNulos(Rs!Prue_Esta_Lav_Encog_Ancho, vbDouble)
    Me.TxtLav_Encog_Largo = FixNulos(Rs!Prue_Esta_Lav_Encog_Largo, vbDouble)
    Me.TxtLav_Revirado = FixNulos(Rs!Prue_Esta_Lav_Revirado, vbDouble)
    
    
End If

Exit Sub
errDatos:
    ErrorHandler Err, "Carga Datos"
End Sub

Sub Grabar()
Dim vTipo_Lavado As String
On Error GoTo errGrabar

    If Trim(TxtLav_Ancho_Tela) = "" Then TxtLav_Ancho_Tela = "0"
    If Trim(TxtLav_PesoBW) = "" Then TxtLav_PesoBW = "0"
    If Trim(TxtLav_PesoAW) = "" Then TxtLav_PesoAW = "0"
    If Trim(TxtLav_Encog_Ancho) = "" Then TxtLav_Encog_Ancho = "0"
    If Trim(TxtLav_Encog_Largo) = "" Then TxtLav_Encog_Largo = "0"
    If Trim(TxtLav_Revirado) = "" Then TxtLav_Revirado = "0"
    If Trim(TxtAca1_PesoBW) = "" Then TxtAca1_PesoBW = "0"
        
    If Trim(Me.TxtPruSeca_Ancho_Tela) = "" Then TxtPruSeca_Ancho_Tela = "0"
    If Trim(Me.TxtPruSeca_PesoBW) = "" Then TxtPruSeca_PesoBW = "0"
    If Trim(Me.TxtPruSeca_PesoAW) = "" Then TxtPruSeca_PesoAW = "0"
    If Trim(Me.TxtPruSeca_Encog_Ancho) = "" Then TxtPruSeca_Encog_Ancho = "0"
    If Trim(Me.TxtPruSeca_Encog_Largo) = "" Then TxtPruSeca_Encog_Largo = "0"
    If Trim(Me.TxtPruSeca_Revirado) = "" Then TxtPruSeca_Revirado = "0"
    If Trim(Me.TxtPruBoil_Ancho_Tela) = "" Then TxtPruBoil_Ancho_Tela = "0"
    If Trim(Me.TxtPruBoil_PesoBW) = "" Then TxtPruBoil_PesoBW = "0"
    If Trim(Me.TxtPruBoil_PesoAW) = "" Then TxtPruBoil_PesoAW = "0"
    If Trim(Me.TxtPruBoil_Encog_Ancho) = "" Then TxtPruBoil_Encog_Ancho = "0"
    If Trim(Me.TxtPruBoil_Encog_Largo) = "" Then TxtPruBoil_Encog_Largo = "0"
    If Trim(Me.TxtPruBoil_Revirado) = "" Then TxtPruBoil_Revirado = "0"
    If Trim(Me.TxtTambler_Ancho_Lavado) = "" Then TxtTambler_Ancho_Lavado = "0"
    If Trim(Me.TxtTambler_Densidad) = "" Then TxtTambler_Densidad = "0"
    If Trim(Me.TxtTambler_Ancho_Proyectado) = "" Then TxtTambler_Ancho_Proyectado = "0"
    
    
    If Trim(Me.TxtAca1_Ancho_Tela) = "" Then TxtAca1_Ancho_Tela = "0"
    If Trim(Me.TxtAca1_PesoBW) = "" Then TxtAca1_Ancho_Tela = "0"
    If Trim(Me.TxtAca1_PesoAW) = "" Then TxtAca1_PesoAW = "0"
    If Trim(Me.TxtAca1_Encog_Ancho) = "" Then TxtAca1_Encog_Ancho = "0"
    If Trim(Me.TxtAca1_Encog_Largo) = "" Then TxtAca1_Encog_Largo = "0"
    If Trim(Me.TxtAca1_Revirado) = "" Then TxtAca1_Revirado = "0"
    If Trim(Me.TxtAca3_Ancho_Tela) = "" Then TxtAca3_Ancho_Tela = "0"
    If Trim(Me.TxtAca3_PesoBW) = "" Then TxtAca3_PesoBW = "0"
    If Trim(Me.TxtAca3_PesoAW) = "" Then TxtAca3_PesoAW = "0"
    If Trim(Me.TxtAca3_Encog_Ancho) = "" Then TxtAca3_Encog_Ancho = "0"
    If Trim(Me.TxtAca3_Encog_Largo) = "" Then TxtAca3_Encog_Largo = "0"
    If Trim(Me.TxtAca3_Revirado) = "" Then TxtAca3_Revirado = "0"
    
    If Trim(Me.TxtReqAca_Ancho_Tela) = "" Then TxtReqAca_Ancho_Tela = "0"
    If Trim(Me.TxtReqAca_Peso_BW) = "" Then TxtReqAca_Peso_BW = "0"
    If Trim(Me.TxtReqAca_PesoAW) = "" Then TxtReqAca_PesoAW = "0"
    If Trim(Me.TxtReqAca_Encog_Ancho) = "" Then TxtReqAca_Encog_Ancho = "0"
    If Trim(Me.TxtReqAca_Encog_Largo) = "" Then TxtReqAca_Encog_Largo = "0"
    If Trim(Me.TxtReqAca_Revirado) = "" Then TxtReqAca_Revirado = "0"
    If Trim(Me.TxtResi1_Ancho_Tela) = "" Then TxtResi1_Ancho_Tela = "0"
    If Trim(Me.TxtResi1_PesoBW) = "" Then TxtResi1_PesoBW = "0"
    If Trim(Me.TxtResi1_PesoAW) = "" Then TxtResi1_PesoAW = "0"
    If Trim(Me.TxtResi1_Encog_Ancho) = "" Then TxtResi1_Encog_Ancho = "0"
    If Trim(Me.TxtResi1_Encog_Largo) = "" Then TxtResi1_Encog_Largo = "0"
    If Trim(Me.TxtResi1_Revirado) = "" Then TxtResi1_Revirado = "0"
    If Trim(Me.TxtResi3_Ancho_Tela) = "" Then TxtResi3_Ancho_Tela = "0"
    If Trim(Me.TxtResi3_PesoBW) = "" Then TxtResi3_PesoBW = "0"
    If Trim(Me.TxtResi3_PesoAW) = "" Then TxtResi3_PesoAW = "0"
    If Trim(Me.TxtResi3_Encog_Ancho) = "" Then TxtResi3_Encog_Ancho = "0"
    If Trim(Me.TxtResi3_Encog_Largo) = "" Then TxtResi3_Encog_Largo = "0"
    If Trim(Me.TxtResi3_Revirado) = "" Then TxtResi3_Revirado = "0"

If sFamite = "DE" Then
    strSQL = "UP_TX_ACTUALIZA_DATOS_TECNICOS_TELAcomb_NUEVOS_ADD '" & _
            sCod_Tela & "','" & sCod_Comb & "','" & TxtPruSeca_Ancho_Tela & "','" & _
            TxtPruSeca_PesoBW & "','" & TxtPruSeca_PesoAW & "','" & TxtPruSeca_Encog_Ancho & "','" & _
            TxtPruSeca_Encog_Largo & "','" & TxtPruSeca_Revirado & "','" & TxtPruBoil_Ancho_Tela & "','" & TxtPruBoil_PesoBW & "','" & _
            TxtPruBoil_PesoAW & "','" & TxtPruBoil_Encog_Ancho & "','" & TxtPruBoil_Encog_Largo & "','" & TxtPruBoil_Revirado & "','" & TxtObservaciones_Relevantes & "','" & TxtObservaciones_Considerables & "','" & _
            TxtTambler_Ancho_Lavado & "','" & TxtTambler_Densidad & "','" & TxtTambler_Ancho_Proyectado & "','" & _
            TxtAca1_Ancho_Tela & "','" & TxtAca1_PesoBW & "','" & TxtAca1_PesoAW & "','" & TxtAca1_Encog_Ancho & "','" & _
            TxtAca1_Encog_Largo & "','" & TxtAca1_Revirado & "','" & TxtAca3_Ancho_Tela & "','" & TxtAca3_PesoBW & "','" & _
            TxtAca3_PesoAW & "','" & TxtAca3_Encog_Ancho & "','" & TxtAca3_Encog_Largo & "','" & TxtAca3_Revirado & "','" & _
            TxtReqAca_Ancho_Tela & "','" & TxtReqAca_Peso_BW & "','" & TxtReqAca_PesoAW & "','" & TxtReqAca_Encog_Ancho & "','" & _
            TxtReqAca_Encog_Largo & "','" & TxtReqAca_Revirado & "','" & TxtResi1_Ancho_Tela & "','" & _
            TxtResi1_PesoBW & "','" & TxtResi1_PesoAW & "','" & TxtResi1_Encog_Ancho & "','" & TxtResi1_Encog_Largo & "','" & TxtResi1_Revirado & "','" & _
            TxtResi3_Ancho_Tela & "','" & TxtResi3_PesoBW & "','" & TxtResi3_PesoAW & "','" & TxtResi3_Encog_Ancho & "','" & TxtResi3_Encog_Largo & "','" & _
            TxtResi3_Revirado & "','" & TxtLav_Ancho_Tela & "','" & TxtLav_PesoBW & "','" & TxtLav_PesoAW & "','" & TxtLav_Encog_Ancho & "','" & _
            TxtLav_Encog_Largo & "','" & TxtLav_Revirado & "','" & vusu & "','" & ComputerName & "'"
Else
    strSQL = "UP_TX_ACTUALIZA_DATOS_TECNICOS_TELA_NUEVOS_ADD '" & _
            sCod_Tela & "','" & TxtPruSeca_Ancho_Tela & "','" & _
            TxtPruSeca_PesoBW & "','" & TxtPruSeca_PesoAW & "','" & TxtPruSeca_Encog_Ancho & "','" & _
            TxtPruSeca_Encog_Largo & "','" & TxtPruSeca_Revirado & "','" & TxtPruBoil_Ancho_Tela & "','" & TxtPruBoil_PesoBW & "','" & _
            TxtPruBoil_PesoAW & "','" & TxtPruBoil_Encog_Ancho & "','" & TxtPruBoil_Encog_Largo & "','" & TxtPruBoil_Revirado & "','" & TxtObservaciones_Relevantes & "','" & TxtObservaciones_Considerables & "','" & _
            TxtTambler_Ancho_Lavado & "','" & TxtTambler_Densidad & "','" & TxtTambler_Ancho_Proyectado & "','" & _
            TxtAca1_Ancho_Tela & "','" & TxtAca1_PesoBW & "','" & TxtAca1_PesoAW & "','" & TxtAca1_Encog_Ancho & "','" & _
            TxtAca1_Encog_Largo & "','" & TxtAca1_Revirado & "','" & TxtAca3_Ancho_Tela & "','" & TxtAca3_PesoBW & "','" & _
            TxtAca3_PesoAW & "','" & TxtAca3_Encog_Ancho & "','" & TxtAca3_Encog_Largo & "','" & TxtAca3_Revirado & "','" & _
            TxtReqAca_Ancho_Tela & "','" & TxtReqAca_Peso_BW & "','" & TxtReqAca_PesoAW & "','" & TxtReqAca_Encog_Ancho & "','" & _
            TxtReqAca_Encog_Largo & "','" & TxtReqAca_Revirado & "','" & TxtResi1_Ancho_Tela & "','" & _
            TxtResi1_PesoBW & "','" & TxtResi1_PesoAW & "','" & TxtResi1_Encog_Ancho & "','" & TxtResi1_Encog_Largo & "','" & TxtResi1_Revirado & "','" & _
            TxtResi3_Ancho_Tela & "','" & TxtResi3_PesoBW & "','" & TxtResi3_PesoAW & "','" & TxtResi3_Encog_Ancho & "','" & TxtResi3_Encog_Largo & "','" & _
            TxtResi3_Revirado & "','" & TxtLav_Ancho_Tela & "','" & TxtLav_PesoBW & "','" & TxtLav_PesoAW & "','" & TxtLav_Encog_Ancho & "','" & _
            TxtLav_Encog_Largo & "','" & TxtLav_Revirado & "','" & vusu & "','" & ComputerName & "','" & sCod_Ruta & "'"
End If
            
ExecuteCommandSQL cCONNECT, strSQL
MsgBox "Se grabó correctamente", vbInformation
Unload Me

Exit Sub
errGrabar:
    ErrorHandler Err, "Grabar"
End Sub

Private Sub TxtAca1_Ancho_Tela_GotFocus()
SelectionText TxtAca1_Ancho_Tela
End Sub

Private Sub TxtAca1_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca1_Encog_Ancho_GotFocus()
SelectionText TxtAca1_Encog_Ancho
End Sub

Private Sub TxtAca1_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca1_Encog_Largo_GotFocus()
SelectionText TxtAca1_Encog_Largo
End Sub

Private Sub TxtAca1_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca1_PesoAW_GotFocus()
SelectionText TxtAca1_PesoAW
End Sub

Private Sub TxtAca1_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca1_PesoBW_GotFocus()
SelectionText TxtAca1_PesoBW
End Sub

Private Sub TxtAca1_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca1_Revirado_GotFocus()
SelectionText TxtAca1_Revirado
End Sub

Private Sub TxtAca1_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_Ancho_Tela_GotFocus()
SelectionText TxtAca3_Ancho_Tela
End Sub

Private Sub TxtAca3_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_Encog_Ancho_GotFocus()
SelectionText TxtAca3_Encog_Ancho
End Sub

Private Sub TxtAca3_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_Encog_Largo_GotFocus()
SelectionText TxtAca3_Encog_Largo
End Sub

Private Sub TxtAca3_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_PesoAW_GotFocus()
SelectionText TxtAca3_PesoAW
End Sub

Private Sub TxtAca3_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAca3_Revirado_GotFocus()
SelectionText TxtAca3_Revirado
End Sub

Private Sub TxtAca3_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_Ancho_Tela_GotFocus()
SelectionText TxtLav_Ancho_Tela
End Sub

Private Sub TxtLav_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_Encog_Ancho_GotFocus()
SelectionText TxtLav_Encog_Ancho
End Sub

Private Sub TxtLav_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_Encog_Largo_GotFocus()
SelectionText TxtLav_Encog_Largo
End Sub

Private Sub TxtLav_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_PesoAW_GotFocus()
SelectionText TxtLav_PesoAW
End Sub

Private Sub TxtLav_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_PesoBW_GotFocus()
SelectionText TxtLav_PesoBW
End Sub

Private Sub TxtLav_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtLav_Revirado_GotFocus()
SelectionText TxtLav_Revirado
End Sub

Private Sub TxtLav_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtObservaciones_Considerables_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtObservaciones_Relevantes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Ancho_Tela_GotFocus()
SelectionText TxtPruBoil_Ancho_Tela
End Sub

Private Sub TxtPruBoil_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Encog_Ancho_GotFocus()
SelectionText TxtPruBoil_Encog_Ancho
End Sub

Private Sub TxtPruBoil_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Encog_Largo_GotFocus()
SelectionText TxtPruBoil_Encog_Largo
End Sub

Private Sub TxtPruBoil_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_PesoAW_GotFocus()
SelectionText TxtPruBoil_PesoAW
End Sub

Private Sub TxtPruBoil_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_PesoBW_GotFocus()
SelectionText TxtPruBoil_PesoBW
End Sub

Private Sub TxtPruBoil_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Revirado_GotFocus()
SelectionText TxtPruBoil_Revirado
End Sub

Private Sub TxtPruBoil_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Ancho_Tela_GotFocus()
SelectionText TxtPruSeca_Ancho_Tela
End Sub

Private Sub TxtPruSeca_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Encog_Ancho_GotFocus()
SelectionText TxtPruSeca_Encog_Ancho
End Sub

Private Sub TxtPruSeca_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Encog_Largo_GotFocus()
SelectionText TxtPruSeca_Encog_Largo
End Sub

Private Sub TxtPruSeca_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_PesoAW_GotFocus()
SelectionText TxtPruSeca_PesoAW
End Sub

Private Sub TxtPruSeca_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_PesoBW_GotFocus()
SelectionText TxtPruSeca_PesoBW
End Sub

Private Sub TxtPruSeca_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Revirado_GotFocus()
SelectionText TxtPruSeca_Revirado
End Sub

Private Sub TxtPruSeca_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_Ancho_Tela_GotFocus()
SelectionText TxtReqAca_Ancho_Tela
End Sub

Private Sub TxtReqAca_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_Encog_Ancho_GotFocus()
SelectionText TxtReqAca_Encog_Ancho
End Sub

Private Sub TxtReqAca_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_Encog_Largo_GotFocus()
SelectionText TxtReqAca_Encog_Largo
End Sub

Private Sub TxtReqAca_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_Peso_BW_GotFocus()
SelectionText TxtReqAca_Peso_BW
End Sub

Private Sub TxtReqAca_Peso_BW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_PesoAW_GotFocus()
SelectionText TxtReqAca_PesoAW
End Sub

Private Sub TxtReqAca_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtReqAca_Revirado_GotFocus()
SelectionText TxtReqAca_Revirado
End Sub

Private Sub TxtReqAca_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_Ancho_Tela_GotFocus()
SelectionText TxtResi1_Ancho_Tela
End Sub

Private Sub TxtResi1_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_Encog_Ancho_GotFocus()
SelectionText TxtResi1_Encog_Ancho
End Sub

Private Sub TxtResi1_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_Encog_Largo_GotFocus()
SelectionText TxtResi1_Encog_Largo
End Sub

Private Sub TxtResi1_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_PesoAW_GotFocus()
SelectionText TxtResi1_PesoAW
End Sub

Private Sub TxtResi1_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_PesoBW_GotFocus()
SelectionText TxtResi1_PesoBW
End Sub

Private Sub TxtResi1_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi1_Revirado_GotFocus()
SelectionText TxtResi1_Revirado
End Sub

Private Sub TxtResi1_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_Ancho_Tela_GotFocus()
SelectionText TxtResi3_Ancho_Tela
End Sub

Private Sub TxtResi3_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_Encog_Ancho_GotFocus()
SelectionText TxtResi3_Encog_Ancho
End Sub

Private Sub TxtResi3_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_Encog_Largo_GotFocus()
SelectionText TxtResi3_Encog_Largo
End Sub

Private Sub TxtResi3_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_PesoAW_GotFocus()
SelectionText TxtResi3_PesoAW
End Sub

Private Sub TxtResi3_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_PesoBW_GotFocus()
SelectionText TxtResi3_PesoBW
End Sub

Private Sub TxtResi3_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtResi3_Revirado_GotFocus()
SelectionText TxtResi3_Revirado
End Sub

Private Sub TxtResi3_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Ancho_Lavado_GotFocus()
SelectionText TxtTambler_Ancho_Lavado
End Sub

Private Sub TxtTambler_Ancho_Lavado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Ancho_Proyectado_GotFocus()
SelectionText TxtTambler_Ancho_Proyectado
End Sub

Private Sub TxtTambler_Ancho_Proyectado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Densidad_GotFocus()
SelectionText TxtTambler_Densidad
End Sub

Private Sub TxtTambler_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub




