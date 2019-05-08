VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmManTelasDatTec 
   Caption         =   "Datos Tecnicos Tela"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   2640
      TabIndex        =   187
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox CmbPrenda 
         Height          =   315
         Left            =   2640
         TabIndex        =   188
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "Prenda Lavada"
         Height          =   195
         Left            =   1080
         TabIndex        =   189
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdProcesos 
      Caption         =   "Procesos Textiles de la Tela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   50
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      Caption         =   "Perchadora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   176
      Top             =   8520
      Width           =   11055
      Begin VB.TextBox TxtPercha_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   81
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Ancho_Salida 
         Height          =   285
         Left            =   3960
         TabIndex        =   82
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Pelo 
         Height          =   285
         Left            =   6840
         TabIndex        =   83
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Contra_Pelo 
         Height          =   285
         Left            =   9720
         TabIndex        =   84
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtPercha_Pases 
         Height          =   285
         Left            =   1440
         TabIndex        =   85
         Top             =   520
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Velocidad 
         Height          =   285
         Left            =   3960
         TabIndex        =   86
         Top             =   520
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Presion 
         Height          =   285
         Left            =   6840
         TabIndex        =   87
         Top             =   520
         Width           =   855
      End
      Begin VB.TextBox TxtPercha_Alim_Rodillo 
         Height          =   285
         Left            =   9720
         TabIndex        =   88
         Top             =   520
         Width           =   735
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Alim. de Rodillo"
         Height          =   195
         Left            =   8400
         TabIndex        =   184
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   183
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   2760
         TabIndex        =   182
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pelo"
         Height          =   195
         Left            =   5760
         TabIndex        =   181
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Contra Pelo"
         Height          =   195
         Left            =   8400
         TabIndex        =   180
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "# Pases"
         Height          =   195
         Left            =   120
         TabIndex        =   179
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   2760
         TabIndex        =   178
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "Presion"
         Height          =   195
         Left            =   5760
         TabIndex        =   177
         Top             =   630
         Width           =   525
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Teñido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   169
      Top             =   600
      Width           =   6255
      Begin VB.TextBox TxtCod_Receta 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1160
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Receta 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   1160
         Width           =   4095
      End
      Begin VB.TextBox TxtCurva_Tenido 
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Top             =   560
         Width           =   855
      End
      Begin VB.TextBox TxtRel_Banos_Kilos 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   560
         Width           =   855
      End
      Begin VB.TextBox TxtRel_Banos_Litro 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   560
         Width           =   855
      End
      Begin VB.TextBox TxtDes_Proveedor 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   1470
         Width           =   3495
      End
      Begin VB.TextBox TxtCod_Proveedor 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1470
         Width           =   1335
      End
      Begin VB.TextBox TxtDes_TipoReceta 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   860
         Width           =   4095
      End
      Begin VB.TextBox TxtCod_TipoReceta 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   860
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Maquina_Tinto 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtCod_Maquina_Tinto 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Lavado"
         Height          =   195
         Left            =   120
         TabIndex        =   186
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   175
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Rel. Baños Kilos"
         Height          =   195
         Left            =   2160
         TabIndex        =   174
         Top             =   645
         Width           =   1155
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Receta"
         Height          =   195
         Left            =   120
         TabIndex        =   173
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Curva Teñido"
         Height          =   195
         Left            =   4320
         TabIndex        =   172
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Rel. Baños Litros"
         Height          =   195
         Left            =   120
         TabIndex        =   171
         Top             =   645
         Width           =   1200
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Maquina Tinto"
         Height          =   195
         Left            =   120
         TabIndex        =   170
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abridor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6360
      TabIndex        =   160
      Top             =   600
      Width           =   4695
      Begin VB.TextBox TxtAbr_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Ancho_Tubular 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   550
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Alimentacion 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Alt_Destorcedor 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1180
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Circunf_Canasta 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Presion_Exprimido 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   550
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Ancho_Salida 
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Estiramiento 
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   1180
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   240
         TabIndex        =   168
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Tubular"
         Height          =   195
         Left            =   240
         TabIndex        =   167
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Alimentacion %"
         Height          =   195
         Left            =   240
         TabIndex        =   166
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alt. Destorcedor"
         Height          =   195
         Left            =   240
         TabIndex        =   165
         Top             =   1275
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Circunf. Canasta"
         Height          =   195
         Left            =   2400
         TabIndex        =   164
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Presion Exprimido"
         Height          =   195
         Left            =   2400
         TabIndex        =   163
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   2400
         TabIndex        =   162
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estiramiento %"
         Height          =   195
         Left            =   2400
         TabIndex        =   161
         Top             =   1275
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acabado / Rama - SECADO"
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
      TabIndex        =   147
      Top             =   2400
      Width           =   5415
      Begin VB.TextBox TxtRama_Ancho_Cadena 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Densidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_SobreAlimSup 
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   1725
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Temperatura 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Densidad_Salida 
         Height          =   285
         Left            =   4080
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Ancho_Salida 
         Height          =   285
         Left            =   4080
         TabIndex        =   26
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Presion 
         Height          =   285
         Left            =   4080
         TabIndex        =   24
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox TxtRama_SobreAlimInf 
         Height          =   285
         Left            =   4080
         TabIndex        =   20
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_SobreAlimSal 
         Height          =   285
         Left            =   4080
         TabIndex        =   22
         Top             =   510
         Width           =   855
      End
      Begin VB.ComboBox cmbRama_Vapor 
         Height          =   315
         Left            =   4080
         TabIndex        =   30
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Cadena"
         Height          =   195
         Left            =   120
         TabIndex        =   159
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   158
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   157
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SobreAlim. Sup. %"
         Height          =   195
         Left            =   120
         TabIndex        =   156
         Top             =   1830
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Temperatura"
         Height          =   195
         Left            =   120
         TabIndex        =   155
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   120
         TabIndex        =   154
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Salida"
         Height          =   195
         Left            =   2760
         TabIndex        =   153
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   2760
         TabIndex        =   152
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Presion"
         Height          =   195
         Left            =   2760
         TabIndex        =   151
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "SobreAli. Inf. %"
         Height          =   195
         Left            =   2760
         TabIndex        =   150
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "SobreAli. Salida %"
         Height          =   195
         Left            =   2760
         TabIndex        =   149
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Vapor"
         Height          =   195
         Left            =   2760
         TabIndex        =   148
         Top             =   1830
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "HidroExtractora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   137
      Top             =   4560
      Width           =   5415
      Begin VB.TextBox TxtHidro_Alimentacion_Baja 
         Height          =   285
         Left            =   4005
         TabIndex        =   48
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Alimentacion_Media 
         Height          =   285
         Left            =   4005
         TabIndex        =   46
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Alimentacion_Alta 
         Height          =   285
         Left            =   4005
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Presion2 
         Height          =   285
         Left            =   1440
         TabIndex        =   51
         Top             =   1455
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Presion1 
         Height          =   285
         Left            =   1440
         TabIndex        =   49
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   855
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Ancho_Salida 
         Height          =   285
         Left            =   1440
         TabIndex        =   45
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtHidro_Ensanchador 
         Height          =   285
         Left            =   4005
         TabIndex        =   50
         Top             =   1185
         Width           =   855
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Aliment. Baja %"
         Height          =   195
         Left            =   2760
         TabIndex        =   146
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Aliment. Media %"
         Height          =   195
         Left            =   2760
         TabIndex        =   145
         Top             =   675
         Width           =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Alimentac. Alta %"
         Height          =   195
         Left            =   2760
         TabIndex        =   144
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Presion 2"
         Height          =   195
         Left            =   120
         TabIndex        =   143
         Top             =   1545
         Width           =   660
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Presion 1"
         Height          =   195
         Left            =   120
         TabIndex        =   142
         Top             =   1275
         Width           =   660
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   120
         TabIndex        =   141
         Top             =   975
         Width           =   705
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   120
         TabIndex        =   140
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   139
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Ensanchador"
         Height          =   195
         Left            =   2760
         TabIndex        =   138
         Top             =   1275
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Secado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5520
      TabIndex        =   127
      Top             =   4560
      Width           =   5535
      Begin VB.TextBox TxtSec_Densidad 
         Height          =   285
         Left            =   4200
         TabIndex        =   59
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Ancho_Salida 
         Height          =   285
         Left            =   1440
         TabIndex        =   54
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtSec_SobreAlimentacion 
         Height          =   285
         Left            =   1440
         TabIndex        =   56
         Top             =   855
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Temp1 
         Height          =   285
         Left            =   1440
         TabIndex        =   58
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Temp2 
         Height          =   285
         Left            =   1440
         TabIndex        =   60
         Top             =   1455
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Temp3 
         Height          =   285
         Left            =   4200
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Velocidad 
         Height          =   285
         Left            =   4200
         TabIndex        =   55
         Top             =   555
         Width           =   855
      End
      Begin VB.TextBox TxtSec_Encogimiento 
         Height          =   285
         Left            =   4200
         TabIndex        =   57
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Densidad"
         Height          =   195
         Left            =   2880
         TabIndex        =   136
         Top             =   1275
         Width           =   675
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "SobreAliment. %"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   975
         Width           =   1140
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Temp. 1"
         Height          =   195
         Left            =   120
         TabIndex        =   132
         Top             =   1275
         Width           =   585
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Temp. 2"
         Height          =   195
         Left            =   120
         TabIndex        =   131
         Top             =   1545
         Width           =   585
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Temp. 3"
         Height          =   195
         Left            =   2880
         TabIndex        =   130
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   2880
         TabIndex        =   129
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Encog."
         Height          =   195
         Left            =   2880
         TabIndex        =   128
         Top             =   975
         Width           =   510
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Compactadora"
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
      Left            =   5520
      TabIndex        =   117
      Top             =   6360
      Width           =   5535
      Begin VB.TextBox TxtCompa_Densidad 
         Height          =   285
         Left            =   4110
         TabIndex        =   78
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Tension 
         Height          =   285
         Left            =   4110
         TabIndex        =   76
         Top             =   550
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Teflon 
         Height          =   285
         Left            =   4110
         TabIndex        =   74
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   90
         Top             =   1460
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Temperatura 
         Height          =   270
         Left            =   1440
         TabIndex        =   79
         Top             =   1150
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Ensanchador 
         Height          =   285
         Left            =   1440
         TabIndex        =   77
         Top             =   860
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Ancho_Salida 
         Height          =   285
         Left            =   1440
         TabIndex        =   75
         Top             =   550
         Width           =   855
      End
      Begin VB.TextBox TxtCompa_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   73
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CmbCompa_Vapor 
         Height          =   315
         Left            =   4110
         TabIndex        =   80
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "SobreAliment. %"
         Height          =   195
         Left            =   2880
         TabIndex        =   126
         Top             =   975
         Width           =   1140
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Tension"
         Height          =   195
         Left            =   2880
         TabIndex        =   125
         Top             =   675
         Width           =   570
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Teflon"
         Height          =   195
         Left            =   2880
         TabIndex        =   124
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   120
         TabIndex        =   123
         Top             =   1545
         Width           =   705
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "Temperatura"
         Height          =   195
         Left            =   120
         TabIndex        =   122
         Top             =   1275
         Width           =   900
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Ensanchador"
         Height          =   195
         Left            =   120
         TabIndex        =   121
         Top             =   975
         Width           =   945
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   120
         TabIndex        =   120
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   119
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Vapor"
         Height          =   195
         Left            =   2880
         TabIndex        =   118
         Top             =   1275
         Width           =   420
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Acabado / Rama"
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
      Left            =   5520
      TabIndex        =   104
      Top             =   2400
      Width           =   5535
      Begin VB.ComboBox cmbRama_Resinado_Vapor 
         Height          =   315
         Left            =   3840
         TabIndex        =   42
         Top             =   1725
         Width           =   1575
      End
      Begin VB.TextBox TxtRama_Resinado_SobreAlimSal 
         Height          =   285
         Left            =   4200
         TabIndex        =   34
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_SobreAlimInf 
         Height          =   285
         Left            =   4200
         TabIndex        =   32
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Presion 
         Height          =   285
         Left            =   4200
         TabIndex        =   36
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Ancho_Salida 
         Height          =   285
         Left            =   4200
         TabIndex        =   38
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Densidad_Salida 
         Height          =   285
         Left            =   4200
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Temperatura 
         Height          =   285
         Left            =   1440
         TabIndex        =   39
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_SobreAlimSup 
         Height          =   285
         Left            =   1440
         TabIndex        =   41
         Top             =   1725
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Densidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Resinado_Ancho_Cadena 
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "Aca. Rama/Seco"
         Height          =   195
         Left            =   2520
         TabIndex        =   116
         Top             =   1830
         Width           =   1245
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "SobreAli. Salida %"
         Height          =   195
         Left            =   2880
         TabIndex        =   115
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "SobreAlim. Inf. %"
         Height          =   195
         Left            =   2880
         TabIndex        =   114
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "Presion"
         Height          =   195
         Left            =   2880
         TabIndex        =   113
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   2880
         TabIndex        =   112
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Salida"
         Height          =   195
         Left            =   2880
         TabIndex        =   111
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   120
         TabIndex        =   110
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "Temperatura"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "SobreAlim. Sup. %"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   1830
         Width           =   1290
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Cadena"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   915
         Width           =   1065
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Acabado / Rama - TERMOFIJADO"
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
      TabIndex        =   91
      Top             =   6360
      Width           =   5415
      Begin VB.ComboBox cmbRama_Termo_Vapor 
         Height          =   315
         Left            =   3990
         TabIndex        =   72
         Top             =   1725
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_SobreAlimSal 
         Height          =   285
         Left            =   3990
         TabIndex        =   64
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_SobreAlimInf 
         Height          =   285
         Left            =   3990
         TabIndex        =   62
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Presion 
         Height          =   285
         Left            =   3990
         TabIndex        =   66
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Ancho_Salida 
         Height          =   285
         Left            =   3990
         TabIndex        =   68
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Densidad_Salida 
         Height          =   285
         Left            =   3990
         TabIndex        =   70
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Velocidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   67
         Top             =   1125
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Temperatura 
         Height          =   285
         Left            =   1440
         TabIndex        =   69
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_SobreAlimSup 
         Height          =   285
         Left            =   1440
         TabIndex        =   71
         Top             =   1725
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Ancho_Entrada 
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Densidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   63
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox TxtRama_Termo_Ancho_Cadena 
         Height          =   285
         Left            =   1440
         TabIndex        =   65
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         Caption         =   "Vapor"
         Height          =   195
         Left            =   2760
         TabIndex        =   103
         Top             =   1890
         Width           =   420
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         Caption         =   "SobreAli.Salida %"
         Height          =   195
         Left            =   2760
         TabIndex        =   102
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         Caption         =   "SobreAlim. Inf. %"
         Height          =   195
         Left            =   2760
         TabIndex        =   101
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         Caption         =   "Presion"
         Height          =   195
         Left            =   2760
         TabIndex        =   100
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label98 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Salida"
         Height          =   195
         Left            =   2760
         TabIndex        =   99
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Salida"
         Height          =   195
         Left            =   2760
         TabIndex        =   98
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label101 
         AutoSize        =   -1  'True
         Caption         =   "Temperatura"
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label102 
         AutoSize        =   -1  'True
         Caption         =   "SobreAlim. Sup. %"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   1830
         Width           =   1290
      End
      Begin VB.Label Label103 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label104 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label105 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Cadena"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   915
         Width           =   1065
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4320
      TabIndex        =   89
      Top             =   9600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmManTelasDatTec.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmManTelasDatTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Public sCod_Tela As String, sDes_tela As String, sFamite As String, sCod_Comb As String, sDes_Comb As String
Public Codigo As String, Descripcion As String, TipoAdd As String

Private Sub cmbRama_Resinado_Vapor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdProcesos_Click()
If UCase(sFamite) <> "DE" Then
    Load FrmManTela_Procesos_Textil
    FrmManTela_Procesos_Textil.vCod_Tela = Me.sCod_Tela
    FrmManTela_Procesos_Textil.txtcod_tela.Text = Me.sCod_Tela
    FrmManTela_Procesos_Textil.txtdes_tela.Text = Me.sDes_tela
    FrmManTela_Procesos_Textil.CARGA_GRID
    FrmManTela_Procesos_Textil.Show vbModal
    Set FrmManTela_Procesos_Textil = Nothing
Else
    Load FrmManTelaComb_Procesos_Textil
    FrmManTelaComb_Procesos_Textil.vCod_Tela = Me.sCod_Tela
    FrmManTelaComb_Procesos_Textil.vCod_Comb = Me.sCod_Comb
    FrmManTelaComb_Procesos_Textil.txtcod_tela.Text = Me.sCod_Tela
    FrmManTelaComb_Procesos_Textil.txtdes_tela.Text = Me.sDes_tela
    FrmManTelaComb_Procesos_Textil.TxtCod_Comb.Text = Me.sCod_Comb
    FrmManTelaComb_Procesos_Textil.TxtDes_Comb.Text = Me.sDes_Comb
    
    FrmManTelaComb_Procesos_Textil.CARGA_GRID
    FrmManTelaComb_Procesos_Textil.Show vbModal
    Set FrmManTelaComb_Procesos_Textil = Nothing
End If
End Sub

Private Sub Form_Load()
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
Call CARGA_COMBOS
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
Dim rs As New ADODB.Recordset
On Error GoTo errDatos

Call CARGA_COMBOS

If UCase(sFamite) <> "DE" Then
    strSQL = "Tx_Muestra_DATOS_TECNICOS_TELA_Nuevo '" & sCod_Tela & "'"
Else
    strSQL = "Tx_Muestra_DATOS_TECNICOS_TELAComb_Nuevo '" & sCod_Tela & "','" & sCod_Comb & "'"
End If


Set rs = Nothing
rs.CursorLocation = adUseClient

rs.Open strSQL, cCONNECT, 3
If rs.RecordCount Then
    Me.TxtCod_Maquina_Tinto = rs!Ten_Cod_Maquina_Tinto
    Me.TxtDes_Maquina_Tinto = rs!des_maquina_tinto
    Me.TxtRel_Banos_Litro = rs!Ten_Rel_Bano_Litros
    Me.TxtRel_Banos_Kilos = rs!Ten_Rel_Bano_Kilos
    Me.TxtCurva_Tenido = rs!Ten_Curva_Tenido
    Me.TxtCod_TipoReceta = rs!Ten_Cod_Tipo_Receta
    Me.TxtDes_TipoReceta = Trim(rs!des_tiporeceta)
    Me.TxtCod_Proveedor = rs!Ten_Cod_Proveedor
    Me.TxtDes_Proveedor = Trim(rs!des_proveedor)

    Me.TxtAbr_Velocidad = rs!Abr_Velocidad
    Me.TxtAbr_Ancho_Tubular = rs!Abr_Ancho_Tubular
    Me.TxtAbr_Alimentacion = rs!Abr_Alimentacion
    Me.TxtAbr_Alt_Destorcedor = rs!Abr_Alt_Destorcedor
    Me.TxtAbr_Circunf_Canasta = rs!Abr_Circunferencia_Canasta
    Me.TxtAbr_Presion_Exprimido = rs!Abr_Presion_Exprimido
    Me.TxtAbr_Ancho_Salida = rs!Abr_Ancho_Salida
    Me.TxtAbr_Estiramiento = rs!Abr_Estiramiento
    
    Me.TxtRama_Ancho_Entrada = rs!AcaRama_Secado_Ancho_Entrada
    Me.TxtRama_Densidad = rs!AcaRama_Secado_Densidad_Entrada
    Me.TxtRama_Ancho_Cadena = rs!AcaRama_Ancho_Cadena
    Me.TxtRama_Velocidad = rs!AcaRama_Secado_Velocidad
    Me.TxtRama_Temperatura = rs!AcaRama_Secado_Temperatura
    Me.TxtRama_SobreAlimSup = rs!AcaRama_Secado_SobreAlimentacion_Sup
    Me.TxtRama_SobreAlimInf = rs!AcaRama_Secado_SobreAlimentacion_Inf
    Me.TxtRama_SobreAlimSal = rs!AcaRama_Secado_SobreAlimentacion_Salida
    Me.TxtRama_Presion = rs!AcaRama_Secado_Presion
    Me.TxtRama_Ancho_Salida = rs!AcaRama_Secado_Ancho_Salida
    Me.TxtRama_Densidad_Salida = rs!AcaRama_Secado_Densidad_Salida
    Call BuscaCombo(rs!AcaRama_Secado_Vapor, 1, cmbRama_Vapor)
    
    Me.TxtRama_Resinado_Ancho_Entrada = rs!AcaRama_Resinado_Secado_Ancho_Entrada
    Me.TxtRama_Resinado_Densidad = rs!AcaRama_Resinado_Secado_Densidad_Entrada
    Me.TxtRama_Resinado_Ancho_Cadena = rs!AcaRama_Resinado_Ancho_Cadena
    Me.TxtRama_Resinado_Velocidad = rs!AcaRama_Resinado_Secado_Velocidad
    Me.TxtRama_Resinado_Temperatura = rs!AcaRama_Resinado_Secado_Temperatura
    Me.TxtRama_Resinado_SobreAlimSup = rs!AcaRama_Resinado_Secado_SobreAlimentacion_Sup
    Me.TxtRama_Resinado_SobreAlimInf = rs!AcaRama_Resinado_Secado_SobreAlimentacion_Inf
    Me.TxtRama_Resinado_SobreAlimSal = rs!AcaRama_Resinado_Secado_SobreAlimentacion_Salida
    Me.TxtRama_Resinado_Presion = rs!AcaRama_Resinado_Secado_Presion
    Me.TxtRama_Resinado_Ancho_Salida = rs!AcaRama_Resinado_Secado_Ancho_Salida
    Me.TxtRama_Resinado_Densidad_Salida = rs!AcaRama_Resinado_Secado_Densidad_Salida
    Call BuscaCombo(rs!AcaRama_Resinado_Secado_Vapor, 1, cmbRama_Resinado_Vapor)
    
    
    Me.TxtRama_Termo_Ancho_Entrada = rs!AcaRama_Termo_Secado_Ancho_Entrada
    Me.TxtRama_Termo_Densidad = rs!AcaRama_Termo_Secado_Densidad_Entrada
    Me.TxtRama_Termo_Ancho_Cadena = rs!AcaRama_Termo_Ancho_Cadena
    Me.TxtRama_Termo_Velocidad = rs!AcaRama_Termo_Secado_Velocidad
    Me.TxtRama_Termo_Temperatura = rs!AcaRama_Termo_Secado_Temperatura
    Me.TxtRama_Termo_SobreAlimSup = rs!AcaRama_Termo_Secado_SobreAlimentacion_Sup
    Me.TxtRama_Termo_SobreAlimInf = rs!AcaRama_Termo_Secado_SobreAlimentacion_Inf
    Me.TxtRama_Termo_SobreAlimSal = rs!AcaRama_Termo_Secado_SobreAlimentacion_Salida
    Me.TxtRama_Termo_Presion = rs!AcaRama_Termo_Secado_Presion
    Me.TxtRama_Termo_Ancho_Salida = rs!AcaRama_Termo_Secado_Ancho_Salida
    Me.TxtRama_Termo_Densidad_Salida = rs!AcaRama_Termo_Secado_Densidad_Salida
    Call BuscaCombo(rs!AcaRama_Termo_Secado_Vapor, 1, cmbRama_Termo_Vapor)
    
    
    Me.TxtHidro_Ancho_Entrada = rs!Hidro_Ancho_Entrada
    Me.TxtHidro_Ancho_Salida = rs!Hidro_Ancho_Salida
    Me.TxtHidro_Velocidad = rs!Hidro_Velocidad
    Me.TxtHidro_Presion1 = rs!Hidro_Presion1
    Me.TxtHidro_Presion2 = rs!Hidro_Presion2
    Me.TxtHidro_Alimentacion_Alta = rs!Hidro_Alimentacion_Alta
    Me.TxtHidro_Alimentacion_Media = rs!Hidro_Alimentacion_Media
    Me.TxtHidro_Alimentacion_Baja = rs!Hidro_Alimentacion_Baja
    Me.TxtHidro_Ensanchador = rs!Hidro_Ensanchador
    Me.TxtSec_Ancho_Entrada = rs!Seca_Ancho_Entrada
    Me.TxtSec_Ancho_Salida = rs!Seca_Ancho_Salida
    Me.TxtSec_SobreAlimentacion = rs!Seca_SobreAlimentacion
    Me.TxtSec_Temp1 = rs!Seca_Temp1
    Me.TxtSec_Temp2 = rs!Seca_Temp2
    Me.TxtSec_Temp3 = rs!Seca_Temp3
    Me.TxtSec_Velocidad = rs!Seca_Velocidad
    Me.TxtSec_Encogimiento = rs!Seca_Encogimiento
    Me.TxtSec_Densidad = rs!Seca_Densidad
    Me.TxtCompa_Ancho_Entrada = rs!Compa_Ancho_Entrada
    Me.TxtCompa_Ancho_Salida = rs!Compa_Ancho_Salida
    Me.TxtCompa_Ensanchador = rs!Compa_Ensanchador
    Me.TxtCompa_Temperatura = rs!Compa_Temperatura
    Me.TxtCompa_Velocidad = rs!Compa_Velocidad
    Me.TxtCompa_Teflon = rs!Compa_Teflon
    Me.TxtCompa_Tension = rs!Compa_Tension
    Me.TxtCompa_Densidad = rs!Compa_Densidad
    Call BuscaCombo(rs!Compa_Vapor, 1, CmbCompa_Vapor)
    Me.TxtPercha_Ancho_Entrada = rs!Perc_Ancho_Entrada
    Me.TxtPercha_Ancho_Salida = rs!Perc_Ancho_Salida
    Me.TxtPercha_Pelo = rs!Perc_Pelo
    Me.TxtPercha_Contra_Pelo = rs!Perc_Contra_Pelo
    Me.TxtPercha_Pases = rs!Perc_Pases
    Me.TxtPercha_Velocidad = rs!Perc_Velocidad
    Me.TxtPercha_Presion = rs!Perc_Presion
    Me.TxtPercha_Alim_Rodillo = rs!Perc_Alimentacion_rodillo
    Me.TxtCod_Receta = rs!cod_Receta
    Me.TxtDes_Receta = rs!des_receta
    
    Call BuscaCombo(rs!prenda_lavada, 1, CmbPrenda)

End If

Exit Sub
errDatos:
    ErrorHandler Err, "Carga Datos"
End Sub

Sub Grabar()
Dim vTipo_Lavado As String
On Error GoTo errGrabar

    If Trim(Me.TxtRel_Banos_Litro) = "" Then TxtRel_Banos_Litro = "0"
    If Trim(Me.TxtRel_Banos_Kilos) = "" Then TxtRel_Banos_Kilos = "0"
    If Trim(Me.TxtCurva_Tenido) = "" Then TxtCurva_Tenido = "0"

    If Trim(Me.TxtAbr_Velocidad) = "" Then TxtAbr_Velocidad = "0"
    If Trim(Me.TxtAbr_Ancho_Tubular) = "" Then TxtAbr_Ancho_Tubular = "0"
    If Trim(Me.TxtAbr_Alimentacion) = "" Then TxtAbr_Alimentacion = "0"
    If Trim(Me.TxtAbr_Alt_Destorcedor) = "" Then TxtAbr_Alt_Destorcedor = "0"
    If Trim(Me.TxtAbr_Circunf_Canasta) = "" Then TxtAbr_Circunf_Canasta = "0"
    If Trim(Me.TxtAbr_Presion_Exprimido) = "" Then TxtAbr_Presion_Exprimido = "0"
    If Trim(Me.TxtAbr_Ancho_Salida) = "" Then TxtAbr_Ancho_Salida = "0"
    If Trim(Me.TxtAbr_Estiramiento) = "" Then TxtAbr_Estiramiento = "0"
    If Trim(TxtRama_Ancho_Entrada) = "" Then TxtRama_Ancho_Entrada = "0"
    If Trim(TxtRama_Densidad) = "" Then TxtRama_Densidad = "0"
    If Trim(TxtRama_Ancho_Cadena) = "" Then TxtRama_Ancho_Cadena = "0"
    If Trim(TxtRama_Velocidad) = "" Then TxtRama_Velocidad = "0"
    If Trim(TxtRama_Temperatura) = "" Then TxtRama_Temperatura = "0"
    If Trim(Me.TxtRama_SobreAlimSup) = "" Then TxtRama_SobreAlimSup = "0"
    If Trim(Me.TxtRama_SobreAlimInf) = "" Then TxtRama_SobreAlimInf = "0"
    If Trim(Me.TxtRama_SobreAlimSal) = "" Then TxtRama_SobreAlimSal = "0"
    If Trim(Me.TxtRama_Presion) = "" Then TxtRama_Presion = "0"
    If Trim(Me.TxtRama_Ancho_Salida) = "" Then TxtRama_Ancho_Salida = "0"
    If Trim(Me.TxtRama_Densidad_Salida) = "" Then TxtRama_Densidad_Salida = "0"
    
    If Trim(Me.TxtHidro_Ancho_Entrada) = "" Then TxtHidro_Ancho_Entrada = "0"
    If Trim(Me.TxtHidro_Ancho_Salida) = "" Then TxtHidro_Ancho_Salida = "0"
    If Trim(Me.TxtHidro_Velocidad) = "" Then TxtHidro_Velocidad = "0"
    If Trim(Me.TxtHidro_Presion1) = "" Then TxtHidro_Presion1 = "0"
    If Trim(Me.TxtHidro_Presion2) = "" Then TxtHidro_Presion2 = "0"
    If Trim(Me.TxtHidro_Alimentacion_Alta) = "" Then TxtHidro_Alimentacion_Alta = "0"
    If Trim(Me.TxtHidro_Alimentacion_Media) = "" Then TxtHidro_Alimentacion_Media = "0"
    If Trim(Me.TxtHidro_Alimentacion_Baja) = "" Then TxtHidro_Alimentacion_Baja = "0"
    If Trim(Me.TxtHidro_Ensanchador) = "" Then TxtHidro_Ensanchador = "0"
    If Trim(Me.TxtSec_Ancho_Entrada) = "" Then TxtSec_Ancho_Entrada = "0"
    If Trim(Me.TxtSec_Ancho_Salida) = "" Then TxtSec_Ancho_Salida = "0"
    If Trim(Me.TxtSec_SobreAlimentacion) = "" Then TxtSec_SobreAlimentacion = "0"
    If Trim(Me.TxtSec_Temp1) = "" Then TxtSec_Temp1 = "0"
    If Trim(Me.TxtSec_Temp2) = "" Then TxtSec_Temp2 = "0"
    If Trim(Me.TxtSec_Temp3) = "" Then TxtSec_Temp3 = "0"
    If Trim(Me.TxtSec_Velocidad) = "" Then TxtSec_Velocidad = "0"
    If Trim(Me.TxtSec_Encogimiento) = "" Then TxtSec_Encogimiento = "0"
    If Trim(Me.TxtSec_Densidad) = "" Then TxtSec_Densidad = "0"
    If Trim(Me.TxtCompa_Ancho_Entrada) = "" Then TxtCompa_Ancho_Entrada = "0"
    If Trim(Me.TxtCompa_Ancho_Salida) = "" Then TxtCompa_Ancho_Salida = "0"
    If Trim(Me.TxtCompa_Ensanchador) = "" Then TxtCompa_Ensanchador = "0"
    If Trim(Me.TxtCompa_Temperatura) = "" Then TxtCompa_Temperatura = "0"
    If Trim(Me.TxtCompa_Velocidad) = "" Then TxtCompa_Velocidad = "0"
    If Trim(Me.TxtCompa_Teflon) = "" Then TxtCompa_Teflon = "0"
    If Trim(Me.TxtCompa_Tension) = "" Then TxtCompa_Tension = "0"
    If Trim(Me.TxtCompa_Densidad) = "" Then TxtCompa_Densidad = "0"
    If Trim(Me.TxtPercha_Ancho_Entrada) = "" Then TxtPercha_Ancho_Entrada = "0"
    If Trim(Me.TxtPercha_Ancho_Salida) = "" Then TxtPercha_Ancho_Salida = "0"
    If Trim(Me.TxtPercha_Pelo) = "" Then TxtPercha_Pelo = "0"
    If Trim(Me.TxtPercha_Contra_Pelo) = "" Then TxtPercha_Contra_Pelo = "0"
    If Trim(Me.TxtPercha_Pases) = "" Then TxtPercha_Pases = "0"
    If Trim(Me.TxtPercha_Velocidad) = "" Then TxtPercha_Velocidad = "0"
    If Trim(Me.TxtPercha_Presion) = "" Then TxtPercha_Presion = "0"
    If Trim(Me.TxtPercha_Alim_Rodillo) = "" Then TxtPercha_Alim_Rodillo = "0"
    
    If Trim(Me.TxtRama_Resinado_Ancho_Entrada) = "" Then TxtRama_Resinado_Ancho_Entrada = "0"
    If Trim(Me.TxtRama_Resinado_Densidad) = "" Then TxtRama_Resinado_Densidad = "0"
    If Trim(Me.TxtRama_Resinado_Ancho_Cadena) = "" Then TxtRama_Resinado_Ancho_Cadena = "0"
    If Trim(Me.TxtRama_Resinado_Velocidad) = "" Then TxtRama_Resinado_Velocidad = "0"
    If Trim(Me.TxtRama_Resinado_Temperatura) = "" Then TxtRama_Resinado_Temperatura = "0"
    If Trim(Me.TxtRama_Resinado_SobreAlimSup) = "" Then TxtRama_Resinado_SobreAlimSup = "0"
    If Trim(Me.TxtRama_Resinado_SobreAlimInf) = "" Then TxtRama_Resinado_SobreAlimInf = "0"
    If Trim(Me.TxtRama_Resinado_SobreAlimSal) = "" Then TxtRama_Resinado_SobreAlimSal = "0"
    If Trim(Me.TxtRama_Resinado_Presion) = "" Then TxtRama_Resinado_Presion = "0"
    If Trim(Me.TxtRama_Resinado_Ancho_Salida) = "" Then TxtRama_Resinado_Ancho_Salida = "0"
    If Trim(Me.TxtRama_Resinado_Densidad_Salida) = "" Then TxtRama_Resinado_Densidad_Salida = "0"
    
    If Trim(Me.TxtRama_Termo_Ancho_Entrada) = "" Then TxtRama_Termo_Ancho_Entrada = "0"
    If Trim(Me.TxtRama_Termo_Densidad) = "" Then TxtRama_Termo_Densidad = "0"
    If Trim(Me.TxtRama_Termo_Ancho_Cadena) = "" Then TxtRama_Termo_Ancho_Cadena = "0"
    If Trim(Me.TxtRama_Termo_Velocidad) = "" Then TxtRama_Termo_Velocidad = "0"
    If Trim(Me.TxtRama_Termo_Temperatura) = "" Then TxtRama_Termo_Temperatura = "0"
    If Trim(Me.TxtRama_Termo_SobreAlimSup) = "" Then TxtRama_Termo_SobreAlimSup = "0"
    If Trim(Me.TxtRama_Termo_SobreAlimInf) = "" Then TxtRama_Termo_SobreAlimInf = "0"
    If Trim(Me.TxtRama_Termo_SobreAlimSal) = "" Then TxtRama_Termo_SobreAlimSal = "0"
    If Trim(Me.TxtRama_Termo_Presion) = "" Then TxtRama_Termo_Presion = "0"
    If Trim(Me.TxtRama_Termo_Ancho_Salida) = "" Then TxtRama_Termo_Ancho_Salida = "0"
    If Trim(Me.TxtRama_Termo_Densidad_Salida) = "" Then TxtRama_Termo_Densidad_Salida = "0"

If UCase(sFamite) <> "DE" Then
    strSQL = "UP_TX_ACTUALIZA_DATOS_TECNICOS_TELA_NUEVOS '" & _
        sCod_Tela & "','"
Else
    strSQL = "UP_TX_ACTUALIZA_DATOS_TECNICOS_TELACOMB_NUEVOS '" & _
        sCod_Tela & "','" & sCod_Comb & "','"
End If

strSQL = strSQL & TxtCod_Maquina_Tinto.Text & "','" & TxtRel_Banos_Litro.Text & "','" & _
        TxtRel_Banos_Kilos & "','" & TxtCurva_Tenido & "','" & TxtCod_TipoReceta & "','" & TxtCod_Proveedor & "','" & _
        TxtAbr_Velocidad & "','" & TxtAbr_Ancho_Tubular & "','" & TxtAbr_Alimentacion & "','" & _
        TxtAbr_Alt_Destorcedor & "','" & TxtAbr_Circunf_Canasta & "','" & TxtAbr_Presion_Exprimido & "','" & _
        TxtAbr_Ancho_Salida & "','" & TxtAbr_Estiramiento & "','" & TxtRama_Ancho_Entrada & "','" & _
        TxtRama_Densidad & "','" & TxtRama_Ancho_Cadena & "','" & TxtRama_Velocidad & "','" & _
        TxtRama_Temperatura & "','" & TxtRama_SobreAlimSup & "','" & TxtRama_SobreAlimInf & "','" & _
        TxtRama_SobreAlimSal & "','" & TxtRama_Presion & "','" & TxtRama_Ancho_Salida & "','" & _
        TxtRama_Densidad_Salida & "','" & cmbRama_Vapor & "','" & TxtHidro_Ancho_Entrada & "','" & _
        TxtHidro_Ancho_Salida & "','" & TxtHidro_Velocidad & "','" & TxtHidro_Presion1 & "','" & _
        TxtHidro_Presion2 & "','" & TxtHidro_Alimentacion_Alta & "','" & TxtHidro_Alimentacion_Media & "','" & _
        TxtHidro_Alimentacion_Baja & "','" & TxtHidro_Ensanchador & "','" & TxtSec_Ancho_Entrada & "','" & _
        TxtSec_Ancho_Salida & "','" & TxtSec_SobreAlimentacion & "','" & TxtSec_Temp1 & "','" & _
        TxtSec_Temp2 & "','" & TxtSec_Temp3 & "','" & TxtSec_Velocidad & "','" & TxtSec_Encogimiento & "','" & _
        TxtSec_Densidad & "','" & TxtCompa_Ancho_Entrada & "','" & TxtCompa_Ancho_Salida & "','" & _
        TxtCompa_Ensanchador & "','" & TxtCompa_Temperatura & "','" & TxtCompa_Velocidad & "','" & _
        TxtCompa_Teflon & "','" & TxtCompa_Tension & "','" & TxtCompa_Densidad & "','" & CmbCompa_Vapor & "','" & _
        TxtPercha_Ancho_Entrada & "','" & TxtPercha_Ancho_Salida & "','" & TxtPercha_Pelo & "','" & _
        TxtPercha_Contra_Pelo & "','" & TxtPercha_Pases & "','" & TxtPercha_Velocidad & "','" & _
        TxtPercha_Presion & "','" & TxtPercha_Alim_Rodillo & "','"
        
strSQL = strSQL & Me.TxtRama_Resinado_Ancho_Entrada & "','" & Me.TxtRama_Resinado_Densidad & "','" & Me.TxtRama_Resinado_Ancho_Cadena & "','" & _
    Me.TxtRama_Resinado_Velocidad & "','" & Me.TxtRama_Resinado_Temperatura & "','" & Me.TxtRama_Resinado_SobreAlimSup & "','" & _
    Me.TxtRama_Resinado_SobreAlimInf & "','" & Me.TxtRama_Resinado_SobreAlimSal & "','" & Me.TxtRama_Resinado_Presion & "','" & _
    Me.TxtRama_Resinado_Ancho_Salida & "','" & Me.TxtRama_Resinado_Densidad_Salida & "','" & cmbRama_Resinado_Vapor & "','" & _
    Me.TxtRama_Termo_Ancho_Entrada & "','" & Me.TxtRama_Termo_Densidad & "','" & Me.TxtRama_Termo_Ancho_Cadena & "','" & _
    Me.TxtRama_Termo_Velocidad & "','" & Me.TxtRama_Termo_Temperatura & "','" & Me.TxtRama_Termo_SobreAlimSup & "','" & _
    Me.TxtRama_Termo_SobreAlimInf & "','" & Me.TxtRama_Termo_SobreAlimSal & "','" & Me.TxtRama_Termo_Presion & "','" & _
    Me.TxtRama_Termo_Ancho_Salida & "','" & Me.TxtRama_Termo_Densidad_Salida & "','" & cmbRama_Termo_Vapor & "','" & vusu & "','" & ComputerName & "','" & TxtCod_Receta.Text & "','" & CmbPrenda.Text & "'"

'strSQL = "UP_TX_ACTUALIZA_DATOS_TECNICOS_TELA_NUEVOS '" & _
        sCod_Tela & "','" & TxtCod_Maquina_Tinto.Text & "','" & TxtRel_Banos_Litro.Text & "','" & _
        TxtRel_Banos_Kilos & "','" & TxtCurva_Tenido & "','" & TxtCod_TipoReceta & "','" & TxtCod_Proveedor & "','" & _
        TxtAbr_Velocidad & "','" & TxtAbr_Ancho_Tubular & "','" & TxtAbr_Alimentacion & "','" & _
        TxtAbr_Alt_Destorcedor & "','" & TxtAbr_Circunf_Canasta & "','" & TxtAbr_Presion_Exprimido & "','" & TxtAbr_Ancho_Salida & "','" & _
        TxtAbr_Estiramiento & "','" & TxtRama_Ancho_Entrada & "','" & TxtRama_Densidad & "','" & TxtRama_Ancho_Cadena & "','" & _
        TxtRama_Velocidad & "','" & TxtRama_Temperatura & "','" & TxtRama_SobreAlimSup & "','" & TxtRama_SobreAlimInf & "','" & _
        TxtRama_SobreAlimSal & "','" & TxtRama_Presion & "','" & TxtRama_Ancho_Salida & "','" & TxtRama_Densidad_Salida & "','" & _
        cmbRama_Vapor & "','" & AcaRama_Resinado_Secado_Ancho_Entrada & "','" & AcaRama_Resinado_Secado_Densidad_Entrada & "','" & _
        AcaRama_Resinado_Ancho_Cadena & "','" & AcaRama_Resinado_Secado_Velocidad & "','" & AcaRama_Resinado_Secado_Temperatura & "','" & AcaRama_Resinado_Secado_SobreAlimentacion_Sup & "','" & _
        AcaRama_Resinado_Secado_SobreAlimentacion_Inf & "','" & AcaRama_Resinado_Secado_SobreAlimentacion_Salida & "','" & AcaRama_Resinado_Secado_Presion & "','" & _
        AcaRama_Resinado_Secado_Ancho_Salida & "','" & AcaRama_Resinado_Secado_Densidad_Salida & "','" & AcaRama_Resinado_Secado_Vapor & "','" & AcaRama_Termo_Secado_Ancho_Entrada & "','" & _
        AcaRama_Termo_Secado_Densidad_Entrada & "','" & AcaRama_Termo_Ancho_Cadena & "','" & AcaRama_Termo_Secado_Velocidad & "','" & AcaRama_Termo_Secado_Temperatura & "','" & _
        AcaRama_Termo_Secado_SobreAlimentacion_Sup & "','" & AcaRama_Termo_Secado_SobreAlimentacion_Inf & "','" & AcaRama_Termo_Secado_SobreAlimentacion_Salida & "','" & AcaRama_Termo_Secado_Presion & "','" & _
        AcaRama_Termo_Secado_Ancho_Salida & "','" & AcaRama_Termo_Secado_Densidad_Salida & "','" & AcaRama_Termo_Secado_Vapor & "','" & Hidro_Ancho_Entrada & "','" & _
        Hidro_Ancho_Salida & "','" & Hidro_Velocidad & "','" & Hidro_Presion1 & "','" & Hidro_Presion2 & "','" & Hidro_Alimentacion_Alta & "','" & _
        Hidro_Alimentacion_Media & "','" & Hidro_Alimentacion_Baja & "','" & Hidro_Ensanchador & "','" & Seca_Ancho_Entrada & "','" & Seca_Ancho_Salida & "','" & _
        Seca_SobreAlimentacion & "','" & Seca_Temp1 & "','" & Seca_Temp2 & "','" & Seca_Temp3 & "','" & Seca_Velocidad & "','" & Seca_Encogimiento & "','" & _
        Seca_Densidad & "','" & Compa_Ancho_Entrada & "','" & Compa_Ancho_Salida & "','" & Compa_Ensanchador & "','" & Compa_Temperatura & "','" & Compa_Velocidad & "','" & _
        Compa_Teflon & "','" & Compa_Tension & "','" & Compa_Densidad & "','" & Compa_Vapor & "','" & Perc_Ancho_Entrada & "','" & Perc_Ancho_Salida & "','" & Perc_Pelo & "','" & _
        Perc_Contra_Pelo & "','" & Perc_Pases & "','" & Perc_Velocidad & "','" & Perc_Presion & "','" & Perc_Alimentacion_rodillo & "','" & Prue_Seca_Ancho_Tela & "','" & Prue_Seca_Peso_BW & "','" & _
        Prue_Seca_Peso_AW & "','" & Prue_Seca_Encog_Ancho & "','" & Prue_Seca_Encog_Largo & "','" & Prue_Seca_Revirado & "','" & Prue_BoilOff_Ancho_Tela & "','" & Prue_BoilOff_Peso_BW & "','" & _
        Prue_BoilOff_Peso_AW & "','" & Prue_BoilOff_Encog_Ancho & "','" & Prue_BoilOff_Encog_Largo & "','" & Prue_BoilOff_Revirado & "','" & Observaciones_Relevantes & "','" & Observaciones_Considerables & "','" & _
        Test_Tam_Ancho_Lavado & "','" & Test_Tam_Densidad & "','" & Test_Tam_Ancho_Proyecta & "','" & vTipo_Lavado & "'"
            
ExecuteCommandSQL cCONNECT, strSQL
MsgBox "Se grabó correctamente", vbInformation
Unload Me

Exit Sub
errGrabar:
    ErrorHandler Err, "Grabar"
End Sub

Public Sub BuscaMaquina()
On Error GoTo Fin
    
    strSQL = "ti_muestra_maquinas_propia '01'"
    TxtCod_Maquina_Tinto = ""
    TxtDes_Maquina_Tinto = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        .DGridLista.Columns("Codigo").Width = 1000
        .DGridLista.Columns("Nombre").Width = 5000
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        ElseIf rstAux.RecordCount = 1 Then
            Codigo = Trim(rstAux!Codigo)
            Descripcion = Trim(rstAux!Nombre)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Maquina_Tinto = Codigo
            TxtDes_Maquina_Tinto = Descripcion
            TxtRel_Banos_Litro.SetFocus
        End If
    End With
    Unload frmBusqGeneral3
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
    On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
    MsgBox Err.Description, vbCritical + vbOKOnly, _
    "Busqueda de Maquina "
End Sub


Private Sub TxtAbr_Alimentacion_GotFocus()
SelectionText TxtAbr_Alimentacion
End Sub

Private Sub TxtAbr_Alimentacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Alt_Destorcedor_GotFocus()
SelectionText TxtAbr_Alt_Destorcedor
End Sub

Private Sub TxtAbr_Alt_Destorcedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Ancho_Salida_GotFocus()
SelectionText TxtAbr_Ancho_Salida
End Sub

Private Sub TxtAbr_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Ancho_Tubular_GotFocus()
SelectionText TxtAbr_Ancho_Tubular
End Sub

Private Sub TxtAbr_Ancho_Tubular_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Circunf_Canasta_GotFocus()
SelectionText TxtAbr_Circunf_Canasta
End Sub

Private Sub TxtAbr_Circunf_Canasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Estiramiento_GotFocus()
SelectionText TxtAbr_Estiramiento
End Sub

Private Sub TxtAbr_Estiramiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Presion_Exprimido_GotFocus()
SelectionText TxtAbr_Presion_Exprimido
End Sub

Private Sub TxtAbr_Presion_Exprimido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtAbr_Velocidad_GotFocus()
SelectionText TxtAbr_Velocidad
End Sub

Private Sub TxtAbr_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCod_Maquina_Tinto_GotFocus()
SelectionText TxtCod_Maquina_Tinto
End Sub

Private Sub TxtCod_Maquina_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaMaquina
    End If
End Sub

Private Sub TxtCod_Proveedor_GotFocus()
SelectionText TxtCod_Proveedor
End Sub

Private Sub txtCod_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaProveedor(1)
End If
End Sub

Private Sub TxtCod_Receta_GotFocus()
SelectionText TxtCod_Receta
End Sub

Private Sub TxtCod_Receta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Receta(1)
End If
End Sub

Private Sub TxtCod_TipoReceta_GotFocus()
SelectionText TxtCod_TipoReceta
End Sub

Private Sub TxtCod_TipoReceta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaTipo_Receta(1)
End If
End Sub

Private Sub TxtCompa_Ancho_Entrada_GotFocus()
SelectionText TxtCompa_Ancho_Entrada
End Sub

Private Sub TxtCompa_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Ancho_Salida_GotFocus()
SelectionText TxtCompa_Ancho_Salida
End Sub

Private Sub TxtCompa_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Densidad_GotFocus()
SelectionText TxtCompa_Densidad
End Sub

Private Sub TxtCompa_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Ensanchador_GotFocus()
SelectionText TxtCompa_Ensanchador
End Sub

Private Sub TxtCompa_Ensanchador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Teflon_GotFocus()
SelectionText TxtCompa_Teflon
End Sub

Private Sub TxtCompa_Teflon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Temperatura_GotFocus()
SelectionText TxtCompa_Temperatura
End Sub

Private Sub TxtCompa_Temperatura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Tension_GotFocus()
SelectionText TxtCompa_Tension
End Sub

Private Sub TxtCompa_Tension_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCompa_Velocidad_GotFocus()
SelectionText TxtCompa_Velocidad
End Sub

Private Sub TxtCompa_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCurva_Tenido_GotFocus()
SelectionText TxtCurva_Tenido
End Sub

Private Sub TxtCurva_Tenido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtDes_Maquina_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaMaquina
    End If
End Sub

Private Sub txtDes_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaProveedor(2)
End If
End Sub

Private Sub TxtDes_Receta_GotFocus()
SelectionText TxtDes_Receta
End Sub

Private Sub TxtDes_Receta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Receta(2)
End If
End Sub

Private Sub TxtDes_TipoReceta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaTipo_Receta(2)
End If
End Sub

Private Sub TxtHidro_Alimentacion_Alta_GotFocus()
SelectionText TxtHidro_Alimentacion_Alta
End Sub

Private Sub TxtHidro_Alimentacion_Alta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Alimentacion_Baja_GotFocus()
SelectionText TxtHidro_Alimentacion_Baja
End Sub

Private Sub TxtHidro_Alimentacion_Baja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Alimentacion_Media_GotFocus()
SelectionText TxtHidro_Alimentacion_Media
End Sub

Private Sub TxtHidro_Alimentacion_Media_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Ancho_Entrada_GotFocus()
SelectionText TxtHidro_Ancho_Entrada
End Sub

Private Sub TxtHidro_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Ancho_Salida_GotFocus()
SelectionText TxtHidro_Ancho_Salida
End Sub

Private Sub TxtHidro_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Ensanchador_GotFocus()
SelectionText TxtHidro_Ensanchador
End Sub

Private Sub TxtHidro_Ensanchador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Presion1_GotFocus()
SelectionText TxtHidro_Presion1
End Sub

Private Sub TxtHidro_Presion1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Presion2_GotFocus()
SelectionText TxtHidro_Presion2
End Sub

Private Sub TxtHidro_Presion2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtHidro_Velocidad_GotFocus()
SelectionText TxtHidro_Velocidad
End Sub

Private Sub TxtHidro_Velocidad_KeyPress(KeyAscii As Integer)
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

Private Sub TxtPercha_Alim_Rodillo_GotFocus()
SelectionText TxtPercha_Alim_Rodillo
End Sub

Private Sub TxtPercha_Alim_Rodillo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Ancho_Entrada_GotFocus()
SelectionText TxtPercha_Ancho_Entrada
End Sub

Private Sub TxtPercha_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Ancho_Salida_GotFocus()
SelectionText TxtPercha_Ancho_Salida
End Sub

Private Sub TxtPercha_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Contra_Pelo_GotFocus()
SelectionText TxtPercha_Contra_Pelo
End Sub

Private Sub TxtPercha_Contra_Pelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Pases_GotFocus()
SelectionText TxtPercha_Pases
End Sub

Private Sub TxtPercha_Pases_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Pelo_GotFocus()
SelectionText TxtPercha_Pelo
End Sub

Private Sub TxtPercha_Pelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Presion_GotFocus()
SelectionText TxtPercha_Presion
End Sub

Private Sub TxtPercha_Presion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPercha_Velocidad_GotFocus()
SelectionText TxtPercha_Velocidad
End Sub

Private Sub TxtPercha_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruBoil_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Ancho_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Encog_Ancho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Encog_Largo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_PesoAW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_PesoBW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPruSeca_Revirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Ancho_Cadena_GotFocus()
SelectionText TxtRama_Ancho_Cadena
End Sub

Private Sub TxtRama_Ancho_Cadena_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Ancho_Entrada_GotFocus()
SelectionText TxtRama_Ancho_Entrada
End Sub

Private Sub TxtRama_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Ancho_Salida_GotFocus()
SelectionText TxtRama_Ancho_Salida
End Sub

Private Sub TxtRama_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Densidad_GotFocus()
SelectionText TxtRama_Densidad
End Sub

Private Sub TxtRama_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Densidad_Salida_GotFocus()
SelectionText TxtRama_Densidad_Salida
End Sub

Private Sub TxtRama_Densidad_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Presion_GotFocus()
SelectionText TxtRama_Presion
End Sub

Private Sub TxtRama_Presion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Ancho_Cadena_GotFocus()
SelectionText TxtRama_Resinado_Ancho_Cadena
End Sub

Private Sub TxtRama_Resinado_Ancho_Cadena_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Ancho_Entrada_GotFocus()
SelectionText TxtRama_Resinado_Ancho_Entrada
End Sub

Private Sub TxtRama_Resinado_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Ancho_Salida_GotFocus()
SelectionText TxtRama_Resinado_Ancho_Salida
End Sub

Private Sub TxtRama_Resinado_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Densidad_GotFocus()
SelectionText TxtRama_Resinado_Densidad
End Sub

Private Sub TxtRama_Resinado_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Densidad_Salida_GotFocus()
SelectionText TxtRama_Resinado_Densidad_Salida
End Sub

Private Sub TxtRama_Resinado_Densidad_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Presion_GotFocus()
SelectionText TxtRama_Resinado_Presion
End Sub

Private Sub TxtRama_Resinado_Presion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_SobreAlimInf_GotFocus()
SelectionText TxtRama_Resinado_SobreAlimInf
End Sub

Private Sub TxtRama_Resinado_SobreAlimInf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_SobreAlimSal_GotFocus()
SelectionText TxtRama_Resinado_SobreAlimSal
End Sub

Private Sub TxtRama_Resinado_SobreAlimSal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_SobreAlimSup_GotFocus()
SelectionText TxtRama_Resinado_SobreAlimSup
End Sub

Private Sub TxtRama_Resinado_SobreAlimSup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Temperatura_GotFocus()
SelectionText TxtRama_Resinado_Temperatura
End Sub

Private Sub TxtRama_Resinado_Temperatura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Resinado_Velocidad_GotFocus()
SelectionText TxtRama_Resinado_Velocidad
End Sub

Private Sub TxtRama_Resinado_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_SobreAlimInf_GotFocus()
SelectionText TxtRama_SobreAlimInf
End Sub

Private Sub TxtRama_SobreAlimInf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_SobreAlimSal_GotFocus()
SelectionText TxtRama_SobreAlimSal
End Sub

Private Sub TxtRama_SobreAlimSal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_SobreAlimSup_GotFocus()
SelectionText TxtRama_SobreAlimSup
End Sub

Private Sub TxtRama_SobreAlimSup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Temperatura_GotFocus()
SelectionText TxtRama_Temperatura
End Sub

Private Sub TxtRama_Temperatura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Ancho_Cadena_GotFocus()
SelectionText TxtRama_Termo_Ancho_Cadena
End Sub

Private Sub TxtRama_Termo_Ancho_Cadena_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Ancho_Entrada_GotFocus()
SelectionText TxtRama_Termo_Ancho_Entrada
End Sub

Private Sub TxtRama_Termo_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Ancho_Salida_GotFocus()
SelectionText TxtRama_Termo_Ancho_Salida
End Sub

Private Sub TxtRama_Termo_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Densidad_GotFocus()
SelectionText TxtRama_Termo_Densidad
End Sub

Private Sub TxtRama_Termo_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Densidad_Salida_GotFocus()
SelectionText TxtRama_Termo_Densidad_Salida
End Sub

Private Sub TxtRama_Termo_Densidad_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Presion_GotFocus()
SelectionText TxtRama_Termo_Presion
End Sub

Private Sub TxtRama_Termo_Presion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_SobreAlimInf_GotFocus()
SelectionText TxtRama_Termo_SobreAlimInf
End Sub

Private Sub TxtRama_Termo_SobreAlimInf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_SobreAlimSal_GotFocus()
SelectionText TxtRama_Termo_SobreAlimSal
End Sub

Private Sub TxtRama_Termo_SobreAlimSal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_SobreAlimSup_GotFocus()
SelectionText TxtRama_Termo_SobreAlimSup
End Sub

Private Sub TxtRama_Termo_SobreAlimSup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Temperatura_GotFocus()
SelectionText TxtRama_Termo_Temperatura
End Sub

Private Sub TxtRama_Termo_Temperatura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Termo_Velocidad_GotFocus()
SelectionText TxtRama_Termo_Velocidad
End Sub

Private Sub TxtRama_Termo_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRama_Velocidad_GotFocus()
SelectionText TxtRama_Velocidad
End Sub

Private Sub TxtRama_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRel_Banos_Kilos_GotFocus()
SelectionText TxtRel_Banos_Kilos
End Sub

Private Sub TxtRel_Banos_Kilos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtRel_Banos_Kilos) = "" Then TxtRel_Banos_Kilos.Text = "0"
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRel_Banos_Litro_GotFocus()
SelectionText TxtRel_Banos_Litro
End Sub

Private Sub TxtRel_Banos_Litro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtRel_Banos_Litro) = "" Then TxtRel_Banos_Litro.Text = "0"
    SendKeys "{TAB}"
End If
End Sub

Public Sub BuscaTipo_Receta(tipo As Integer)
On Error GoTo Fin
    
    If tipo = 1 Then
        strSQL = "SELECT Cod_TipoReceta as Codigo, Des_TipoReceta as Nombre FROM TI_TIPO_RECETA WHERE flg_operativo ='S' and Cod_TipoReceta like '%" & Trim(TxtCod_TipoReceta) & "%'"
    Else
        strSQL = "SELECT Cod_TipoReceta as Codigo, Des_TipoReceta as Nombre FROM TI_TIPO_RECETA WHERE flg_operativo ='S' and Des_TipoReceta like '%" & Trim(TxtDes_TipoReceta) & "%'"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        .DGridLista.Columns("Codigo").Width = 1000
        .DGridLista.Columns("Nombre").Width = 5000
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        ElseIf rstAux.RecordCount = 1 Then
            Codigo = Trim(rstAux!Codigo)
            Descripcion = Trim(rstAux!Nombre)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_TipoReceta = Codigo
            TxtDes_TipoReceta = Descripcion
            TxtCod_Receta.SetFocus
        End If
    End With
    Unload frmBusqGeneral3
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
    On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
    MsgBox Err.Description, vbCritical + vbOKOnly, _
    "Busqueda de Tipo Receta "
End Sub

Public Sub BuscaProveedor(tipo As Integer)
On Error GoTo Fin
    
    If tipo = 1 Then
        strSQL = "SELECT Cod_Proveedor as Codigo, Des_Proveedor as Nombre FROM tx_proveedor WHERE Cod_proveedor like '%" & Trim(TxtCod_Proveedor) & "%'"
    Else
        strSQL = "SELECT Cod_Proveedor as Codigo, Des_Proveedor as Nombre FROM tx_proveedor WHERE Des_proveedor like '%" & Trim(TxtDes_Proveedor) & "%'"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        .DGridLista.Columns("Codigo").Width = 1200
        .DGridLista.Columns("Nombre").Width = 5000
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        ElseIf rstAux.RecordCount = 1 Then
            Codigo = Trim(rstAux!Codigo)
            Descripcion = Trim(rstAux!Nombre)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Proveedor = Codigo
            TxtDes_Proveedor = Trim(Descripcion)
            TxtAbr_Velocidad.SetFocus
        End If
    End With
    Unload frmBusqGeneral3
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
    On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
    MsgBox Err.Description, vbCritical + vbOKOnly, _
    "Busqueda de Tipo Receta "
End Sub

Sub CARGA_COMBOS()
CmbCompa_Vapor.Clear
CmbCompa_Vapor.AddItem ("N")
CmbCompa_Vapor.AddItem ("S")
CmbCompa_Vapor.ListIndex = 0

cmbRama_Vapor.Clear
cmbRama_Vapor.AddItem ("N")
cmbRama_Vapor.AddItem ("S")
cmbRama_Vapor.ListIndex = 0

cmbRama_Resinado_Vapor.Clear
cmbRama_Resinado_Vapor.AddItem ("S-SUAVIZADO")
cmbRama_Resinado_Vapor.AddItem ("I-SILICONADO")
cmbRama_Resinado_Vapor.AddItem ("R-RESINADO")
cmbRama_Resinado_Vapor.ListIndex = -1

cmbRama_Termo_Vapor.Clear
cmbRama_Termo_Vapor.AddItem ("N")
cmbRama_Termo_Vapor.AddItem ("S")
cmbRama_Termo_Vapor.ListIndex = 0

CmbPrenda.Clear
CmbPrenda.AddItem ("N")
CmbPrenda.AddItem ("S")
CmbPrenda.ListIndex = 0

End Sub

Private Sub TxtSec_Ancho_Entrada_GotFocus()
SelectionText TxtSec_Ancho_Entrada
End Sub

Private Sub TxtSec_Ancho_Entrada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Ancho_Salida_GotFocus()
SelectionText TxtSec_Ancho_Salida
End Sub

Private Sub TxtSec_Ancho_Salida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Densidad_GotFocus()
SelectionText TxtSec_Densidad
End Sub

Private Sub TxtSec_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Encogimiento_GotFocus()
SelectionText TxtSec_Encogimiento
End Sub

Private Sub TxtSec_Encogimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_SobreAlimentacion_GotFocus()
SelectionText TxtSec_SobreAlimentacion
End Sub

Private Sub TxtSec_SobreAlimentacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Temp1_GotFocus()
SelectionText TxtSec_Temp1
End Sub

Private Sub TxtSec_Temp2_GotFocus()
SelectionText TxtSec_Temp2
End Sub

Private Sub TxtSec_Temp3_GotFocus()
SelectionText TxtSec_Temp3
End Sub


Private Sub TxtSec_Temp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Temp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Temp3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSec_Velocidad_GotFocus()
SelectionText TxtSec_Velocidad
End Sub

Private Sub TxtSec_Velocidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Ancho_Lavado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Ancho_Proyectado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtTambler_Densidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Public Sub Busca_Receta(tipo As Integer)
On Error GoTo Fin
    
    If tipo = 1 Then
        strSQL = "SELECT Cod_Receta as Codigo, Descripcion FROM lv_recetas WHERE Cod_Receta like '%" & Trim(TxtCod_Receta.Text) & "%'"
    Else
        strSQL = "SELECT Cod_Receta as Codigo, Descripcion FROM lv_recetas WHERE descripcion like '%" & Trim(TxtDes_Receta.Text) & "%'"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        .DGridLista.Columns("Codigo").Width = 1000
        .DGridLista.Columns("Descripcion").Width = 5000
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        ElseIf rstAux.RecordCount = 1 Then
            Codigo = Trim(rstAux!Codigo)
            Descripcion = Trim(rstAux!Descripcion)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Receta = Codigo
            TxtDes_Receta = Descripcion
            TxtCod_Proveedor.SetFocus
        End If
    End With
    Unload frmBusqGeneral3
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
    On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
    MsgBox Err.Description, vbCritical + vbOKOnly, _
    "Busqueda de Tipo Receta "
End Sub

