VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmDatosTecDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Tecnicos"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frmDatosTecDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   10395
      Begin VB.TextBox TxtTela 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   48
         Top             =   200
         Width           =   3495
      End
      Begin VB.TextBox txtPartida 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   46
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Tela :"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Partida :"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   3780
      TabIndex        =   37
      Top             =   5880
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1005
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~1~0~3~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   550
      ControlSeparator=   150
   End
   Begin VB.Frame Frame1 
      Height          =   5130
      Left            =   105
      TabIndex        =   38
      Top             =   630
      Width           =   10410
      Begin VB.ComboBox cboSolidez_3LV 
         Height          =   315
         Left            =   7545
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2895
         Width           =   1260
      End
      Begin VB.ComboBox cboSolidez_2LV 
         Height          =   315
         Left            =   6075
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2910
         Width           =   1260
      End
      Begin VB.ComboBox cboSolidez_1LV 
         Height          =   315
         Left            =   4635
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2910
         Width           =   1260
      End
      Begin VB.ComboBox cboSolidez_Planta 
         Height          =   315
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2910
         Width           =   1260
      End
      Begin VB.ComboBox cboSolidez_Tinto 
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2910
         Width           =   1260
      End
      Begin VB.TextBox txtAncho_Total 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9015
         TabIndex        =   63
         Text            =   "0"
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtgramaje_Acab 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9015
         TabIndex        =   62
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Ancho 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9015
         TabIndex        =   61
         Text            =   "0"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Largo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9015
         TabIndex        =   60
         Text            =   "0"
         Top             =   2265
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9015
         TabIndex        =   59
         Text            =   "0"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   33
         Text            =   "0"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Largo_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   32
         Text            =   "0"
         Top             =   2265
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Ancho_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   31
         Text            =   "0"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtgramaje_Acab_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   28
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Total_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   29
         Text            =   "0"
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Util_3LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7545
         TabIndex        =   30
         Text            =   "0"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   26
         Text            =   "0"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Largo_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   25
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Ancho_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   24
         Text            =   "0"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtgramaje_Acab_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   21
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Total_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   22
         Text            =   "0"
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Util_2LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6075
         TabIndex        =   23
         Text            =   "0"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   19
         Text            =   "0"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Largo_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   18
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Ancho_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   17
         Text            =   "0"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtgramaje_Acab_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   14
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Total_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   15
         Text            =   "0"
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Util_1LV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4635
         TabIndex        =   16
         Text            =   "0"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   12
         Text            =   "0"
         Top             =   2595
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Largo_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   11
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_Ancho_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   10
         Text            =   "0"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtgramaje_Acab_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   7
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Total_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   8
         Text            =   "0"
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Util_Planta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3180
         TabIndex        =   9
         Text            =   "0"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.TextBox txtRevirado_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Text            =   "0"
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_largo_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "0"
         Top             =   2265
         Width           =   1095
      End
      Begin VB.TextBox txtEncog_ancho_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Text            =   "0"
         Top             =   1935
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   195
         TabIndex        =   49
         Top             =   3420
         Width           =   8985
         Begin VB.ComboBox cboFlg_Status_Apro 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   315
            Width           =   1785
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   495
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   705
            Width           =   7005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   150
            TabIndex        =   51
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   150
            TabIndex        =   50
            Top             =   735
            Width           =   1110
         End
      End
      Begin VB.TextBox txtAncho_Util_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txtAncho_Total_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Text            =   "0"
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox TxtGramaje_acab_Tinto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LAVADO"
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
         Left            =   4665
         TabIndex        =   66
         Top             =   210
         Width           =   4005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ANTES DE LAVADO"
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
         Left            =   1800
         TabIndex        =   65
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label Label3 
         Caption         =   "Programado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9015
         TabIndex        =   64
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "1ra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4980
         TabIndex        =   58
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Solidez :"
         Height          =   210
         Left            =   255
         TabIndex        =   57
         Top             =   2970
         Width           =   1080
      End
      Begin VB.Label Label18 
         Caption         =   "Revirado :"
         Height          =   210
         Left            =   240
         TabIndex        =   56
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label17 
         Caption         =   "Encog. Largo :"
         Height          =   210
         Left            =   240
         TabIndex        =   55
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label16 
         Caption         =   "Encog. Ancho :"
         Height          =   210
         Left            =   240
         TabIndex        =   54
         Top             =   1980
         Width           =   1125
      End
      Begin VB.Label Label15 
         Caption         =   "3ra."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7905
         TabIndex        =   53
         Top             =   615
         Width           =   405
      End
      Begin VB.Label Label14 
         Caption         =   "Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3390
         TabIndex        =   52
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Ancho Util :"
         Height          =   210
         Left            =   240
         TabIndex        =   43
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "2da."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6450
         TabIndex        =   42
         Top             =   615
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Tintoreria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Top             =   615
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Ancho Total (Mt) :"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Gramaje (Gr/Mt2) :"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1050
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDatosTecDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_TipOrdTra As String
Public vCod_OrdTra As String
Public vCod_Tela As String
Public vDes_tela As String
Public vCod_Comb As String
Public vCod_Color As String
Public vPartida As String
Public sAccion As String
Dim strSQL As String

Private Sub cboFlg_Status_Apro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSolidez_1LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSolidez_Planta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSolidez_Tinto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSolidez_2LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSolidez_3LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Dim Rs As ADODB.Recordset, RsTelas As ADODB.Recordset
On Error GoTo hand

cboFlg_Status_Apro.Clear
cboFlg_Status_Apro.AddItem "Aprobado" & Space(100) & "A"
cboFlg_Status_Apro.AddItem "Desaprobado" & Space(100) & "D"
cboFlg_Status_Apro.ListIndex = 0

FillStatus cboFlg_Status_Apro
FillStatus cboSolidez_Tinto
FillStatus cboSolidez_Planta
FillStatus cboSolidez_1LV
FillStatus cboSolidez_2LV
FillStatus cboSolidez_3LV

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

sAccion = "I"

txtPartida.Text = Trim(vPartida)
TxtTela = Trim(vCod_Tela) & "-" & Trim(vDes_tela)

Rs.Open "SELECT * FROM TX_DATOS_TECNICOS_TELAS " & _
        "WHERE Cod_TipOrdTra = '" & vCod_TipOrdTra & "' " & _
        "AND Cod_OrdTra = '" & vCod_OrdTra & "' " & _
        "AND Cod_Tela = '" & vCod_Tela & "' " & _
        "AND Cod_comb = '" & vCod_Comb & "' " & _
        "AND Cod_Color = '" & vCod_Color & "'"
With Rs
If .RecordCount Then
    .MoveFirst
    
    sAccion = "U"
    
    TxtGramaje_acab_Tinto = !Gramaje_acab_Tinto
    txtAncho_Total_Tinto = !Ancho_Total_Tinto
    txtAncho_Util_Tinto = !Ancho_Util_Tinto
    txtEncog_ancho_Tinto = !Encog_ancho_Tinto
    txtEncog_largo_Tinto = !Encog_largo_Tinto
    txtRevirado_Tinto = !Revirado_Tinto
    BuscaCombo !Solidez_Tinto, 2, cboSolidez_Tinto
    
    txtgramaje_Acab_Planta = !gramaje_Acab_Planta
    txtAncho_Total_Planta = !Ancho_Total_Planta
    txtAncho_Util_Planta = !Ancho_Util_Planta
    txtEncog_Ancho_Planta = !Encog_Ancho_Planta
    txtEncog_Largo_Planta = !Encog_Largo_Planta
    txtRevirado_Planta = !Revirado_Planta
    BuscaCombo !Solidez_planta, 2, cboSolidez_Planta
    
    txtgramaje_Acab_1LV = !gramaje_Acab_1LV
    txtAncho_Total_1LV = !Ancho_Total_1LV
    txtAncho_Util_1LV = !Ancho_Util_1LV
    txtEncog_Ancho_1LV = !Encog_Ancho_1LV
    txtEncog_Largo_1LV = !Encog_Largo_1LV
    txtRevirado_1LV = !Revirado_1LV
    BuscaCombo !Solidez_1LV, 2, cboSolidez_1LV
    'txtSolidez_1LV = Trim(!Solidez_1LV)
    
    txtgramaje_Acab_2LV = !gramaje_Acab_2LV
    txtAncho_Total_2LV = !Ancho_Total_2LV
    txtAncho_Util_2LV = !Ancho_Util_2LV
    txtEncog_Ancho_2LV = !Encog_Ancho_2LV
    txtEncog_Largo_2LV = !Encog_Largo_2LV
    txtRevirado_2LV = !Revirado_2LV
    BuscaCombo !Solidez_2LV, 2, cboSolidez_2LV
    'txtSolidez_2LV = Trim(!Solidez_2LV)
    
    txtgramaje_Acab_3LV = !gramaje_Acab_3LV
    txtAncho_Total_3LV = !Ancho_Total_3LV
    txtAncho_Util_3LV = !Ancho_Util_3LV
    txtEncog_Ancho_3LV = !Encog_Ancho_3LV
    txtEncog_Largo_3LV = !Encog_Largo_3LV
    txtRevirado_3LV = !Revirado_3LV
    BuscaCombo !Solidez_3LV, 2, cboSolidez_3LV
    'txtSolidez_3LV = Trim(!Solidez_3LV)
    
    'Call BuscaCombo(Rs("cod_Tipenc"), 1, CboTipo)
    
    'Esto es para los datos Adicionales
'    If IsNull(Rs("Gramaje_Lavado").Value) Then
'        Me.txtGramaje_Lavado.Text = ""
'    Else
'        Me.txtGramaje_Lavado.Text = Rs("Gramaje_Lavado").Value
'    End If
    
    If IsNull(Rs("Observaciones").Value) Then
        Me.txtObservaciones.Text = ""
    Else
        Me.txtObservaciones.Text = Rs("Observaciones").Value
    End If
    
    If IsNull(Rs("Flg_Status_Apro").Value) Then
        Me.cboFlg_Status_Apro.ListIndex = -1
    Else
        Call BuscaCombo(Rs("Flg_Status_Apro").Value, 2, cboFlg_Status_Apro)
    End If
End If
.Close
End With

Set RsTelas = New ADODB.Recordset
RsTelas.Open "SELECT Gramaje_Acab, Ancho_Acab, Encog_Ancho, Encog_Largo " & _
             "FROM TX_TELA WHERE Cod_Tela = '" & vCod_Tela & "'", cConnect, _
             adOpenStatic
If RsTelas.RecordCount Then
    txtgramaje_Acab = RsTelas!gramaje_Acab
    txtAncho_Total = RsTelas!ancho_acab
    txtEncog_Ancho = RsTelas!encog_Ancho
    txtEncog_Largo = RsTelas!encog_largo
End If
RsTelas.Close

Set Rs = Nothing
Set RsTelas = Nothing
Exit Sub

hand:
Set Rs = Nothing
Set RsTelas = Nothing

ErrorHandler Err, "Form_load"
End Sub

Private Sub FillStatus(ByRef cboAux As ComboBox)
    cboAux.Clear
    cboAux.AddItem "Aprobado" & Space(100) & "A"
    cboAux.AddItem "Desaprobado" & Space(100) & "D"
    cboAux.ListIndex = 0
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        SALVAR_DATOS
    Case "CANCELAR"
        Unload Me
End Select
End Sub

Sub SALVAR_DATOS()
On Error GoTo hand

strSQL = "EXEC UP_MAN_TX_DATOS_TECNICOS_TELAS '" & sAccion & "', '" & vCod_TipOrdTra & "', '" & _
vCod_OrdTra & "', '" & vCod_Tela & "', '" & vCod_Comb & "', '" & vCod_Color & _
"', " & TxtGramaje_acab_Tinto & ", " & txtAncho_Total_Tinto & ", " & _
txtAncho_Util_Tinto & ", " & txtEncog_ancho_Tinto & ", " & txtEncog_largo_Tinto & _
", " & txtRevirado_Tinto & ", '" & Trim(Right(cboSolidez_Tinto, 3)) & "', " & _
txtgramaje_Acab_Planta & ", " & txtAncho_Total_Planta & ", " & _
txtAncho_Util_Planta & ", " & txtEncog_Ancho_Planta & ", " & txtEncog_Largo_Planta & ", " & _
txtRevirado_Planta & ", '" & Trim(Right(cboSolidez_Planta, 3)) & "', " & txtgramaje_Acab_1LV & ", " & _
txtAncho_Total_1LV & ", " & txtAncho_Util_1LV & ", " & txtEncog_Ancho_1LV & ", " & _
txtEncog_Largo_1LV & ", " & txtRevirado_1LV & ", '" & Trim(Right(cboSolidez_1LV, 3)) & "', " & _
txtgramaje_Acab_2LV & ", " & txtAncho_Total_2LV & ", " & txtAncho_Util_2LV & ", " & _
txtEncog_Ancho_2LV & ", " & txtEncog_Largo_2LV & ", " & txtRevirado_2LV & ", '" & _
Trim(Right(cboSolidez_2LV, 3)) & "', " & txtgramaje_Acab_3LV & ", " & txtAncho_Total_3LV & ", " & _
txtAncho_Util_3LV & ", " & txtEncog_Ancho_3LV & ", " & txtEncog_Largo_3LV & ", " & _
txtRevirado_3LV & ", '" & Trim(Right(cboSolidez_3LV, 3)) & "', '" & Trim(Right(cboFlg_Status_Apro, 3)) _
& "', '" & txtObservaciones & "'"

ExecuteSQL cConnect, strSQL

Unload Me

Exit Sub
hand:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Guardar Datos Tecnicos"
    'ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub txtAncho_Total_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Util_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_ancho_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_largo_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtGramaje_acab_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRevirado_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSolidez_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Total_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Util_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_ancho_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_largo_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtGramaje_acab_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRevirado_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSolidez_Planta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Total_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Util_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_ancho_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_largo_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtGramaje_acab_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRevirado_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSolidez_1LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Total_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Util_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_ancho_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_largo_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtGramaje_acab_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRevirado_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSolidez_2LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Total_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAncho_Util_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_ancho_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEncog_largo_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub TxtGramaje_acab_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRevirado_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSolidez_3LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
