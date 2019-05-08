VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmTelaBitacora 
   Caption         =   "Telas - Bitacora"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnhilados 
      Caption         =   "Ver Hilados"
      Height          =   540
      Left            =   6840
      TabIndex        =   7
      Top             =   4515
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   540
      Left            =   8190
      TabIndex        =   6
      Top             =   4515
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   9465
      Begin GridEX20.GridEX GridEX1 
         Height          =   3240
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   5715
         Version         =   "2.0"
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
         Column(1)       =   "FrmTelaBitacora.frx":0000
         Column(2)       =   "FrmTelaBitacora.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmTelaBitacora.frx":016C
         FormatStyle(2)  =   "FrmTelaBitacora.frx":02A4
         FormatStyle(3)  =   "FrmTelaBitacora.frx":0354
         FormatStyle(4)  =   "FrmTelaBitacora.frx":0408
         FormatStyle(5)  =   "FrmTelaBitacora.frx":04E0
         FormatStyle(6)  =   "FrmTelaBitacora.frx":0598
         FormatStyle(7)  =   "FrmTelaBitacora.frx":0678
         FormatStyle(8)  =   "FrmTelaBitacora.frx":0724
         ImageCount      =   0
         PrinterProperties=   "FrmTelaBitacora.frx":07D4
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.TextBox TxtCod_Tela 
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
         Height          =   330
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox TxtDes_Tela 
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
         Height          =   330
         Left            =   2940
         TabIndex        =   1
         Top             =   240
         Width           =   6210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
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
         Left            =   315
         TabIndex        =   3
         Top             =   315
         Width           =   390
      End
   End
End
Attribute VB_Name = "FrmTelaBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub btnhilados_Click()
            If GridEX1.RowCount = 0 Then Exit Sub
            Load FrmMuestraBitacora
            FrmMuestraBitacora.stela = Trim(txtcod_tela.Text)
            FrmMuestraBitacora.ssolicitud = GridEX1.Value(GridEX1.Columns("nro_solicitud").Index)
            FrmMuestraBitacora.CARGA_GRID
            FrmMuestraBitacora.Show 1
            Set FrmTelaBitacora = Nothing
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Sub CARGA_GRID()
strSQL = "EXEC TX_SEL_TELA_BITACORA '" & txtcod_tela & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
GridEX1.Columns("nro_solicitud").Width = "650"
GridEX1.Columns("nro_solicitud").Caption = "Solicitud"
GridEX1.Columns("Cod_Usuario_modificacion").Caption = "Usuario Modif."
GridEX1.Columns("Cod_Usuario_modificacion").Width = "1400"
GridEX1.Columns("Pc_Cambios").Width = "1400"
GridEX1.Columns("Fecha_Realizacion").Width = "1300"
GridEX1.Columns("Fecha_Realizacion").Caption = "Fecha Cambios"

GridEX1.Columns("gramaje_antiguo").Width = "950"
GridEX1.Columns("Ancho_antiguo").Width = "950"
GridEX1.Columns("gramaje_Acab").Width = "950"
GridEX1.Columns("ancho_Acab").Width = "950"

GridEX1.Columns("cod_tela").Visible = False
End Sub


