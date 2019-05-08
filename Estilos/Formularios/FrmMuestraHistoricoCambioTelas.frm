VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraHistoricoCambioTelas 
   Caption         =   "Solicitudes Aprobadas"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8625
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
         Left            =   2835
         TabIndex        =   6
         Top             =   240
         Width           =   5580
      End
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
         TabIndex        =   5
         Top             =   240
         Width           =   1905
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
         TabIndex        =   4
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   540
      Left            =   7350
      TabIndex        =   1
      Top             =   4515
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8625
      Begin GridEX20.GridEX GridEX1 
         Height          =   3240
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   8355
         _ExtentX        =   14737
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
         Column(1)       =   "FrmMuestraHistoricoCambioTelas.frx":0000
         Column(2)       =   "FrmMuestraHistoricoCambioTelas.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmMuestraHistoricoCambioTelas.frx":016C
         FormatStyle(2)  =   "FrmMuestraHistoricoCambioTelas.frx":02A4
         FormatStyle(3)  =   "FrmMuestraHistoricoCambioTelas.frx":0354
         FormatStyle(4)  =   "FrmMuestraHistoricoCambioTelas.frx":0408
         FormatStyle(5)  =   "FrmMuestraHistoricoCambioTelas.frx":04E0
         FormatStyle(6)  =   "FrmMuestraHistoricoCambioTelas.frx":0598
         FormatStyle(7)  =   "FrmMuestraHistoricoCambioTelas.frx":0678
         FormatStyle(8)  =   "FrmMuestraHistoricoCambioTelas.frx":0724
         ImageCount      =   0
         PrinterProperties=   "FrmMuestraHistoricoCambioTelas.frx":07D4
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Estado = R (Cambios Realizados); Estado = P (Cambios Pendientes)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   105
      TabIndex        =   7
      Top             =   4575
      Width           =   2430
   End
End
Attribute VB_Name = "FrmMuestraHistoricoCambioTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub Command1_Click()
Unload Me
End Sub

Sub CARGA_GRID()
strSQL = "EXEC tx_sel_TELA_SOLICITUDCAMBIOS '" & TxtCod_Tela & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
GridEX1.Columns("nro_solicitud").Width = "700"
GridEX1.Columns("nro_solicitud").Caption = "Solicitud"
GridEX1.Columns("Cod_Usuario_Permiso").Caption = "Autorizado por"
GridEX1.Columns("Cod_Usuario_Permiso").Width = "2000"
GridEX1.Columns("Cod_Usuario_modificacion").Caption = "Usuario Modificacion"
GridEX1.Columns("Cod_Usuario_modificacion").Width = "2000"
GridEX1.Columns("flg_status").Caption = "Estado"
GridEX1.Columns("flg_status").Width = "800"
GridEX1.Columns("fecha_solicitud").Width = "1300"
End Sub

