VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraBitacora 
   Caption         =   "Telas - Bitacora"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   540
      Left            =   8190
      TabIndex        =   2
      Top             =   3675
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin GridEX20.GridEX GridEX1 
         Height          =   3240
         Left            =   105
         TabIndex        =   1
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
         Column(1)       =   "FrmMuestraBitacora.frx":0000
         Column(2)       =   "FrmMuestraBitacora.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmMuestraBitacora.frx":016C
         FormatStyle(2)  =   "FrmMuestraBitacora.frx":02A4
         FormatStyle(3)  =   "FrmMuestraBitacora.frx":0354
         FormatStyle(4)  =   "FrmMuestraBitacora.frx":0408
         FormatStyle(5)  =   "FrmMuestraBitacora.frx":04E0
         FormatStyle(6)  =   "FrmMuestraBitacora.frx":0598
         FormatStyle(7)  =   "FrmMuestraBitacora.frx":0678
         FormatStyle(8)  =   "FrmMuestraBitacora.frx":0724
         ImageCount      =   0
         PrinterProperties=   "FrmMuestraBitacora.frx":07D4
      End
   End
End
Attribute VB_Name = "FrmMuestraBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public stela As String
Public ssolicitud As String

Private Sub Command1_Click()
Unload Me
End Sub

Sub CARGA_GRID()
strSQL = "EXEC tx_muestra_bitacora_hilostel '" & ssolicitud & "','" & stela & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Num_Secuencia").Width = "650"
GridEX1.Columns("Num_Secuencia").Caption = "Secuencia"

GridEX1.Columns("Hilado").Caption = "Hilado"
GridEX1.Columns("Hilado").Width = "1400"

GridEX1.Columns("Nombre").Caption = "Nombre"
GridEX1.Columns("Nombre").Width = "1400"

GridEX1.Columns("Porcentaje").Width = "1300"
GridEX1.Columns("Porcentaje").Caption = "Porcentaje"

GridEX1.Columns("Long_Malla").Width = "950"
GridEX1.Columns("Long_Malla").Caption = "Long Malla"

GridEX1.Columns("Num_Agujas").Width = "950"
GridEX1.Columns("Num_Agujas").Caption = "Num Agujas"

GridEX1.Columns("Num_Alimentadores").Width = "950"
GridEX1.Columns("Num_Alimentadores").Caption = "Num Alimentadores"

GridEX1.Columns("Parafinado").Width = "950"
GridEX1.Columns("Parafinado").Caption = "Parafinado"

GridEX1.Columns("Torsion").Width = "950"
GridEX1.Columns("Torsion").Caption = "Torsion"
End Sub


