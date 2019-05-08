VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmOrdCorte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Corte"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   435
      Left            =   5865
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3645
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7050
      Begin GridEX20.GridEX gexOC 
         Height          =   3390
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   5980
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmOrdCorte.frx":0000
         Column(2)       =   "frmOrdCorte.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmOrdCorte.frx":016C
         FormatStyle(2)  =   "frmOrdCorte.frx":02A4
         FormatStyle(3)  =   "frmOrdCorte.frx":0354
         FormatStyle(4)  =   "frmOrdCorte.frx":0408
         FormatStyle(5)  =   "frmOrdCorte.frx":04E0
         FormatStyle(6)  =   "frmOrdCorte.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmOrdCorte.frx":0678
      End
   End
End
Attribute VB_Name = "frmOrdCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String
Public sCod_OrdProv As String
Public sCod_Tela As String
Public sCod_Combo As String
Public scod_color As String
Public sCod_Medida As String
Public sCod_Calidad As String
Public SCOD_PROVEEDOR As String

Private Sub Command1_Click()
    Unload Me
End Sub

Sub CARGA_GRID()
On Error GoTo hand
Dim Reg As ADODB.Recordset

Set Reg = New ADODB.Recordset

Reg.CursorLocation = adUseClient
    Reg.Open "EXEC sm_stock_girado_partida_ordencorte '" & sCod_Almacen & "','" & sCod_OrdProv & "','" & sCod_Tela & "','" & sCod_Combo & "','" & scod_color & "','" & sCod_Medida & "','" & sCod_Calidad & "','" & SCOD_PROVEEDOR & "'", cConnect

Set gexOC.ADORecordset = Reg
Call ConfigurarGrid

Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
    gexOC.Columns("O/Corte").Width = 700
    gexOC.Columns("Status Orden").Width = 800
'    gexOC.Columns("O/Corte").Width = 700
'    gexOC.Columns("O/Corte").Width = 700
End Sub

