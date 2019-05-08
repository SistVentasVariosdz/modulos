VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPackingAtrasados 
   Caption         =   "Packing con Atraso de Facturación "
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   11271
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmPackingAtrasados.frx":0000
      Column(2)       =   "frmPackingAtrasados.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmPackingAtrasados.frx":016C
      FormatStyle(2)  =   "frmPackingAtrasados.frx":02A4
      FormatStyle(3)  =   "frmPackingAtrasados.frx":0354
      FormatStyle(4)  =   "frmPackingAtrasados.frx":0408
      FormatStyle(5)  =   "frmPackingAtrasados.frx":04E0
      FormatStyle(6)  =   "frmPackingAtrasados.frx":0598
      FormatStyle(7)  =   "frmPackingAtrasados.frx":0678
      FormatStyle(8)  =   "frmPackingAtrasados.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmPackingAtrasados.frx":07D4
   End
End
Attribute VB_Name = "frmPackingAtrasados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_almacen As String


Public Sub Buscar()

On Error GoTo drDepurar

Dim ssql As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle


ssql = "VENTAS_MUESTRA_PACKING_PENDIENTES_FACTURAR '" & sCod_almacen & "'"
GridEX1.ClearFields


Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)
  
'

GridEX1.Columns("NOM_CLIENTE").Width = 4000
GridEX1.Columns("ABR_CLIENTE").Width = 1000
GridEX1.Columns("NUM_PACKING").Width = 1000
GridEX1.Columns("NUM_MOVSTK").Width = 700
GridEX1.Columns("Fec_movstk").Width = 1000
GridEX1.Columns("PRENDAS").Width = 1500

GridEX1.Columns("NOM_CLIENTE").Caption = "Cliente"
GridEX1.Columns("ABR_CLIENTE").Caption = "Abreviatura"
GridEX1.Columns("NUM_PACKING").Caption = "Packing"
GridEX1.Columns("NUM_MOVSTK").Caption = "Movimiento"
GridEX1.Columns("Fec_movstk").Caption = "Fecha"
GridEX1.Columns("PRENDAS").Caption = "Prendas"

GridEX1.ContinuousScroll = True

Exit Sub

drDepurar:
  errores err.Number
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oFrm As New Frm_Toolbar
oFrm.CambiarContenedor Me
Set oFrm = Nothing

End Sub
