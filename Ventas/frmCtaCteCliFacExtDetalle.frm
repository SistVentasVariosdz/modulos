VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCtaCteCliFacExtDetalle 
   Caption         =   "Detalle Cancelacion por Factoring Exterior"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   6165
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmCtaCteCliFacExtDetalle.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmCtaCteCliFacExtDetalle.frx":0352
      Column(2)       =   "frmCtaCteCliFacExtDetalle.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCtaCteCliFacExtDetalle.frx":04BE
      FormatStyle(2)  =   "frmCtaCteCliFacExtDetalle.frx":05F6
      FormatStyle(3)  =   "frmCtaCteCliFacExtDetalle.frx":06A6
      FormatStyle(4)  =   "frmCtaCteCliFacExtDetalle.frx":075A
      FormatStyle(5)  =   "frmCtaCteCliFacExtDetalle.frx":0832
      FormatStyle(6)  =   "frmCtaCteCliFacExtDetalle.frx":08EA
      FormatStyle(7)  =   "frmCtaCteCliFacExtDetalle.frx":09CA
      FormatStyle(8)  =   "frmCtaCteCliFacExtDetalle.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmCtaCteCliFacExtDetalle.frx":12CE
      PrinterProperties=   "frmCtaCteCliFacExtDetalle.frx":1620
   End
End
Attribute VB_Name = "frmCtaCteCliFacExtDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String
Public sNumCorre As String

Public Function Buscar() As Boolean
On Error GoTo errores

strSQL = "CN_VENTAS_DETALLE_CANCELACION_POR_FACTORING_EXTERIOR '" & sNumCorre & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
configura
'GridEX1.FrozenColumns = 3

Exit Function
errores:
    errores err.Number
End Function


Public Sub configura()

GridEX1.Columns("fec_cancelacion").Width = 1100
GridEX1.Columns("IMP_FACTURA_CANCELADO_A_BANCO").Width = 1200
GridEX1.Columns("num_cuota").Width = 800
GridEX1.Columns("sec_Pago").Width = 800
GridEX1.Columns("Tipo_pago").Width = 3000
GridEX1.Columns("NUM_DOCUM_OTROS").Width = 1500
GridEX1.Columns("Fec_Desembolso").Width = 1100
GridEX1.Columns("IMP_FACTURA_NEGOCIADO").Width = 1200

End Sub


