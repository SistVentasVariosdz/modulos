VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_StockCriticos 
   Caption         =   "Stocks Criticos"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   15465
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   12840
      TabIndex        =   2
      Top             =   8040
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_StockCriticos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin GridEX20.GridEX GridEX1 
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   13573
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Frm_StockCriticos.frx":0099
         Column(2)       =   "Frm_StockCriticos.frx":0161
         FormatStylesCount=   6
         FormatStyle(1)  =   "Frm_StockCriticos.frx":0205
         FormatStyle(2)  =   "Frm_StockCriticos.frx":033D
         FormatStyle(3)  =   "Frm_StockCriticos.frx":03ED
         FormatStyle(4)  =   "Frm_StockCriticos.frx":04A1
         FormatStyle(5)  =   "Frm_StockCriticos.frx":0579
         FormatStyle(6)  =   "Frm_StockCriticos.frx":0631
         ImageCount      =   0
         PrinterProperties=   "Frm_StockCriticos.frx":0711
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2640
      Top             =   6960
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_StockCriticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
BUSCAR
End Sub


Sub BUSCAR()
On Error GoTo Errores
Dim sSQL As String
Dim vBookmark As Variant

sSQL = "lg_encuentra_items_stock_debajo_punto_reorden '30'"
sSQL = VBsprintf(sSQL, Scod_ordtra)

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True
 
 GridEX1.Columns("cod_item").Width = 1000
 GridEX1.Columns("cod_item").Caption = "Codigo"
 
 
 GridEX1.Columns("des_item").Width = 4000
 GridEX1.Columns("des_item").Caption = "Item"
 
 
 GridEX1.Columns("UN").Width = 500
 GridEX1.Columns("UN").Caption = "Un"
 
 
 GridEX1.Columns("Punto_Reorden").Width = 1500
 GridEX1.Columns("Punto_Reorden").Caption = "Critico"

 GridEX1.Columns("Stock").Width = 1500
 GridEX1.Columns("Stock").Caption = "Stock"
 
 GridEX1.Columns("Fec_Ult_Compra").Width = 1500
 GridEX1.Columns("Fec_Ult_Compra").Caption = "Fecha"
 

 GridEX1.Columns("Ultima_OC").Width = 1000
 GridEX1.Columns("Ultima_OC").Caption = "O/C"
 
 GridEX1.Columns("Proveedor").Width = 3500
 GridEX1.Columns("Proveedor").Caption = "Proveedor"


GridEX1.FrozenColumns = 2

Exit Sub

Errores:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "CANCELAR"
        Unload Me
End Select
End Sub



Sub Reporte()
On Error GoTo ErrorImpresion

    Screen.MousePointer = 11
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\Rpt_Quimicos_Faltantes.xlt"
    oo.Visible = True
    
    oo.Run "REPORTE", GridEX1.ADORecordset
    
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = 0
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub
