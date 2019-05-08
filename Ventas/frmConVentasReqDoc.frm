VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConVentasReqDoc 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6705
   ClientLeft      =   315
   ClientTop       =   855
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11415
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4680
      TabIndex        =   1
      Top             =   6120
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmConVentasReqDoc.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10610
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmConVentasReqDoc.frx":0090
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmConVentasReqDoc.frx":03E2
      Column(2)       =   "frmConVentasReqDoc.frx":04AA
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasReqDoc.frx":054E
      FormatStyle(2)  =   "frmConVentasReqDoc.frx":0686
      FormatStyle(3)  =   "frmConVentasReqDoc.frx":0736
      FormatStyle(4)  =   "frmConVentasReqDoc.frx":07EA
      FormatStyle(5)  =   "frmConVentasReqDoc.frx":08C2
      FormatStyle(6)  =   "frmConVentasReqDoc.frx":097A
      FormatStyle(7)  =   "frmConVentasReqDoc.frx":0A5A
      FormatStyle(8)  =   "frmConVentasReqDoc.frx":0F12
      ImageCount      =   1
      ImagePicture(1) =   "frmConVentasReqDoc.frx":135E
      PrinterProperties=   "frmConVentasReqDoc.frx":16B0
   End
End
Attribute VB_Name = "frmConVentasReqDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCond As String

Public Function Buscar() As Boolean

On Error GoTo errores

Dim strSql As String
Dim fmtCon As JSFmtCondition

strSql = strCond

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

GridEX1.Columns("Nro_Documento").Width = 1500
GridEX1.Columns("Fecha_Emision").Width = 1215
GridEX1.Columns("Cod_Moneda").Width = 570
GridEX1.Columns("Tipo_Cambio").Width = 1065
GridEX1.Columns("Imp_Gastos_Financieros").Width = 840
GridEX1.Columns("Imp_Neto").Format = "###,###.00"
GridEX1.Columns("Imp_Neto").Width = 1080
GridEX1.Columns("Imp_IGV").Width = 840
GridEX1.Columns("Imp_IGV").Format = "###,###.00"
GridEX1.Columns("Imp_Total").Width = 930
GridEX1.Columns("Imp_Total").Format = "###,###.00"
GridEX1.Columns("Guias").Width = 1620
GridEX1.Columns("Pedidos").Width = 1200
GridEX1.Columns("Num_Corre").Visible = False
GridEX1.Columns("Tipo").Visible = False
GridEX1.Columns("Flg_Por_Cobrar").Visible = False

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0

Exit Function
errores:
    errores err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "IMPRIMIR"
 If GridEX1.RowCount = 0 Then Exit Sub
 Reporte
Case Is = "SALIR"
  Unload Me
End Select
End Sub

Private Sub GridEX1_DblClick()

If GridEX1.RowCount = 0 Then Exit Sub

If GridEX1.RowCount = 0 Then Exit Sub

Load frmMuestraDetalleDocumVentas
With frmMuestraDetalleDocumVentas
  .Caption = Trim(Me.Caption) & "  " & GridEX1.Value(GridEX1.Columns("Nro_Documento").Index)
  .strSql = "Ventas_Muestra_Detalle_Factura_Items '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
  .Num_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
  .Buscar
  .FunctButt1.Visible = False
  .Show 1
  Buscar
End With


End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    oo.Workbooks.Open vRuta & "\ReporteDocumentosReq.xlt"
    oo.Run "REPORTE", GridEX1.ADORecordset, Me.Caption
    
    oo.Visible = True
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    
    Exit Sub
    Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error err.Number
End Sub

