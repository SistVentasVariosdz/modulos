VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConVentasReqCliGrup 
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
      Left            =   4080
      TabIndex        =   1
      Top             =   6120
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmConVentasReqCliGrup.frx":0000
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
      ImagePicture1   =   "frmConVentasReqCliGrup.frx":0090
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmConVentasReqCliGrup.frx":03E2
      Column(2)       =   "frmConVentasReqCliGrup.frx":04AA
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasReqCliGrup.frx":054E
      FormatStyle(2)  =   "frmConVentasReqCliGrup.frx":0686
      FormatStyle(3)  =   "frmConVentasReqCliGrup.frx":0736
      FormatStyle(4)  =   "frmConVentasReqCliGrup.frx":07EA
      FormatStyle(5)  =   "frmConVentasReqCliGrup.frx":08C2
      FormatStyle(6)  =   "frmConVentasReqCliGrup.frx":097A
      FormatStyle(7)  =   "frmConVentasReqCliGrup.frx":0A5A
      FormatStyle(8)  =   "frmConVentasReqCliGrup.frx":0F12
      ImageCount      =   1
      ImagePicture(1) =   "frmConVentasReqCliGrup.frx":135E
      PrinterProperties=   "frmConVentasReqCliGrup.frx":16B0
   End
End
Attribute VB_Name = "frmConVentasReqCliGrup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCond As String
Dim strSql As String

Public Function Buscar() As Boolean

On Error GoTo errores


Dim fmtCon As JSFmtCondition

strSql = "Ventas_Muestra_Segun_Requerimiento_Grupos_Clientes " & strCond

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

GridEX1.Columns("Tipo").Visible = False
GridEX1.Columns("Codigo").Width = 630
GridEX1.Columns("Nombre").Width = 4680
GridEX1.Columns("Importe_Soles").Width = 1410
GridEX1.Columns("Cantidad").Width = 1185
GridEX1.Columns("Importe_Soles").Caption = "Valor Venta Soles"
GridEX1.Columns("Importe_Dolares").Width = 1545
GridEX1.Columns("Importe_Dolares").Caption = "Valor Venta Dolares"
GridEX1.Columns("Porcentaje").Width = 960
GridEX1.Columns("Cod_Tipanex").Visible = False
GridEX1.Columns("Cod_Anxo").Visible = False
GridEX1.Columns("origen").Visible = False

GridEX1.Columns("Importe_Soles").Width = 1650
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Width = 1800
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Porcentaje").Width = 1215
GridEX1.Columns("Porcentaje").Format = "###,###.00"

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

If GridEX1.Value(GridEX1.Columns("tipo").Index) = "2" Then Exit Sub

With frmConVentasReqDoc
  Load frmConVentasReqDoc
  .Caption = "Documento de Venta del Cliente " & GridEX1.Value(GridEX1.Columns("Nombre").Index)
  .strCond = "Ventas_Muestra_Segun_Requerimiento_Grupos_Facturas_Cliente " & strCond & ",'" & GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "'"
  .Buscar
  .Show 1
End With


End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    oo.Workbooks.Open vRuta & "\ReporteRankingGuposDetalle.XLT"
    'oo.Run "REPORTE", GridEX1.ADORecordset, Me.Caption
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

