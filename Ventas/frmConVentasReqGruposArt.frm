VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConVentasReqGruposArt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6780
   ClientLeft      =   300
   ClientTop       =   735
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   12345
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4680
      TabIndex        =   1
      Top             =   6240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmConVentasReqGruposArt.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   10821
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmConVentasReqGruposArt.frx":0090
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmConVentasReqGruposArt.frx":03E2
      Column(2)       =   "frmConVentasReqGruposArt.frx":04AA
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasReqGruposArt.frx":054E
      FormatStyle(2)  =   "frmConVentasReqGruposArt.frx":0686
      FormatStyle(3)  =   "frmConVentasReqGruposArt.frx":0736
      FormatStyle(4)  =   "frmConVentasReqGruposArt.frx":07EA
      FormatStyle(5)  =   "frmConVentasReqGruposArt.frx":08C2
      FormatStyle(6)  =   "frmConVentasReqGruposArt.frx":097A
      FormatStyle(7)  =   "frmConVentasReqGruposArt.frx":0A5A
      FormatStyle(8)  =   "frmConVentasReqGruposArt.frx":0F12
      ImageCount      =   1
      ImagePicture(1) =   "frmConVentasReqGruposArt.frx":135E
      PrinterProperties=   "frmConVentasReqGruposArt.frx":16B0
   End
End
Attribute VB_Name = "frmConVentasReqGruposArt"
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

strSql = "Ventas_Muestra_Segun_Requerimiento_Grupos_Art " & strCond

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
GridEX1.Columns("Cod_Articulo").Width = 1200
GridEX1.Columns("Des_Art").Width = 5400
GridEX1.Columns("Cantidad").Width = 1110
GridEX1.Columns("Porcentaje").Width = 900
GridEX1.Columns("Importe_Soles").Width = 1410
GridEX1.Columns("Importe_Soles").Caption = "Valor Venta Soles"
GridEX1.Columns("Importe_Dolares").Width = 1560
GridEX1.Columns("Importe_Dolares").Caption = "Valor Venta Dolares"
GridEX1.Columns("Tipo").Visible = False
GridEX1.Columns("Cantidad").Format = "###,###.00"
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Porcentaje").Format = "###,###.0000"

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0

Exit Function
errores:
    errores Err.Number
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

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\ReporteRankingProductos.xlt"

    oo.Run "REPORTE", GridEX1.ADORecordset, "RANKING DE ARTICULOS DE " & Me.Caption
    
    oo.Visible = True
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    
    Exit Sub
    Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error Err.Number
End Sub

