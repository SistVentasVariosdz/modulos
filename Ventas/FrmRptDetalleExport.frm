VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptDetalleExport 
   Caption         =   "Emision Ventas Exportacion Agrupado"
   ClientHeight    =   6825
   ClientLeft      =   405
   ClientTop       =   1110
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   7680
         TabIndex        =   14
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton optFactura 
            Caption         =   "&Agrupado Factura"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optItem 
            Caption         =   "&Agrupado  Item"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.OptionButton optRangoFechas 
         Caption         =   "&Rango de Fechas"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optMensual 
         Caption         =   "&Mensual"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame frMensual 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   2775
         Begin MSComCtl2.DTPicker DTAnoMes 
            Height          =   330
            Left            =   960
            TabIndex        =   2
            Top             =   120
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "MM / yyyy"
            Format          =   94109699
            CurrentDate     =   37987
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año/Mes :"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   195
            Width           =   750
         End
      End
      Begin VB.Frame frRangoFecha 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   750
            TabIndex        =   0
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   3030
            TabIndex        =   1
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   37543
         End
         Begin VB.Label Label1 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   2400
            TabIndex        =   8
            Top             =   270
            Width           =   615
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   10440
         TabIndex        =   3
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   630
      Left            =   2445
      TabIndex        =   4
      Top             =   6120
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   1111
      Custom          =   $"FrmRptDetalleExport.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   600
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4980
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8784
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmRptDetalleExport.frx":01C0
      Column(2)       =   "FrmRptDetalleExport.frx":0288
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptDetalleExport.frx":032C
      FormatStyle(2)  =   "FrmRptDetalleExport.frx":0464
      FormatStyle(3)  =   "FrmRptDetalleExport.frx":0514
      FormatStyle(4)  =   "FrmRptDetalleExport.frx":05C8
      FormatStyle(5)  =   "FrmRptDetalleExport.frx":06A0
      FormatStyle(6)  =   "FrmRptDetalleExport.frx":0758
      FormatStyle(7)  =   "FrmRptDetalleExport.frx":0838
      FormatStyle(8)  =   "FrmRptDetalleExport.frx":08E4
      ImageCount      =   0
      PrinterProperties=   "FrmRptDetalleExport.frx":0994
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptDetalleExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCod_Anxo As String
Public codigo As String, Descripcion As String
Dim strSQL As String

Private Sub dtpFecEmiIni_Change()
  dtpFecEmiFin = dtpFecEmiIni
End Sub

Private Sub Form_Load()
  DTAnoMes = Date
  
  dtpFecEmiIni = Date
  dtpFecEmiFin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "DUA"
  Call Reporte_DUA
Case "IMPRIMIR"
  Call Reporte
Case "CONFRESU"
    Call Reporte_ConfeResum("R")
Case "DESPPOST"
    Call Reporte_ConfeResum("P")
Case "SALIR"
  Unload Me
End Select
End Sub

Private Function Reporte_ConfeResum(ByVal sTipo As String)
Dim sEmpresa As String
Dim uu As Object
Dim Ado As Object
Set Ado = CreateObject("ADODB.Recordset}")
Dim dFecIni As String, dFecFin As String
Dim sTitulo As String
On Error GoTo Fall
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA = '" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
    
If optMensual Then
  dFecIni = CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes))
  dFecFin = DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT)
Else
  dFecIni = dtpFecEmiIni
  dFecFin = dtpFecEmiFin
End If

If sTipo = "R" Then
    sTitulo = "VENTAS POR ARTICULO EXPORTACION - CONFECCIONES RESUMIDO"
    strSQL = "Ventas_Muestra_Resumido_Exportacion_Confecciones_OP '" & dFecIni & "','" & dFecFin & "'"
Else
    sTitulo = "VENTAS POR ARTICULO EXPORTACION - CONFECCIONES RESUMIDO DESPACHO POSTERIOR"
    strSQL = "Ventas_Muestra_Resumido_Exportacion_Confecciones_OP_Despacho_posterior '" & dFecIni & "','" & dFecFin & "'"
End If
Set Ado = CargarRecordSetDesconectado(strSQL, cCONNECT)

'If GridEX1.RowCount = 0 Then Exit Function

Set uu = CreateObject("excel.application")
    uu.Workbooks.Open vRuta & "\RptAgruDetExpt_ConfResum.XLT"
    uu.Visible = True
    uu.displayalerts = False
    uu.Run "Reporte", Ado, sEmpresa, sTitulo
    Set uu = Nothing
Exit Function
Fall:
MsgBox err.Description
End Function

Sub CARGA_GRID()

Dim oGroup As GridEX20.JSGroup
Dim dFecIni As String, dFecFin As String

On Error GoTo errCarga

If optMensual Then
  dFecIni = CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes))
  dFecFin = DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT)
Else
  dFecIni = dtpFecEmiIni
  dFecFin = dtpFecEmiFin
End If

strSQL = "Ventas_Muestra_Detallado_Exportacion '" & dFecIni & "','" & dFecFin & "','" & IIf(OptFactura, "X", "") & "'"
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

'GridEX1.ColumnHeaderHeight = 500

Set oGroup = gridex1.Groups.Add(gridex1.Columns("Cod_Grupo").Index, jgexSortAscending)

gridex1.Columns("Cod_Grupo").Visible = False
gridex1.Columns("Des_Producto").Visible = False

gridex1.Columns("Codigo").Width = 1140
gridex1.Columns("Factura").Width = 795
gridex1.Columns("Cantidad").Width = 765
gridex1.Columns("Fecha").Width = 945
gridex1.Columns("Precio").Width = 585
gridex1.Columns("Fob_USD").Width = 840
gridex1.Columns("Fle_USD").Width = 840
gridex1.Columns("Seg_USD").Width = 840
gridex1.Columns("Cif_USD").Width = 840
gridex1.Columns("Tc_Fob").Width = 840
gridex1.Columns("Fob_SOL").Width = 840
gridex1.Columns("Fle_SOL").Width = 840
gridex1.Columns("Seg_SOL").Width = 840
gridex1.Columns("Cif_SOL").Width = 840

gridex1.DefaultGroupMode = jgexDGMExpanded

gridex1.BackColorRowGroup = &H80000005

MuestraSubTotales



Exit Sub
Resume
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Private Sub MuestraSubTotales()

Dim colTemp As JSColumn

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Pais")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL"

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Cantidad")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Fob_USD")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Fle_USD")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Seg_USD")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Cif_USD")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Seg_USD")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Fob_SOL")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Fle_SOL")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Seg_SOL")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Cif_SOL")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""
End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object, lvTitulo As String, bItem As Boolean, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim dFecIni As String, dFecFin As String
Dim strSQL As String
Dim sEmpresa As String

    strSQL = "SELECT DES_COMP_EMP FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If optMensual Then
  dFecIni = CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes))
  dFecFin = DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT)
Else
  dFecIni = dtpFecEmiIni
  dFecFin = dtpFecEmiFin
End If

If gridex1.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")

If optItem Then
  lvTitulo = "VENTAS POR ARTICULO EXPORTACION "
  bItem = True
Else
  lvTitulo = "VENTAS POR FACTURA EXPORTACION "
   bItem = False
End If

If optMensual Then
 lvTitulo = lvTitulo + UCase(" mes de " & Format(DTAnoMes, "mmmm"))
Else
 lvTitulo = UCase(" desde el " & dtpFecEmiIni & " hasta el " & dtpFecEmiFin)
End If

Screen.MousePointer = vbHourglass

strSQL = "Ventas_Muestra_Detallado_Exportacion '" & dFecIni & "','" & dFecFin & "','" & IIf(OptFactura, "X", "") & "'"
Set RS = CargarRecordSetDesconectado(strSQL, cCONNECT)

If Not (RS.BOF Or RS.EOF) Then
  oo.Workbooks.Open vRuta & "\RptAgrupadoDetalladoExportacion.XLT"
  oo.Visible = True
  oo.displayalerts = False
  oo.Run "reporte", RS, lvTitulo, bItem, sEmpresa
  Set oo = Nothing
  Screen.MousePointer = vbDefault
Else
  Screen.MousePointer = vbDefault
  MsgBox "No hay registro ha Imprimir", vbInformation, "AVISO"
End If

Exit Sub
hand:
    Screen.MousePointer = vbDefault
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub Reporte_DUA()
On Error GoTo hand
Dim oo As Object, lvTitulo As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim dFecIni As String, dFecFin As String
Dim sRuta_Logo As String

If optMensual Then
  dFecIni = CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes))
  dFecFin = DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT)
Else
  dFecIni = dtpFecEmiIni
  dFecFin = dtpFecEmiFin
End If

Set RS = CargarRecordSetDesconectado("Ventas_Muestra_DUA '" & dFecIni & "','" & dFecFin & "'", cCONNECT)

If (RS.BOF And RS.EOF) Then
  MsgBox "NO HAY REGISTROS Q IMPRIMIR PARA ESTE PERIODO", vbInformation, "AVISO"
  Exit Sub
End If

strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo,'') FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA = '" & vemp & "'"
sRuta_Logo = DevuelveCampo(strSQL, cCONNECT)

Set oo = CreateObject("excel.application")

lvTitulo = "REPORTE DE EXPORTACION DUA"

If optMensual Then
 lvTitulo = lvTitulo + UCase(" mes de " & Format(DTAnoMes, "mmmm"))
Else
 lvTitulo = UCase(" desde el " & dtpFecEmiIni & " hasta el " & dtpFecEmiFin)
End If

oo.Workbooks.Open vRuta & "\RptExportacionDUA.XLT"
oo.Visible = True
oo.displayalerts = False
oo.Run "reporte", RS, lvTitulo, sRuta_Logo


Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub optMensual_Click()
  frMensual.Visible = True
  frRangoFecha.Visible = False
End Sub

Private Sub optRangoFechas_Click()
  frMensual.Visible = False
  frRangoFecha.Visible = True
End Sub
