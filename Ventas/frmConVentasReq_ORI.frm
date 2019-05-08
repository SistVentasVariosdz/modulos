VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmConVentasReq 
   Caption         =   "Ranking de Clientes"
   ClientHeight    =   7545
   ClientLeft      =   2580
   ClientTop       =   1290
   ClientWidth     =   12390
   Icon            =   "frmConVentasReq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   12390
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12225
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   7920
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optNacional 
         Caption         =   "Nacional"
         Height          =   195
         Left            =   6480
         TabIndex        =   16
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optanexocont 
         Caption         =   "Por Anexo Contable"
         Height          =   195
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optOrdTipoVenta 
         Caption         =   "Ord por Grupos"
         Height          =   195
         Left            =   7920
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optOrdCliente 
         Caption         =   "Ord. por Cliente"
         Height          =   195
         Left            =   6480
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame frRangoFecha 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1920
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   3975
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   80543745
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   2640
            TabIndex        =   9
            Top             =   90
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   80543745
            CurrentDate     =   37543
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta :"
            Height          =   195
            Left            =   2040
            TabIndex        =   11
            Top             =   120
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.Frame frMensual 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2250
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         Begin MSComCtl2.DTPicker DTAnoMes 
            Height          =   330
            Left            =   900
            TabIndex        =   5
            Top             =   120
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "MM / yyyy"
            Format          =   80543747
            CurrentDate     =   37987
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año/Mes :"
            Height          =   195
            Left            =   60
            TabIndex        =   6
            Top             =   195
            Width           =   750
         End
      End
      Begin VB.OptionButton optMensual 
         Caption         =   "&Mensual"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optRangoFechas 
         Caption         =   "&Rango de Fechas"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5580
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   9843
      Version         =   "2.0"
      RecordNavigator =   -1  'True
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
      Column(1)       =   "frmConVentasReq.frx":030A
      Column(2)       =   "frmConVentasReq.frx":03D2
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasReq.frx":0476
      FormatStyle(2)  =   "frmConVentasReq.frx":05AE
      FormatStyle(3)  =   "frmConVentasReq.frx":065E
      FormatStyle(4)  =   "frmConVentasReq.frx":0712
      FormatStyle(5)  =   "frmConVentasReq.frx":07EA
      FormatStyle(6)  =   "frmConVentasReq.frx":08A2
      FormatStyle(7)  =   "frmConVentasReq.frx":0982
      FormatStyle(8)  =   "frmConVentasReq.frx":0A2E
      ImageCount      =   0
      PrinterProperties=   "frmConVentasReq.frx":0ADE
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   1191
      Custom          =   $"frmConVentasReq.frx":0CB6
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   650
      ControlSeparator=   40
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10875
      Top             =   6240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmConVentasReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String
Dim strOrigen As String
Dim dFecIni As Date, dFecFin As Date
Dim strSQL As String



Private Sub dtpFecEmiIni_Change()
  dtpFecEmiFin = dtpFecEmiIni
End Sub

Private Sub Form_Load()
  DTAnoMes.Value = Date
  dtpFecEmiFin = Date
  dtpFecEmiIni = Date
  strOrigen = "N"
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub BUSCAR()

On Error GoTo drDepurar

Dim sSQL As String
Dim fmtCon As JSFmtCondition

Encuentra_Fechas



If optOrdCliente Then
  sSQL = "Ventas_Muestra_Segun_Requerimiento '" & dFecIni & "','" & dFecFin & "','" & strOrigen & "'"
  If strOrigen = "E" Then sSQL = sSQL + ",'1',''"
  
  strSQL = sSQL
  
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
  If strOrigen = "N" Or strOrigen = "" Then GridEX1.Columns("Pais").Visible = False
  
  GridEX1.Columns("Nro").Width = 390
  GridEX1.Columns("Codigo").Width = 1335
  GridEX1.Columns("Nombre").Width = 3750
  
  GridEX1.Columns("cod_tipanex").Visible = False
  GridEX1.Columns("cod_anxo").Visible = False
  GridEX1.Columns("Origen").Visible = False
  
ElseIf optanexocont Then
  sSQL = "Ventas_Muestra_Segun_Requerimiento '" & dFecIni & "','" & dFecFin & "','" & strOrigen & "'"
  If strOrigen = "E" Then
    sSQL = sSQL + ",'1','','S'"
  Else
    sSQL = sSQL + "'1','NULL','S'"
  End If
  strSQL = sSQL
  
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
  If strOrigen = "N" Or strOrigen = "" Then GridEX1.Columns("Pais").Visible = False
  
  GridEX1.Columns("Nro").Width = 390
  GridEX1.Columns("Codigo").Width = 1335
  GridEX1.Columns("Nombre").Width = 3750
  
  GridEX1.Columns("cod_tipanex").Visible = False
  GridEX1.Columns("cod_anxo").Visible = False
  GridEX1.Columns("Origen").Visible = False
  
Else

  sSQL = "Ventas_Muestra_Segun_Requerimiento_Grupos '" & dFecIni & "','" & dFecFin & "','" & strOrigen & "'"
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
  
  GridEX1.Columns("Cod_Grupo_Ventas").Visible = False
  GridEX1.Columns("Grupo").Width = 2280
  
End If

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Tipo").Visible = False

GridEX1.Columns("Importe_Soles").Width = 1185
GridEX1.Columns("Importe_Soles").Caption = "Valor Venta Soles"
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Width = 1365
GridEX1.Columns("Importe_Dolares").Caption = "Valor Venta Dolares"
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Cantidad").Width = 1020
GridEX1.Columns("Cantidad").Format = "###,###.00"

GridEX1.Columns("Porcentaje").Width = 900
GridEX1.Columns("Porcentaje").Format = "###,###.0000"

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0
  
Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub

Private Sub CONFIGURA_GRID()

On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

GridEX1.DefaultGroupMode = jgexDGMExpanded


  
Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Grupo").Index, jgexSortAscending)

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cod").Index, jgexSortAscending)

MuestraSubTotales

GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed


GridEX1.ContinuousScroll = True
  
Exit Sub

drDepurar:
  errores err.Number
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

Encuentra_Fechas

VB.Screen.MousePointer = vbHourglass

Dim oo As Object
Set oo = CreateObject("excel.application")

If optOrdCliente.Value = True Then
  oo.Workbooks.Open vRuta & "\ReporteRankingClientes.xlt"
  oo.Visible = True
  oo.Run "REPORTE", GridEX1.ADORecordset, "Ranking de Clientes desde el " & dFecIni & " Hasta el " & dFecFin, strSQL, cCONNECT, ""

ElseIf optanexocont.Value = True Then
  oo.Workbooks.Open vRuta & "\ReporteRankingClientes.xlt"
  oo.Visible = True
  oo.Run "REPORTE", GridEX1.ADORecordset, "Ranking de Clientes desde el " & dFecIni & " Hasta el " & dFecFin, strSQL, cCONNECT, "S"
  
Else
  oo.Workbooks.Open vRuta & "\ReporteRankingGupos.xlt"
  oo.Visible = True
  oo.Run "REPORTE", GridEX1.ADORecordset, "Ranking de Grupos desde el " & dFecIni & " Hasta el " & dFecFin
End If


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


Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      BUSCAR
    Case "AUTORIZARPAGO"
      CONFIGURA_GRID
    Case "DOC"
        Doc
    Case "IMPRIMIR"
        If GridEX1.RowCount = 0 Then Exit Sub
        Reporte
    Case "IMPRIMIRDAOT"
    If optOrdCliente.Value = True Then
        BuscarDAOT
    End If
        Reporte
    Case "GENERARDAOT"
        DAOT
    Case "SALIR"
       Unload Me
    End Select
End Sub


Private Sub MuestraSubTotales()
Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter

Set colTemp = GridEX1.Columns("Fecha")

colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "


GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Soles")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Dolares")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub

Private Sub SetColores()

'Dim fmtCon As JSFmtCondition
'Dim fmtCond2 As JSFmtCondition
'Dim fmtCond3 As JSFmtCondition
'
'Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("SEL").Index, jgexEqual, -1)
'
'    With GridEX1.FmtConditions
'            .ApplyGroupCondition = True
'            .ShowGroupConditionCount = True
'            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
'            Set fmtCon = .GroupCondition
'    End With
'    fmtCon.SetCondition GridEX1.Columns("SEL").Index, jgexEqual, -1
'    fmtCon.FormatStyle.FontBold = True
'    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub
Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String
    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            Select Case opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & opcion & ")"
End Sub
Sub Encuentra_Fechas()
  If optMensual Then
    dFecIni = CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes))
    dFecFin = DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT)
  Else
    dFecIni = dtpFecEmiIni
    dFecFin = dtpFecEmiFin
  End If
End Sub
Private Sub GridEX1_DblClick()

If GridEX1.RowCount = 0 Then Exit Sub

If GridEX1.Value(GridEX1.Columns("tipo").Index) = "2" Then Exit Sub

Encuentra_Fechas

If optOrdCliente Then
  With frmConVentasReqGrupos
    Load frmConVentasReqGrupos
    .Caption = GridEX1.Value(GridEX1.Columns("Nombre").Index)
    .strCond = "'" & dFecIni & "','" & dFecFin & "','" & strOrigen & "','" & GridEX1.Value(GridEX1.Columns("cod_tipanex").Index) & "','" & GridEX1.Value(GridEX1.Columns("cod_anxo").Index) & "'"
    .BUSCAR
    .Show 1
  End With
ElseIf optanexocont Then
  With frmConVentasReqCliGrup
    Load frmConVentasReqCliGrup
    .Caption = "Ranking de Clientes por Anexo Contable # " & GridEX1.Value(GridEX1.Columns("cod_anxo").Index) & " desde el " & dFecIni & " hasta el " & dFecFin
    .strCond = "'" & dFecIni & "','" & dFecFin & "','" & strOrigen & "','" & GridEX1.Value(GridEX1.Columns("cod_anxo").Index) & "'"
    .BUSCAR
    .Show 1
  End With
Else
  With frmConVentasReqCliGrup
    Load frmConVentasReqCliGrup
    .Caption = "Ranking de Clientes del Grupo " & GridEX1.Value(GridEX1.Columns("Grupo").Index) & " desde el " & dFecIni & " hasta el " & dFecFin
    .strCond = "'" & dFecIni & "','" & dFecFin & "','" & strOrigen & "','" & GridEX1.Value(GridEX1.Columns("Cod_Grupo_ventas").Index) & "'"
    .BUSCAR
    .Show 1
  End With
End If

End Sub
Sub Doc()

Encuentra_Fechas

If GridEX1.RowCount = 0 Then Exit Sub

If GridEX1.Value(GridEX1.Columns("tipo").Index) = "2" Then Exit Sub

If optOrdCliente Then
  With frmConVentasReqDoc
    Load frmConVentasReqDoc
    .Caption = UCase("Documento de Ventas del Cliente " & Trim(GridEX1.Value(GridEX1.Columns("Nombre").Index)) & " desde el " & dFecIni & " Hasta el " & dFecFin)
    .strCond = "Ventas_Muestra_Segun_Requerimiento_Facturas_Clientes '" & dFecIni & "','" & dFecFin & "','" & strOrigen & "','" & GridEX1.Value(GridEX1.Columns("cod_tipanex").Index) & "','" & GridEX1.Value(GridEX1.Columns("cod_anxo").Index) & "'"
    .BUSCAR
    .Show 1
  End With
End If

End Sub

Private Sub optExtranjero_Click()
  strOrigen = "E"
End Sub

Private Sub optMensual_Click()
  frMensual.Visible = True
  frRangoFecha.Visible = False
End Sub

Private Sub optNacional_Click()
  strOrigen = "N"
End Sub

Private Sub optRangoFechas_Click()
  frMensual.Visible = False
  frRangoFecha.Visible = True
End Sub

Private Sub optTodos_Click()
  strOrigen = ""
End Sub



Public Sub DAOT()
On Error GoTo errorx
Dim oRs As ADODB.Recordset
Dim sSQL As String
Dim cCONNECTDBF As String
Dim sDBFCONTABHIAL As String
Dim oRs2 As ADODB.Recordset

' CREAMOS CADENA DE CONEXION
sDBFCONTABHIAL = "C:\DBF\"
cCONNECTDBF = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;DEFAULTDIR= $;Database=$"
cCONNECTDBF = VBsprintf(cCONNECTDBF, sDBFCONTABHIAL, sDBFCONTABHIAL)

Encuentra_Fechas

sSQL = "Ventas_Muestra_Segun_Requerimiento '" & dFecIni & "','" & dFecFin & "','D',0"
Set oRs = GetRecordset(cCONNECT, sSQL)

Set oRs2 = oRs.Clone(adLockReadOnly)

DAOT_DETALLE oRs, cCONNECTDBF, False
DAOT_DETALLE oRs2, cCONNECTDBF, True

Exit Sub
errorx:
    If err.Number = -2147467259 Then
        Resume Next
    End If
        
    errores err.Number

End Sub

Private Sub DAOT_DETALLE(ByVal oRs As ADODB.Recordset, ByVal cCONNECTDBF As String, ByVal bMensaje As Boolean)
Dim sSQL As String
Dim oCampo As ADODB.FIELD
Dim Rs_DATOSDBF As Object
Set Rs_DATOSDBF = CreateObject("ADODB.Recordset")

' BORRAMOS DBFS
sSQL = "DELETE FROM INGRESOS"
ExecuteSQL cCONNECTDBF, sSQL


'ABRIMOS COMM
Rs_DATOSDBF.LockType = adLockOptimistic
Rs_DATOSDBF.ActiveConnection = cCONNECTDBF
sSQL = "SELECT * FROM INGRESOS"
Rs_DATOSDBF.Open sSQL

'GRABAMOS COMM

If Not oRs Is Nothing Then

    Do While Not oRs.EOF
        Rs_DATOSDBF.AddNew
        For Each oCampo In oRs.Fields
            Rs_DATOSDBF.Fields(oCampo.Name).Value = Mid(RTrim(oCampo.Value), 1, 40)
        Next
        Rs_DATOSDBF.Update
        oRs.MoveNext
    Loop
    Rs_DATOSDBF.Close
    Set Rs_DATOSDBF = Nothing

    oRs.Close
    Set oRs = Nothing
    
    
End If
Set oRs = Nothing

If bMensaje Then
    Aviso "Proceso Culminó satisfactoriamente", 2
End If

Exit Sub
errorx:
    If err.Number = -2147467259 Then
        Resume Next
    End If
        
    errores err.Number
    
    If Not oRs Is Nothing Then
        oRs.Close
        Set oRs = Nothing
    End If
    
    If Not Rs_DATOSDBF Is Nothing Then
        Rs_DATOSDBF.Close
        Set Rs_DATOSDBF = Nothing
    End If


End Sub




Private Sub BuscarDAOT()

On Error GoTo drDepurar

Dim sSQL As String
Dim fmtCon As JSFmtCondition

Encuentra_Fechas

sSQL = "Ventas_Muestra_Segun_Requerimiento '" & dFecIni & "','" & dFecFin & "','R'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
GridEX1.Columns("Nro").Width = 390
GridEX1.Columns("Codigo").Width = 1335
GridEX1.Columns("Nombre").Width = 3750

GridEX1.Columns("cod_tipanex").Visible = False
GridEX1.Columns("cod_anxo").Visible = False
GridEX1.Columns("Origen").Visible = False

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Tipo").Visible = False

GridEX1.Columns("Importe_Soles").Width = 1185
GridEX1.Columns("Importe_Soles").Caption = "Valor Venta Soles"
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Width = 1365
GridEX1.Columns("Importe_Dolares").Caption = "Valor Venta Dolares"
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Cantidad").Width = 1020
GridEX1.Columns("Cantidad").Format = "###,###.00"

GridEX1.Columns("Porcentaje").Width = 900
GridEX1.Columns("Porcentaje").Format = "###,###.0000"

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0
  
Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub

Private Sub txtPais_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

