VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransaccionesAddCuadre 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6180
   ClientLeft      =   1365
   ClientTop       =   1065
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9810
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   480
      TabIndex        =   1
      Top             =   5520
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesAddCuadre.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   9340
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmTransaccionesAddCuadre.frx":0265
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransaccionesAddCuadre.frx":05B7
      Column(2)       =   "frmTransaccionesAddCuadre.frx":067F
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransaccionesAddCuadre.frx":0723
      FormatStyle(2)  =   "frmTransaccionesAddCuadre.frx":085B
      FormatStyle(3)  =   "frmTransaccionesAddCuadre.frx":090B
      FormatStyle(4)  =   "frmTransaccionesAddCuadre.frx":09BF
      FormatStyle(5)  =   "frmTransaccionesAddCuadre.frx":0A97
      FormatStyle(6)  =   "frmTransaccionesAddCuadre.frx":0B4F
      FormatStyle(7)  =   "frmTransaccionesAddCuadre.frx":0C2F
      FormatStyle(8)  =   "frmTransaccionesAddCuadre.frx":10E7
      ImageCount      =   1
      ImagePicture(1) =   "frmTransaccionesAddCuadre.frx":1533
      PrinterProperties=   "frmTransaccionesAddCuadre.frx":1885
   End
End
Attribute VB_Name = "frmTransaccionesAddCuadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, lfAceptar As Boolean, strCod_Anexo As String, strCod_TipAnexo, _
       intNum_Transaccion As Long, dFecha As Date, strCod_Moneda As String

Sub CARGA_GRID()
Dim colTemp As JSColumn
Dim fmtCon As JSFmtCondition
Dim oGroup As GridEX20.JSGroup

On Error GoTo errores

Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Set oGroup = gridex1.Groups.Add(gridex1.Columns("Grupo").Index, jgexSortAscending)

gridex1.DefaultGroupMode = jgexDGMExpanded

gridex1.Columns("Cod").Width = 420
gridex1.Columns("Des_Concepto").Width = 3240
gridex1.Columns("Documento").Width = 1500
gridex1.Columns("TipoCambio").Width = 975
gridex1.Columns("Imp_Debe").Width = 945
gridex1.Columns("Imp_Debe").Format = "###,###.00"
gridex1.Columns("Imp_Haber").Width = 945
gridex1.Columns("Imp_Haber").Format = "###,###.00"
gridex1.Columns("Observaciones").Width = 1935
gridex1.Columns("Grupo").Visible = False
gridex1.Columns("Tipo").Visible = False

gridex1.GroupFooterStyle = jgexTotalsGroupFooter

Set colTemp = gridex1.Columns("TipoCambio")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = "TOTAL "

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Imp_Debe")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Imp_Haber")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

Set fmtCon = gridex1.FmtConditions.Add(gridex1.Columns("tipo").Index, jgexEqual, "1")
fmtCon.FormatStyle.BackColor = &HFFFF00

Set fmtCon = gridex1.FmtConditions.Add(gridex1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &H8080FF

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim lvSql As String

On Error GoTo DrpDepurar

Select Case ActionName
Case Is = "ACEPTAR"
  lfAceptar = True
  Unload Me
Case Is = "AGREGAADELANDO"
  Carga_Mantenimieno "Ventas_Muestra_Anticipos_Pedientes ", "TM_Ventas_MAN_Transacciones_Cobranzas_Adelantos ", True, False, " ( Adicion Adelanto )"
'Case Is = "AGREGACANJE"
'  Carga_Mantenimieno "Ventas_Muestra_Docum_Pedientes_Pagos ", "TM_Ventas_MAN_Transacciones_Cobranzas_Canjes ", False, True, " ( Adicion Canje )"
Case Is = "AGREGANOTA"
  Carga_Mantenimieno "Ventas_Muestra_Docum_Pedientes_Cobranzas_Notas ", "TM_Ventas_MAN_Transacciones_Cobranzas_Notas_Abono ", False, True, " ( Adicion Nota Abono )"
Case Is = "AGREGAFINAN"
  Carga_Mantenimieno_Finan "", "", False, False, " ( Adicion Financiamiento )"
Case Is = "VERDETALLE"
  If gridex1.RowCount = 0 Then Exit Sub
  Carga_Detalle gridex1.Value(gridex1.Columns("Cod").Index)
Case Is = "CANCELAR"
  lfAceptar = False
  Unload Me
End Select

Exit Sub

DrpDepurar:

errores err.Number

End Sub

Sub Carga_Mantenimieno(Store_Carga As String, Store As String, dAnticipo As Boolean, dDoc As Boolean, strTitulo As String)
  With frmTransaccionesAddCuadreMan
    .strCod_TipAnex = strCod_TipAnexo
    .strCod_Anexo = strCod_Anexo
    .strCod_Moneda = strCod_Moneda
    .intNum_Transaccion = intNum_Transaccion
    .dFecha = dFecha
    .strStore_Carga = Store_Carga
    .strStore = Store
    .frAnticipo.Visible = dAnticipo
    .frDocumento.Visible = dDoc
    .strOption = "I"
    .Caption = Me.Caption & strTitulo
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With
End Sub
Sub Carga_Mantenimieno_Finan(Store_Carga As String, Store As String, dAnticipo As Boolean, dDoc As Boolean, strTitulo As String)
  With frmTransaccionesAddCuadreManFinan
    .strCod_TipAnexo = strCod_TipAnexo
    .strCod_Anexo = strCod_Anexo
    .strCod_Moneda = strCod_Moneda
    .intNum_Transaccion = intNum_Transaccion
    .dFecha = dFecha
    .strStore_Carga = Store_Carga
    .strStore = Store
    .strOption = "I"
    .Caption = Me.Caption & strTitulo
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With
End Sub

Sub Carga_Detalle(strCod_Concepto As String)

Dim strTipo_Cod As String, strSotreMan As String, strCodx As String, strCaption As String

Select Case strCod_Concepto
Case Is = gcAnticipos
  strTipo_Cod = "A"
  strSotreMan = "TM_Ventas_MAN_Transacciones_Cobranzas_Adelantos "
  strCodx = gcAnticipos
  strCaption = " Anticipo"
Case Is = gcNota_Credito_Clientes
  strTipo_Cod = "N"
  strCodx = gcNota_Credito_Clientes
  strSotreMan = "TM_Ventas_MAN_Transacciones_Cobranzas_Notas_Abono "
  strCaption = " Notas Abono"
Case Is = gcCanjes
  strTipo_Cod = "C"
  strSotreMan = "TM_Ventas_MAN_Transacciones_Cobranzas_Canjes "
  strCodx = gcCanjes
  strCaption = " Canjes "
End Select

If strTipo_Cod <> "" Then
  With frmTransaccionesAddCuadreDet
    .strTipo_Det = strCodx
    .strSQL = "Ventas_Muestra_Temporales_Cobranzas " & intNum_Transaccion & ",'" & strTipo_Cod & "'"
    .intNum_Transaccion = intNum_Transaccion
    .strCod_Anexo = strCod_Anexo
    .strCod_TipAnexo = strCod_TipAnexo
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .dTipo_Cambio = gridex1.Value(gridex1.Columns("TipoCambio").Index)
    .CARGA_GRID
    .StrSql_Man = strSotreMan
    .Caption = Me.Caption & strCaption
    .Show 1
    CARGA_GRID
  End With
Else
  MsgBox "No hay Detalle para este Concepto", vbInformation, "AVISO"
End If

End Sub
