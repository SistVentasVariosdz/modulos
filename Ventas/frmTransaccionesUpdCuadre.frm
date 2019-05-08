VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransaccionesUpdCuadre 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6630
   ClientLeft      =   555
   ClientTop       =   900
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10095
   Begin GridEX20.GridEX GridEX1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   11456
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmTransaccionesUpdCuadre.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransaccionesUpdCuadre.frx":0352
      Column(2)       =   "frmTransaccionesUpdCuadre.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransaccionesUpdCuadre.frx":04BE
      FormatStyle(2)  =   "frmTransaccionesUpdCuadre.frx":05F6
      FormatStyle(3)  =   "frmTransaccionesUpdCuadre.frx":06A6
      FormatStyle(4)  =   "frmTransaccionesUpdCuadre.frx":075A
      FormatStyle(5)  =   "frmTransaccionesUpdCuadre.frx":0832
      FormatStyle(6)  =   "frmTransaccionesUpdCuadre.frx":08EA
      FormatStyle(7)  =   "frmTransaccionesUpdCuadre.frx":09CA
      FormatStyle(8)  =   "frmTransaccionesUpdCuadre.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmTransaccionesUpdCuadre.frx":12CE
      PrinterProperties=   "frmTransaccionesUpdCuadre.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   6510
      Left            =   8880
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   11483
      Custom          =   $"frmTransaccionesUpdCuadre.frx":17F8
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesUpdCuadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, dFecha As Date, intSecuencia As Integer, strCod_Moneda As String, strCod_Anexo As String, strCod_TipAnexo


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
gridex1.Columns("Sec_Detalle").Visible = False
gridex1.Columns("Num_Corre").Visible = True
gridex1.Columns("Debe_Haber").Visible = False
gridex1.Columns("Fec_Transaccion").Visible = False
gridex1.Columns("Secuencia").Visible = False
gridex1.Columns("Ser_Docum").Visible = False
gridex1.Columns("Num_Docum_Ventas").Visible = False
gridex1.Columns("Cod_Moneda").Visible = False
gridex1.Columns("Moneda_Doc").Visible = False
gridex1.Columns("Importe").Visible = False
gridex1.Columns("Des_TipDoc").Visible = False
gridex1.Columns("Cod_TipDoc").Visible = False
gridex1.Columns("TipoCambio_Ori").Visible = False

gridex1.GroupFooterStyle = jgexTotalsGroupFooter

Set colTemp = gridex1.Columns("TipoCambio")
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

Dim lvSql As String, strTipo_Cod As String

On Error GoTo DrpDepurar

Select Case ActionName

Case Is = "ADICIONARDOCUM"
  Carga_Mantenimieno " Ventas_Muestra_Docum_Pedientes_Cobranzas ", " CN_Ventas_Trans_Cobranz_DETALLE_PRINCIPAL ", False, True, False, " ( Adiciona Documento ) ", "1"
Case Is = "ADICIONARCONCEPTO"
  Carga_Mantenimieno "", " CN_Ventas_Trans_Cobranz_DETALLE_PRINCIPAL ", False, False, True, " ( Adiciona Concepto ) ", "4"
Case Is = "AGREGAADELANTO"
  Carga_Mantenimieno "Ventas_Muestra_Anticipos_Pedientes ", "CN_Ventas_MAN_Trans_Cobranzas_Adelant_PRINCIPAL ", True, False, False, " ( Adicion Adelanto )", "2"
Case Is = "AGREGACANJE"
  Carga_Mantenimieno "Ventas_Muestra_Docum_Pedientes_Pagos ", "CN_Ventas_MAN_Trans_Cobranzas_Canjes_PRINCIPAL ", False, True, False, " ( Adicion Canje )", "3"
Case Is = "AGREGANOTAABONO"
  Carga_Mantenimieno "Ventas_Muestra_Docum_Pedientes_Cobranzas_Notas ", "CN_Ventas_MAN_Trans_Cobranzas_Notas_Abono_PRINCIPAL ", False, True, False, " ( Adicion Nota Abono )", "3"
Case Is = "AGREGAFINAN"
  Carga_Mantenimieno_Fina "", "", False, False, False, " ( Adiciona Financiamiento )", ""
Case Is = "MODIFICAR"
  If gridex1.RowCount = 0 Then Exit Sub
  strTipo_Cod = ""
  Select Case gridex1.Value(gridex1.Columns("Cod").Index)
  Case Is = gcAnticipos
    strTipo_Cod = "A"
  Case Is = gcNota_Credito_Clientes
    strTipo_Cod = "N"
  Case Is = gcCanjes
    strTipo_Cod = "C"
  End Select
  If strTipo_Cod <> "" Then
    MsgBox "No se puede Modificar este concepto entre a Ver Detalle ", vbInformation, "AVISO"
    Exit Sub
  End If
  
  If Trim(gridex1.Value(gridex1.Columns("Nro_Financiamiento").Index)) <> "" Then
    MsgBox "No se puede Modificar este concepto ", vbInformation, "AVISO"
    Exit Sub
  End If
  
  With frmTransaccionesUpdCuadreMan
    .Caption = Me.Caption
    .strOption = "U"
    .intSecuencia_Det = gridex1.Value(gridex1.Columns("Sec_Detalle").Index)
    .strStore = " CN_Ventas_Trans_Cobranz_DETALLE_PRINCIPAL "
    .strNum_Corre = gridex1.Value(gridex1.Columns("Num_Corre").Index)
    .txtCod_Cobranza.Text = gridex1.Value(gridex1.Columns("Cod").Index)
    .txtDes_Cobranza.Text = gridex1.Value(gridex1.Columns("Des_Concepto").Index)
    .dFecha = gridex1.Value(gridex1.Columns("Fec_Transaccion").Index)
    .intSecuencia = gridex1.Value(gridex1.Columns("Secuencia").Index)
    .TxtObservacion = gridex1.Value(gridex1.Columns("Observaciones").Index)
    .TxtTipo_Cambio.Text = gridex1.Value(gridex1.Columns("TipoCambio").Index)
    .txtCod_Moneda.Text = gridex1.Value(gridex1.Columns("Moneda_Doc").Index)
    .strCod_Moneda = strCod_Moneda
    .txtOtro_Tipo_Cambio.Text = gridex1.Value(gridex1.Columns("TipoCambio_Otros").Index)
    
    If Trim(gridex1.Value(gridex1.Columns("Num_Corre").Index)) = "" Then
      .frDocumento.Visible = False
      .txtCod_Cobranza.Enabled = False
      .txtDes_Cobranza.Enabled = False
      .txtImporte.Text = gridex1.Value(gridex1.Columns("Importe").Index)
      .strTipo_Det = "1"
    Else
      .frDocumento.Visible = True
      .txtCod_TipDoc.Enabled = False
      .txtDes_TipDoc.Enabled = False
      .txtSer_Docum.Enabled = False
      .txtNum_Docum.Enabled = False
      .txtImp_Convertido.Enabled = False
      .txtCod_TipDoc = gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
      .txtDes_TipDoc = gridex1.Value(gridex1.Columns("Des_TipDoc").Index)
      .txtSer_Docum = gridex1.Value(gridex1.Columns("Ser_Docum").Index)
      .txtNum_Docum = gridex1.Value(gridex1.Columns("Num_Docum_Ventas").Index)
      .strTipo_Det = "1"
'        If GridEX1.Value(GridEX1.Columns("Moneda_Doc").Index) = strCod_Moneda Then
          .txtImporte.Text = gridex1.Value(gridex1.Columns("Importe").Index)
'        Else
'          If GridEX1.Value(GridEX1.Columns("Moneda_Doc").Index) = "SOL" Then
'            .txtImporte.Text = GridEX1.Value(GridEX1.Columns("Importe").Index) / GridEX1.Value(GridEX1.Columns("TipoCambio").Index)
'          Else
'            .txtImporte.Text = GridEX1.Value(GridEX1.Columns("Importe").Index) * GridEX1.Value(GridEX1.Columns("TipoCambio").Index)
'          End If
'        End If
      .frConcepto.Visible = False
      .Calcula_Importe_Converido
    End If
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With
Case Is = "VERDETALLE"
  If gridex1.RowCount = 0 Then Exit Sub
  Carga_Detalle gridex1.Value(gridex1.Columns("Cod").Index)
Case Is = "ELIMINAR"
  If gridex1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta seguro de Eliminar este Concepto", vbYesNo, "IMPORTANTE") = vbYes Then
    lvSql = "CN_Ventas_Trans_Cobranz_DETALLE_PRINCIPAL  'D','" & gridex1.Value(gridex1.Columns("Fec_Transaccion").Index) & "'," _
            & gridex1.Value(gridex1.Columns("Secuencia").Index) & "," & gridex1.Value(gridex1.Columns("Sec_Detalle").Index) & ",'" _
            & gridex1.Value(gridex1.Columns("Cod").Index) & "','','" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "'," _
            & gridex1.Value(gridex1.Columns("Importe").Index) & ",'" & gridex1.Value(gridex1.Columns("Observaciones").Index) & "'," _
            & gridex1.Value(gridex1.Columns("TipoCambio").Index)
    ExecuteCommandSQL cCONNECT, lvSql
    CARGA_GRID
  End If

Case Is = "TIPOCAMBIO"

  With frmTransaccionesUpdCuadreManTipoCambio
    .Caption = Me.Caption
    .strOption = "U"
    .strStore = " CN_Ventas_Trans_Cobranz_DETALLE_PRINCIPAL_Tipo_Cambio "
    .dFecha = gridex1.Value(gridex1.Columns("Fec_Transaccion").Index)
    .intSecuencia = gridex1.Value(gridex1.Columns("Secuencia").Index)
    .intSecuencia_Det = gridex1.Value(gridex1.Columns("Sec_Detalle").Index)
    .TxtTipo_Cambio.Text = gridex1.Value(gridex1.Columns("TipoCambio_Ori").Index)
    .TxtTipo_Cambio_Otro.Text = gridex1.Value(gridex1.Columns("TipoCambio_Otros").Index)
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With

Case Is = "SALIR"

  Unload Me
End Select

Exit Sub
Resume

DrpDepurar:

errores err.Number

End Sub

Sub Carga_Mantenimieno(Store_Carga As String, Store As String, dAnticipo As Boolean, dDoc As Boolean, dConcepto As Boolean, strTitulo As String, Cod_Det As String)
  With frmTransaccionesUpdCuadreMan
    .Caption = Trim(Me.Caption) & strTitulo
    
    .strCod_TipAnexo = strCod_TipAnexo
    .strCod_Anexo = strCod_Anexo
    .strCod_Moneda = strCod_Moneda
    .intSecuencia = intSecuencia
    .dFecha = dFecha
    
    .strStore_Carga = Store_Carga
    .strStore = Store
    
    .strTipo_Det = Cod_Det
    .frAnticipo.Visible = dAnticipo
    .frDocumento.Visible = dDoc
    .frConcepto.Visible = dConcepto
    
    .strOption = "I"
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With
End Sub

Sub Carga_Mantenimieno_Fina(Store_Carga As String, Store As String, dAnticipo As Boolean, dDoc As Boolean, dConcepto As Boolean, strTitulo As String, Cod_Det As String)
  With frmTransaccionesUpdCuadreManFinan
    .Caption = Trim(Me.Caption) & strTitulo
    
    .strCod_TipAnexo = strCod_TipAnexo
    .strCod_Anexo = strCod_Anexo
    .strCod_Moneda = strCod_Moneda
    .intSecuencia = intSecuencia
    .dFecha = dFecha
    
    .strStore_Carga = Store_Carga
    .strStore = Store
    
    .strTipo_Det = Cod_Det
    
    
    .strOption = "I"
    .Show 1
    If .lfAceptar Then CARGA_GRID
  End With
End Sub

Sub Carga_Detalle(strCod_Concepto As String)

Dim strTipo_Cod As String, strSotreMan As String, strCodx As String, strCaption As String, Cod_Det As String

Select Case strCod_Concepto
Case Is = gcAnticipos
  strTipo_Cod = "A"
  strSotreMan = "Cn_Ventas_MAN_Trans_Cobranzas_Adelant_PRINCIPAL "
  strCodx = gcAnticipos
  strCaption = " Anticipo"
  Cod_Det = "2"
Case Is = gcNota_Credito_Clientes
  strTipo_Cod = "N"
  strCodx = gcNota_Credito_Clientes
  strSotreMan = "CN_Ventas_MAN_Trans_Cobranzas_Notas_Abono_PRINCIPAL "
  strCaption = " Notas Abono"
  Cod_Det = "3"
Case Is = gcCanjes
  strTipo_Cod = "C"
  strSotreMan = "CN_Ventas_MAN_Trans_Cobranzas_Canjes_PRINCIPAL "
  strCodx = gcCanjes
  strCaption = " Canjes "
  Cod_Det = "3"
End Select

If strTipo_Cod <> "" Then
  With frmTransaccionesUpdCuadreDet
    .strTipo_Det = strCodx
    .strSQL = "Ventas_Muestra_Fijos_Cobranzas '" & dFecha & "'," & intSecuencia & ",'" & strTipo_Cod & "'"
    .strCod_Anexo = strCod_Anexo
    .strCod_TipAnexo = strCod_TipAnexo
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .intSecuencia = intSecuencia
    .dTipo_Cambio = gridex1.Value(gridex1.Columns("TipoCambio").Index)
    .CARGA_GRID
    .StrSql_Man = strSotreMan
    .Caption = Me.Caption & strCaption
    .strCod_Det = Cod_Det
    .Show 1
    CARGA_GRID
  End With
Else
  MsgBox "No hay Detalle para este Concepto", vbInformation, "AVISO"
End If

End Sub

