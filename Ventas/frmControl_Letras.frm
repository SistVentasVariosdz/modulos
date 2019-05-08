VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControl_Letras 
   Caption         =   "Registro de Letras"
   ClientHeight    =   5730
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   90
      TabIndex        =   2
      Top             =   -15
      Width           =   11925
      Begin VB.OptionButton optBanco 
         Caption         =   "Banco"
         Height          =   225
         Left            =   4215
         TabIndex        =   10
         Top             =   405
         Width           =   960
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   465
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox TxtDes_Banco 
         Height          =   285
         Left            =   5895
         TabIndex        =   8
         Top             =   405
         Width           =   3855
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   285
         Left            =   5160
         TabIndex        =   7
         Top             =   405
         Width           =   615
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1860
         TabIndex        =   6
         Top             =   855
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   1410
         MaxLength       =   1
         TabIndex        =   5
         Top             =   855
         Width           =   375
      End
      Begin VB.OptionButton optstatus 
         Caption         =   "Status"
         Height          =   225
         Left            =   225
         TabIndex        =   4
         Top             =   915
         Width           =   1155
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   405
         Left            =   10755
         TabIndex        =   3
         Top             =   375
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   300
         Left            =   1410
         TabIndex        =   11
         Top             =   450
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   38590
      End
      Begin MSComCtl2.DTPicker DTPFecha1 
         Height          =   300
         Left            =   1515
         TabIndex        =   12
         Top             =   1305
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   38590
      End
      Begin MSComCtl2.DTPicker DTPFecha2 
         Height          =   300
         Left            =   4035
         TabIndex        =   13
         Top             =   1305
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   38590
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha de Inicio"
         Height          =   405
         Left            =   555
         TabIndex        =   15
         Top             =   1305
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Fecha de Fin"
         Height          =   405
         Left            =   3090
         TabIndex        =   14
         Top             =   1305
         Width           =   795
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3105
      Left            =   90
      TabIndex        =   0
      Top             =   1860
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   5477
      Version         =   "2.0"
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
      Column(1)       =   "frmControl_Letras.frx":0000
      Column(2)       =   "frmControl_Letras.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmControl_Letras.frx":016C
      FormatStyle(2)  =   "frmControl_Letras.frx":02A4
      FormatStyle(3)  =   "frmControl_Letras.frx":0354
      FormatStyle(4)  =   "frmControl_Letras.frx":0408
      FormatStyle(5)  =   "frmControl_Letras.frx":04E0
      FormatStyle(6)  =   "frmControl_Letras.frx":0598
      FormatStyle(7)  =   "frmControl_Letras.frx":0678
      FormatStyle(8)  =   "frmControl_Letras.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmControl_Letras.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   90
      TabIndex        =   1
      Top             =   5025
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1191
      Custom          =   $"frmControl_Letras.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1075
      ControlHeigth   =   650
      ControlSeparator=   75
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10635
      Top             =   -15
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmControl_Letras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String
Public sTipoBusq As String

Private Sub cmdBuscar_Click()
  buscar
End Sub
Sub buscar()

Dim strSQL
On Error GoTo errores

Dim sOpcion As Integer

If optFecha.Value = True Then
    sOpcion = 1
ElseIf optBanco.Value = True Then
    sOpcion = 2
Else
    sOpcion = 3
End If



strSQL = "Ventas_Muestra_Planilla_Letras '" & sOpcion & "','" & DTPFecha & "','" & TxtCod_Banco.Text & "','" & txtCod_Origen.Text & "','" & DTPFecha1 & "','" & DTPFecha2 & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Dim colTemp As JSColumn

GridEX1.ColumnHeaderHeight = 600

GridEX1.Columns("Num_Planilla_Letra").Width = 800
GridEX1.Columns("Cod_Banco").Width = 800
GridEX1.Columns("Fec_Presentacion").Width = 1000
GridEX1.Columns("Sec_Cuenta_Banco").Width = 800
GridEX1.Columns("Cod_TipDoc").Width = 1000
GridEX1.Columns("Flg_Status_Letras").Width = 1000
GridEX1.Columns("Nom_Funcionario").Width = 2000
GridEX1.Columns("BANCO").Width = 1500
GridEX1.Columns("CUENTA").Width = 1600
GridEX1.Columns("LETRA").Width = 1000


GridEX1.Columns("Num_Planilla_Letra").Caption = "Nº Planilla     Letra"
GridEX1.Columns("Cod_Banco").Caption = "Cod. Banco"
GridEX1.Columns("Fec_Presentacion").Caption = "Fecha"
GridEX1.Columns("Sec_Cuenta_Banco").Caption = "Cuenta"
GridEX1.Columns("Cod_TipDoc").Caption = "Tipo Doc."
GridEX1.Columns("Flg_Status_Letras").Caption = "Flg. Status Letras"
GridEX1.Columns("Nom_Funcionario").Caption = "Funcionario"
GridEX1.Columns("BANCO").Caption = "Banco"
GridEX1.Columns("CUENTA").Caption = "Nº Cuenta"
GridEX1.Columns("LETRA").Caption = "Nº Letra"


Exit Sub
Resume
errores:
    errores err.Number
End Sub

Public Sub ReporteContinental()

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\RptLetrasContinental1.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "yyyy"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "mm"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "dd"), vemp, GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index), cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"

End Sub

Public Sub ReporteCredito()

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\RptLetrasCredito.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "yyyy"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "mm"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "dd"), vemp, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"

End Sub

Public Sub ReporteHSBC()

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\RptLetrasHSBC.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "yyyy"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "mm"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "dd"), vemp, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"

End Sub

Public Sub ReporteBIF()

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\RptLetrasBIF.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "yyyy"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "mm"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "dd"), vemp, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"

End Sub

Public Sub ReporteScotiabank()

On Error GoTo ErrorImpresion
Dim oo As Object

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\RptLetrasScotiabank.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "yyyy"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "mm"), Format(GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index), "dd"), vemp, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
cmdBuscar.SetFocus
End Sub

Private Sub Form_Load()

DTPFecha.Value = Date
DTPFecha1.Value = Date
DTPFecha2.Value = Date

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'
Dim varSecuencia As Integer

On Error GoTo hand

Select Case ActionName
  Case "AGREGAR"
    With frmAddLetras
      .StrOption = "I"
      .sTipoBusq = 1
      .DTPFecha = Date
      .Show 1
      If .lfSalvar Then
        buscar
      Else
        FunctButt1.SetFocus
      End If
    End With

  Case "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load frmAddLetras
'      .txtCod_TipCobra.Enabled = False
      frmAddLetras.sTipoBusq = 2
      frmAddLetras.sNum_Planilla_Letra = GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index)
      frmAddLetras.TxtCod_Banco.Text = GridEX1.Value(GridEX1.Columns("Cod_Banco").Index)
      frmAddLetras.TxtDes_Banco.Text = GridEX1.Value(GridEX1.Columns("BANCO").Index)
      frmAddLetras.txtCuenta_Cod.Text = GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index)
      frmAddLetras.txtCuenta_Des.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)
      frmAddLetras.txtCod_Origen.Text = GridEX1.Value(GridEX1.Columns("Flg_Status_Letras").Index)
      frmAddLetras.txtDes_Origen.Text = GridEX1.Value(GridEX1.Columns("LETRA").Index)
      frmAddLetras.DTPFecha = GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index)
      frmAddLetras.TxtObservacion.Text = GridEX1.Value(GridEX1.Columns("Nom_Funcionario").Index)
      frmAddLetras.Show vbModal
      Set frmAddLetras = Nothing
      Call buscar

  Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub

'    If DevuelveCampo("select count(*) from Cn_Ventas_Transacciones_Cobranzas_Detalle where Fec_Transaccion = '" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "' and secuencia = " & GridEX1.Value(GridEX1.Columns("Secuencia").Index), cCONNECT) > 0 Then
'      MsgBox "Elimine Primero el Detalle de la Transaccion", vbInformation, "IMPORTATEN"
'      Exit Sub
'    End If

    If MsgBox("Esta seguro de eliminar la Letra a la Planilla Actual", vbYesNo, "IMPORTANTE") = vbYes Then
      lvSql = "Ventas_Generar_Control_Letras '" & 3 & "','" & GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) & "','" _
          & GridEX1.Value(GridEX1.Columns("Fec_Presentacion").Index) & "','" & GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index) & "', '81', '" & GridEX1.Value(GridEX1.Columns("Flg_Status_Letras").Index) & "','" & GridEX1.Value(GridEX1.Columns("Nom_Funcionario").Index) & "'"


      Call ExecuteCommandSQL(cCONNECT, lvSql)
      Call buscar
    End If

  Case "DETALLE"
        If GridEX1.RowCount > 0 Then
            frmDetalleLetras.sCuenta = GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index)
            frmDetalleLetras.sCOD_BANCO = GridEX1.Value(GridEX1.Columns("Cod_Banco").Index)
            frmDetalleLetras.sFlg_Status_Letras = GridEX1.Value(GridEX1.Columns("Flg_Status_Letras").Index)
            frmDetalleLetras.sNum_Planilla_Letra = GridEX1.Value(GridEX1.Columns("Num_Planilla_Letra").Index)
            frmDetalleLetras.Show vbModal
            Load frmDetalleLetras
            Set frmDetalleLetras = Nothing

        End If

  Case "IMPRIMIR"
      If GridEX1.RowCount = 0 Then Exit Sub
      If GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) = 3 Then
      ReporteCredito
      End If
      If GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) = 2 Then
      ReporteScotiabank
      End If
      If GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) = 11 Then
      ReporteBIF
      End If
      If GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) = 19 Then
      ReporteHSBC
      End If
      If GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) = 4 Then
      ReporteContinental
      End If
  Case "SALIR"
    Unload Me
End Select

Exit Sub





Resume
hand:

errores err.Number

End Sub

Sub Trans_Finanzas()

On Error GoTo dprDepurar

Dim ssql As String, iSecuencia As Integer

If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Transferir esta transaccion ha Finanzas", vbYesNo + vbCritical, "AVISO") = vbYes Then
  iSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
  ssql = "Fi_Genera_Movimiento_Finanza '" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & ",'S'"
  ExecuteCommandSQL cCONNECT, ssql
  buscar
  Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, iSecuencia)
  MsgBox "La Transferencia se hizo satisfactoriamente", vbInformation, "AVISO"
End If

Exit Sub

dprDepurar:

  Error err.Number

End Sub

Private Sub inpFec_Emi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Encuentra_Parte
  End If
End Sub

Private Sub inpFec_Emi_LostFocus()
  Encuentra_Parte
End Sub

Private Sub optFecha_Click()
    sTipoBusq = "1"
End Sub

Private Sub optParteCobranza_Click()
    sTipoBusq = "2"
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
  cmdBuscar.SetFocus
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
    Encuentra_Parte
  End If
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If KeyAscii = 13 Then Call Busca_Opcion("flg_status_letra", "descripcion", "cn_status_letras where flg_planilla_letra='s' and ", txtCod_Origen, txtDes_Origen, 1, Me)
        cmdBuscar.SetFocus
  End If
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub Encuentra_Parte()
  txtNum_Parte.Text = DevuelveCampo("Select isnull(MAX(Num_Parte_Cobranza),'') from Cn_Ventas_Partes_Cobranza where Fec_Transaccion = '" & inpFec_Emi.Text & "' and Origen = '" & txtCod_Origen & "'  ", cCONNECT)
End Sub

Private Sub MuestraVoucher2()

On Error GoTo errx
Dim ssql As String
Dim rsAsientos As ADODB.Recordset


If GridEX1.RowCount = 0 Then Exit Sub

ssql = "FI_Muestra_Data_Asientos_Cobranzas '$' ,'$'"
ssql = VBsprintf(ssql, GridEX1.Value(GridEX1.Columns("fecha").Index), GridEX1.Value(GridEX1.Columns("secuencia").Index))

Set rsAsientos = GetDataSet(cCONNECT, ssql)

With rsAsientos

  If .BOF Or .EOF Then
    MsgBox "No se le ha Generado Voucher", vbInformation, "AVISO"
    Exit Sub
  End If

  Load frmShowVoucher
  frmShowVoucher.sCod_TipoDiario = !Cod_TipoDiario
  frmShowVoucher.sano = !Ano_Contable
  frmShowVoucher.smes = !Mes_Contable
  frmShowVoucher.lNum_Registro = !Num_Registro
  frmShowVoucher.sFec_Transaccion = GridEX1.Value(GridEX1.Columns("fecha").Index)
  frmShowVoucher.sSecuencia = GridEX1.Value(GridEX1.Columns("secuencia").Index)
  'frmShowVoucher.Num_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
  'frmShowVoucher.dImporte = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
  'frmShowVoucher.sFlg_Status = GridEX1.Value(GridEX1.Columns("Estatus_Letra").Index)
  frmShowVoucher.buscar
  frmShowVoucher.FunctButt1.ChangeProperty "ENABLED", 1, False
  frmShowVoucher.Show vbModal
  Set frmShowVoucher = Nothing

End With

Set rsAsientos = Nothing

Exit Sub

errx:
    errores err.Number

End Sub



