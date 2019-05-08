VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdelantos 
   Caption         =   "Registro de Adelantos Clientes"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   1035
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   10275
   Begin VB.TextBox txtDes_TipAne 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   5265
   End
   Begin VB.TextBox txtDes_TipAnex 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "C"
      Top             =   480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtNum_Ruc 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7395
      MaxLength       =   11
      TabIndex        =   3
      Top             =   480
      Width           =   1545
   End
   Begin VB.TextBox txtCod_TipAne 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      MaxLength       =   4
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "C"
      Top             =   480
      Width           =   360
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   9120
      TabIndex        =   4
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox txtDes_Status 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   105
      Width           =   1575
   End
   Begin VB.TextBox txtCod_Status 
      Height          =   285
      Left            =   870
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "P"
      Top             =   105
      Width           =   375
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   9419
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
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmAdelantos.frx":0000
      Column(2)       =   "frmAdelantos.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmAdelantos.frx":016C
      FormatStyle(2)  =   "frmAdelantos.frx":02A4
      FormatStyle(3)  =   "frmAdelantos.frx":0354
      FormatStyle(4)  =   "frmAdelantos.frx":0408
      FormatStyle(5)  =   "frmAdelantos.frx":04E0
      FormatStyle(6)  =   "frmAdelantos.frx":0598
      FormatStyle(7)  =   "frmAdelantos.frx":0678
      FormatStyle(8)  =   "frmAdelantos.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmAdelantos.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1530
      TabIndex        =   5
      Top             =   6360
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   900
      Custom          =   $"frmAdelantos.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1680
      Top             =   6360
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   525
      Width           =   570
   End
   Begin VB.Label Label28 
      Caption         =   "R.U.C."
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   495
      Width           =   495
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Status :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmAdelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, strCod_Anxo As String

Private Sub cmdBuscar_Click()
  Buscar
End Sub
Sub Buscar()
Dim strSQL
On Error GoTo errores

strSQL = "Ventas_Man_Adelantos 'V','" & txtCod_TipAne & "','" & strCod_Anxo & "',0,'" & txtCod_Status & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Dim colTemp As JSColumn

GridEX1.Columns("Cliente").Width = 2820
GridEX1.Columns("Ruc").Width = 1230
GridEX1.Columns("Nro_Anticipo").Width = 1050
GridEX1.Columns("Fecha").Width = 945
GridEX1.Columns("Cod_Moneda").Width = 825
GridEX1.Columns("Cod_Moneda").Caption = "Moneda"
GridEX1.Columns("Imp_Anticipo").Width = 1050
GridEX1.Columns("Imp_Cancelado").Width = 1245
GridEX1.Columns("descripcion").Width = 975
GridEX1.Columns("descripcion").Caption = "Status"
GridEX1.Columns("Cod_Tipanex").Visible = False
GridEX1.Columns("Cod_Anxo").Visible = False
GridEX1.Columns("Moneda").Visible = False
GridEX1.Columns("Flg_Status").Visible = False

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Private Sub Form_Load()

Dim oFrm As New Frm_Toolbar
oFrm.CambiarContenedor Me
Set oFrm = Nothing

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo hand

Select Case ActionName
  Case Is = "AGREGAR"
    If strCod_Anxo = "" Then
      MsgBox "Tiene que seleccionar un Cliente", vbInformation, "AVISO"
      Exit Sub
    End If
    With frmAdelantosAdd
      .strOption = "I"
      .txtNro_Anticipo = DevuelveCampo("select Ult_Nro_Anticipo + 1 from  vt_control", cCONNECT)
      .txtNro_Anticipo.Enabled = False
      .strCod_TipAnex = txtCod_TipAne
      .strCod_Anexo = strCod_Anxo
      .txtFecha.Text = Date
      .intNum_Anticipo = 0
      .txtDes_TipAne = txtDes_TipAne
      .txtDes_TipAne.Enabled = False
      .Caption = "Adicion Anticipo al Cliente " & txtDes_TipAne
      .Show 1
      If .lfAceptar Then Buscar
    End With
    
  Case Is = "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    With frmAdelantosAdd
      .Caption = "Actualiza un Anticipo del Cliente " & txtDes_TipAne
      .strCod_TipAnex = GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index)
      .strCod_Anexo = GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index)
      .txtNro_Anticipo.Text = GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index)
      .txtNro_Anticipo.Enabled = False
      .txtFecha.Text = GridEX1.Value(GridEX1.Columns("Fecha").Index)
      .txtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index)
      .txtDes_Moneda.Text = GridEX1.Value(GridEX1.Columns("Moneda").Index)
      .Txt_Importe.Text = GridEX1.Value(GridEX1.Columns("Imp_Anticipo").Index)
      .TxtObservacion.Text = GridEX1.Value(GridEX1.Columns("Observacion").Index)
      .txtCod_TipCobra = GridEX1.Value(GridEX1.Columns("Cod_TipCobranza").Index)
      .txtDes_TipCobra = GridEX1.Value(GridEX1.Columns("des_TipCobranza").Index)
      .txtSec_Parte = GridEX1.Value(GridEX1.Columns("Secuencia_Parte").Index)
      .strOption = "U"
      .txtAnticipo_Moneda_Deposito.Text = GridEX1.Value(GridEX1.Columns("Imp_Anticipo_Moneda_Deposito").Index)
      .txtTipoCambioNegociado.Text = GridEX1.Value(GridEX1.Columns("Tipo_Cambio_Negociado").Index)
      If Val(.txtAnticipo_Moneda_Deposito.Text) <> 0 Then
        .fraAnticipoOtraMoneda.Visible = True
      Else
        .fraAnticipoOtraMoneda.Visible = False
      End If
      .Show 1
      If .lfAceptar Then Buscar
      Buscar
    End With
  Case Is = "VERCANCELACIONES"
      If GridEX1.RowCount = 0 Then Exit Sub
      Load frmAdelantosCancelaciones
      With frmAdelantosCancelaciones
        .Caption = "CANCELACION DE ADELANTOS " & Trim(GridEX1.Value(GridEX1.Columns("Cliente").Index)) & " NRO " & GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index)
        .strSQL = "Cn_Adelantos_Detalle_Cancelaciones '" & GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "'," & GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index)
        .CARGA_GRID
        .Show 1
      End With
  Case Is = "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    If MsgBox("Esta seguro de Eliminar este Adelanto", vbYesNo, "IMPORTANTE") = vbYes Then
      lvSql = "Ventas_Man_Adelantos 'D','" & GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "'," & GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index)
      Call ExecuteCommandSQL(cCONNECT, lvSql)
      Buscar
    End If
  Case Is = "IMPRIMIR"
    Imprimir
  Case Is = "SALIR"
    Unload Me
End Select

Exit Sub
Resume
hand:

errores err.Number

End Sub

Private Sub TxtCod_Status_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("flg_status", "Descripcion", "Cn_Ventas_Adelantos_Status where ", txtCod_Status, txtDes_Status, 1, Me)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub TxtDes_Status_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("flg_status", "Descripcion", "Cn_Ventas_Adelantos_Status where ", txtCod_Status, txtDes_Status, 2, Me)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
End Sub

Private Sub Imprimir()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim strSQL As String
Dim sEmpresa As String
    strSQL = "SELECT DES_COMP_EMP FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)


    Ruta = vRuta & "\AnticipoClientes.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "reporte", GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index), GridEX1.Value(GridEX1.Columns("CLIENTE").Index), GridEX1.Value(GridEX1.Columns("IMP_ANTICIPO").Index), GridEX1.Value(GridEX1.Columns("COD_MONEDA").Index), GridEX1.Value(GridEX1.Columns("FECHA").Index), sEmpresa
    
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub
