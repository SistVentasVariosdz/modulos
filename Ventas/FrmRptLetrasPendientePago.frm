VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptLetrasPendientePago 
   Caption         =   "Letras Pendiente de Pago"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   14220
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   5520
      TabIndex        =   3
      Top             =   7560
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"FrmRptLetrasPendientePago.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14175
      Begin VB.OptionButton Opt_Vendedor 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox Txt_DesUsuario 
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox Txt_Cod_Usuario 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optFechaVen 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Vencimiento"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   270
         Width           =   1695
      End
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   12840
         TabIndex        =   2
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
      Begin VB.Frame frCliente 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   150
         Width           =   7455
         Begin VB.TextBox txtCod_TipAne 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4800
            MaxLength       =   4
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "C"
            Top             =   120
            Width           =   360
         End
         Begin VB.TextBox txtNum_Ruc 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5835
            MaxLength       =   11
            TabIndex        =   1
            Top             =   120
            Width           =   1545
         End
         Begin VB.TextBox txtDes_TipAne 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   600
            MaxLength       =   30
            TabIndex        =   0
            Top             =   120
            Width           =   4065
         End
         Begin VB.TextBox txtDes_TipAnex 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4080
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "C"
            Top             =   120
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0FFFF&
            Caption         =   "R.U.C."
            Height          =   255
            Left            =   5280
            TabIndex        =   12
            Top             =   135
            Width           =   495
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cliente :"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   165
            Width           =   570
         End
      End
      Begin VB.Frame frMensual 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   7455
         Begin MSComCtl2.DTPicker DTAnoMes 
            Height          =   330
            Left            =   960
            TabIndex        =   14
            Top             =   0
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "MM / yyyy"
            Format          =   89653251
            CurrentDate     =   37987
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año/Mes :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   75
            Width           =   750
         End
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6540
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   11536
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
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
      Column(1)       =   "FrmRptLetrasPendientePago.frx":0090
      Column(2)       =   "FrmRptLetrasPendientePago.frx":0158
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptLetrasPendientePago.frx":01FC
      FormatStyle(2)  =   "FrmRptLetrasPendientePago.frx":0334
      FormatStyle(3)  =   "FrmRptLetrasPendientePago.frx":03E4
      FormatStyle(4)  =   "FrmRptLetrasPendientePago.frx":0498
      FormatStyle(5)  =   "FrmRptLetrasPendientePago.frx":0570
      FormatStyle(6)  =   "FrmRptLetrasPendientePago.frx":0628
      FormatStyle(7)  =   "FrmRptLetrasPendientePago.frx":0708
      FormatStyle(8)  =   "FrmRptLetrasPendientePago.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "FrmRptLetrasPendientePago.frx":0864
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptLetrasPendientePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCod_Anxo As String
Public codigo As String, Descripcion As String
Dim strSQL As String

Private Sub Form_Load()
  DTAnoMes = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()

Dim oGroup As GridEX20.JSGroup
Dim dFecIni As String, dFecFin As String

On Error GoTo errCarga

If optFechaVen Then
  dFecIni = "'" & Trim(CStr(CDate("01/" & Format(Month(DTAnoMes), "00") & "/" & Year(DTAnoMes)))) & "'"
  dFecFin = "'" & DevuelveCampo("Select dbo.tg_obtiene_dia_ultimo_ano_mes('" & Format(Year(DTAnoMes), "0000") & "','" & Format(Month(DTAnoMes), "00") & "')", cCONNECT) & "'"
End If

If optCliente Or Opt_Vendedor Then
  dFecIni = "NULL"
  dFecFin = "NULL"
End If


If optCliente Then
    strSQL = "Cn_Ventas_Muestra_Letras_Pendientes_Res '" & txtCod_TipAne & "','" & strCod_Anxo & "'," & dFecIni & "," & dFecFin & ",'','','1'"
End If
If optFechaVen Then
    strSQL = "Cn_Ventas_Muestra_Letras_Pendientes_Res '" & txtCod_TipAne & "','" & strCod_Anxo & "'," & dFecIni & "," & dFecFin & ",'','','2'"
End If
If Opt_Vendedor Then
    strSQL = "Cn_Ventas_Muestra_Letras_Pendientes_Res '" & txtCod_TipAne & "','" & strCod_Anxo & "'," & dFecIni & "," & dFecFin & ",'" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "','3'"
End If

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.ColumnHeaderHeight = 500

If optCliente Then

  Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cli").Index, jgexSortAscending)

  GridEX1.Columns("Cli").Visible = False
  GridEX1.Columns("Fecha_Emision").Width = 1065
  GridEX1.Columns("Fecha_Emision").Caption = "Fecha Emision"
  GridEX1.Columns("Cliente").Visible = False
  GridEX1.Columns("Ruc").Visible = False
Else
  GridEX1.Columns("Cliente").Width = 3045
  GridEX1.Columns("Ruc").Width = 1320
End If


GridEX1.Columns("Letra").Width = 795
GridEX1.Columns("Fecha_Vencimiento").Width = 1065
GridEX1.Columns("Moneda").Width = 720
GridEX1.Columns("Fecha_Vencimiento").Caption = "Fecha Vencimiento"
GridEX1.Columns("Saldo_Soles").Width = 1035
GridEX1.Columns("Saldo_Soles").Caption = "Saldo  Soles"
GridEX1.Columns("Saldo_Dolares").Width = 1035
GridEX1.Columns("Saldo_Dolares").Caption = "Saldo Dolares"
GridEX1.Columns("Status_Letra").Width = 1245
GridEX1.Columns("Status_Letra").Caption = "Status Letra"
GridEX1.Columns("Banco").Width = 2505
GridEX1.Columns("Letra_Banco").Width = 1380

GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.BackColorRowGroup = &H80000005

MuestraSubTotales



Exit Sub
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Private Sub MuestraSubTotales()

Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Fecha_Vencimiento")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Saldo_Soles")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Saldo_Dolares")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub


Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim strSQL As String
Dim sEmpresa As String

    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If GridEX1.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")


If optCliente Or Opt_Vendedor Then
  oo.Workbooks.Open vRuta & "\RptLetrasPendientesPAgo.XLT"
  oo.Visible = True
  oo.DisplayAlerts = False
  oo.Run "reporte", GridEX1.ADORecordset, sEmpresa
Else
  oo.Workbooks.Open vRuta & "\RptLetrasPendientesPAgoFecVend.XLT"
  oo.Visible = True
  oo.DisplayAlerts = False
  oo.Run "reporte", GridEX1.ADORecordset, UCase(Format(DTAnoMes, "MMMM")), sEmpresa
End If

Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub optCliente_Click()
frCliente.Visible = True
GridEX1.ClearFields
End Sub

Private Sub optFechaVen_Click()
frCliente.Visible = False
GridEX1.ClearFields
End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
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




Public Sub Busca_Trabajador()
On Error GoTo Fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim opcion As String
Dim strSQL As String
strSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!codigo)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!codigo), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!codigo), 4)
            'stip_Trabajador = Left(rstAux!codigo, 1)
            'scod_trabajador = Right(rstAux!codigo, 4)
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
    "Búsqueda de Color (" & opcion & ")"
End Sub



