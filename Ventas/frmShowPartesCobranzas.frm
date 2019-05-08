VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowPartesCobranzas 
   Caption         =   "Partes de Cobranzas"
   ClientHeight    =   7230
   ClientLeft      =   1635
   ClientTop       =   1725
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   6780
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   435
      Left            =   5220
      TabIndex        =   4
      Top             =   195
      Width           =   1185
   End
   Begin VB.TextBox txtCod_Origen 
      Height          =   285
      Left            =   990
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "N"
      Top             =   90
      Width           =   375
   End
   Begin VB.TextBox txtDes_Origen 
      Height          =   285
      Left            =   1740
      TabIndex        =   1
      Top             =   90
      Width           =   1575
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5430
      Left            =   90
      TabIndex        =   6
      Top             =   900
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   9578
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
      Column(1)       =   "frmShowPartesCobranzas.frx":0000
      Column(2)       =   "frmShowPartesCobranzas.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowPartesCobranzas.frx":016C
      FormatStyle(2)  =   "frmShowPartesCobranzas.frx":02A4
      FormatStyle(3)  =   "frmShowPartesCobranzas.frx":0354
      FormatStyle(4)  =   "frmShowPartesCobranzas.frx":0408
      FormatStyle(5)  =   "frmShowPartesCobranzas.frx":04E0
      FormatStyle(6)  =   "frmShowPartesCobranzas.frx":0598
      FormatStyle(7)  =   "frmShowPartesCobranzas.frx":0678
      FormatStyle(8)  =   "frmShowPartesCobranzas.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmShowPartesCobranzas.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   270
      TabIndex        =   5
      Top             =   6495
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1191
      Custom          =   $"frmShowPartesCobranzas.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   650
      ControlSeparator=   110
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   480
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94109697
      CurrentDate     =   37543
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   3270
      TabIndex        =   3
      Top             =   480
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
      Left            =   240
      TabIndex        =   9
      Top             =   510
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta :"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   510
      Width           =   615
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6135
      Top             =   6510
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Origen :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   135
      Width           =   555
   End
End
Attribute VB_Name = "frmShowPartesCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String

Private Sub cmdBuscar_Click()
  Buscar
End Sub
Sub Buscar()
Dim strSQL
On Error GoTo errores

strSQL = "CN_VENTAS_MUESTRA_PARTES_COBRANZA '" & txtCod_Origen & "','" & dtpDesde & "','" & dtpHasta & "'"
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Exit Sub
Resume
errores:
    errores err.Number
End Sub


Private Sub Form_Load()
  dtpDesde = Date
  dtpHasta = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "CERRAR"
        If gridex1.RowCount = 0 Then Exit Sub
            With frmTransaccionesStatus
              .txtCod_Origen = txtCod_Origen
              .txtDes_Origen = txtDes_Origen
              .txtFecha_Cierre.Text = DevuelveCampo("select isnull(max(Fec_Transaccion),getdate()) from CN_VENTAS_PARTES_COBRANZA where Origen = '" & txtCod_Origen & "'", cCONNECT)
              .txtFecha_Nuevo.Text = Date
              .Show 1
            End With
        Case "ABRIR"
          If gridex1.RowCount = 0 Then Exit Sub
          With frmTransaccionesStatusReversion
            .txtCod_Origen = txtCod_Origen
            .txtDes_Origen = txtDes_Origen
            .txtNum_Parte = gridex1.Value(gridex1.Columns("NUM_PARTE_COBRANZA").Index)
            .Show 1
            Buscar
          End With
        Case "GENERARPARTEADIC"
            GenerarParteAdicional
        Case "IMPRIMIRPENDIENTES"
            Reporte
        Case "SALIR"
            Unload Me
    End Select

End Sub
Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
    Me.cmdBuscar.SetFocus
  End If
End Sub


Private Sub GenerarParteAdicional()
On Error GoTo errx
Dim sSQL As String
Dim vResp As Variant

vResp = MsgBox("Desea Generar un Parte Adicional para la Fecha indicada: " & gridex1.Value(gridex1.Columns("FEC_TRANSACCION").Index), vbOKCancel, "Confirmación")

If vResp <> vbOK Then
    Exit Sub
End If

sSQL = "CN_VENTAS_PARTES_COBRANZA_GENERACION_ADICIONAL '$','$'"
sSQL = VBsprintf(sSQL, txtCod_Origen, gridex1.Value(gridex1.Columns("NUM_PARTE_COBRANZA").Index))

ExecuteCommandSQL cCONNECT, sSQL

Buscar

Exit Sub
errx:
    errores err.Number
    
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

VB.Screen.MousePointer = vbHourglass

Dim oo As Object, strSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim RS1 As Object
Set RS1 = CreateObject("ADODB.Recordset")
Set oo = CreateObject("excel.application")

strSQL = "CN_Muestra_Partes_Cobranzas_por_Status '" & txtCod_Origen & "','C'"

Set RS = CargarRecordSetDesconectado(strSQL, cCONNECT)

If RS.RecordCount = 0 Then
  Screen.MousePointer = vbNormal
  MsgBox "No hay Registros que imprimir", vbInformation, "AVISO"
  Exit Sub
End If

oo.Workbooks.Open vRuta & "\rptListadoPartesCobranza.xlt"
oo.Run "REPORTE", RS, txtCod_Origen, txtDes_Origen, cCONNECT

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


