VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasPacking 
   Caption         =   "Revisión Packings vs Facturas Exportación Prendas"
   ClientHeight    =   7005
   ClientLeft      =   2580
   ClientTop       =   1290
   ClientWidth     =   10515
   Icon            =   "frmVentasPacking.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   75
      TabIndex        =   1
      Top             =   120
      Width           =   10320
      Begin VB.CheckBox chkerror 
         Caption         =   "Ver Sólo Erradas"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   58982403
         CurrentDate     =   37887
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   525
         Left            =   6375
         TabIndex        =   4
         Top             =   150
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   926
         Custom          =   $"frmVentasPacking.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1250
         ControlHeigth   =   500
         ControlSeparator=   40
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año - Mes : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1050
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5820
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   10266
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
      Column(1)       =   "frmVentasPacking.frx":03FA
      Column(2)       =   "frmVentasPacking.frx":04C2
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmVentasPacking.frx":0566
      FormatStyle(2)  =   "frmVentasPacking.frx":069E
      FormatStyle(3)  =   "frmVentasPacking.frx":074E
      FormatStyle(4)  =   "frmVentasPacking.frx":0802
      FormatStyle(5)  =   "frmVentasPacking.frx":08DA
      FormatStyle(6)  =   "frmVentasPacking.frx":0992
      FormatStyle(7)  =   "frmVentasPacking.frx":0A72
      FormatStyle(8)  =   "frmVentasPacking.frx":0B1E
      ImageCount      =   0
      PrinterProperties=   "frmVentasPacking.frx":0BCE
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10875
      Top             =   5985
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmVentasPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String
Dim strOrigen As String
Dim dFecIni As Date, dFecFin As Date
Public serror1 As String

Private Sub Form_Load()
dtpAnoMes.Value = Date
End Sub

Private Sub Buscar()
Dim strSQL As String
On Error GoTo Fin

    strSQL = "Costos_Revisa_Packing_Despachados_Exportacion_Facturables '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & serror1 & "'"

    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    'GridEX1.Columns("Num_Corre_Planilla").Width = 0
    GridEX1.FrozenColumns = 4
    
    GridEX1.Columns("Nom_Cliente").Width = 2000
    GridEX1.Columns("Nom_Cliente").Caption = "NomCliente"
    GridEX1.Columns("Nom_Cliente").HeaderAlignment = jgexAlignCenter
      
    GridEX1.Columns("Num_Packing").Width = 600
    GridEX1.Columns("Num_Packing").Caption = "NºPacking"
    GridEX1.Columns("Num_Packing").HeaderAlignment = jgexAlignCenter
    
    GridEX1.Columns("Fec_EmidOc").Width = 1000
    GridEX1.Columns("Fec_EmidOc").Caption = "FecEmidOc"
    GridEX1.Columns("Fec_EmidOc").HeaderAlignment = jgexAlignCenter
   
    GridEX1.Columns("Fec_DESPACHO").Width = 1000
    GridEX1.Columns("Fec_DESPACHO").Caption = "FecDespacho"
    GridEX1.Columns("Fec_DESPACHO").HeaderAlignment = jgexAlignCenter
    
    GridEX1.Columns("Factura").Width = 1200
    GridEX1.Columns("Factura").Caption = "Factura"
    GridEX1.Columns("Factura").HeaderAlignment = jgexAlignCenter
   
    GridEX1.Columns("Moneda").Width = 700
    GridEX1.Columns("Moneda").Caption = "Moneda"
    GridEX1.Columns("Moneda").HeaderAlignment = jgexAlignCenter
    
    GridEX1.Columns("Prendas").Width = 700
    GridEX1.Columns("Prendas").Caption = "Prendas"
    GridEX1.Columns("Prendas").HeaderAlignment = jgexAlignCenter
    
    GridEX1.Columns("Clase_PO").Width = 800
    GridEX1.Columns("Clase_PO").Caption = "Clase PO"
    GridEX1.Columns("Clase_PO").HeaderAlignment = jgexAlignCenter
   
    GridEX1.Columns("cod_tipo_venta").Width = 1000
    GridEX1.Columns("cod_tipo_venta").Caption = "Cod Tip. Venta"
    GridEX1.Columns("cod_tipo_venta").HeaderAlignment = jgexAlignCenter
    
    GridEX1.Columns("Num_Corre_Venta").Width = 1000
    GridEX1.Columns("Num_Corre_Venta").Caption = "Num Corre Venta"
    GridEX1.Columns("Num_Corre_Venta").HeaderAlignment = jgexAlignCenter
    
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub


Private Sub Mostrar()
Dim strSQL As String
On Error GoTo Fin
    
    Set GridEX1.ADORecordset = Nothing

    strSQL = "CF_MUESTRA_MOVS_SALIDA_DESPACHO_CLIENTES_APT  '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "'"

    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    'GridEX1.Columns("Num_Corre_Planilla").Width = 0
   ' GridEX1.FrozenColumns = 4
    
'    GridEX1.Columns("Nom_Cliente").Width = 2000
'    GridEX1.Columns("Nom_Cliente").Caption = "NomCliente"
'    GridEX1.Columns("Nom_Cliente").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Num_Packing").Width = 600
'    GridEX1.Columns("Num_Packing").Caption = "NºPacking"
'    GridEX1.Columns("Num_Packing").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Fec_EmidOc").Width = 1000
'    GridEX1.Columns("Fec_EmidOc").Caption = "FecEmidOc"
'    GridEX1.Columns("Fec_EmidOc").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Fec_DESPACHO").Width = 1000
'    GridEX1.Columns("Fec_DESPACHO").Caption = "FecDespacho"
'    GridEX1.Columns("Fec_DESPACHO").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Factura").Width = 1200
'    GridEX1.Columns("Factura").Caption = "Factura"
'    GridEX1.Columns("Factura").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Moneda").Width = 700
'    GridEX1.Columns("Moneda").Caption = "Moneda"
'    GridEX1.Columns("Moneda").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Prendas").Width = 700
'    GridEX1.Columns("Prendas").Caption = "Prendas"
'    GridEX1.Columns("Prendas").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Clase_PO").Width = 800
'    GridEX1.Columns("Clase_PO").Caption = "Clase PO"
'    GridEX1.Columns("Clase_PO").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("cod_tipo_venta").Width = 1000
'    GridEX1.Columns("cod_tipo_venta").Caption = "Cod Tip. Venta"
'    GridEX1.Columns("cod_tipo_venta").HeaderAlignment = jgexAlignCenter
'
'    GridEX1.Columns("Num_Corre_Venta").Width = 1000
'    GridEX1.Columns("Num_Corre_Venta").Caption = "Num Corre Venta"
'    GridEX1.Columns("Num_Corre_Venta").HeaderAlignment = jgexAlignCenter
    If GridEX1.RowCount > 0 Then
    Call Reporte
    End If
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "VER_SOLO_ERRADOS"
    If chkerror.Value = 1 Then
        serror1 = "S"
    Else
        serror1 = "N"
    End If
      Buscar
      
    Case "MOVGENFACT"
     Mostrar
    Case "SALIR"
       Unload Me
    End Select
End Sub

Sub Reporte()
Dim strSQL As String
Dim oo As Object
Dim Ruta As String
Dim sEmpresa As String
On Error GoTo Errox
 
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
    
Ruta = ""
Ruta = vRuta & "\RPT_MUESTRA_MOVS_SALIDA_DESPACHO_CLIENTES_APT.XLT"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "reporte", GridEX1.ADORecordset, Format(dtpAnoMes.Value, "YYYY"), Right("00" & Format(dtpAnoMes.Value, "MM"), 2), sEmpresa
Set oo = Nothing
Exit Sub

Errox:
    ErrorHandler err, "Reporte"
End Sub


