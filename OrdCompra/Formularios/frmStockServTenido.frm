VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmStockServTenido 
   Caption         =   "Stocks en Servicio de Teñido"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerDetalle 
      Caption         =   "&Ver Detalle"
      Height          =   525
      Left            =   1695
      TabIndex        =   4
      Top             =   6435
      Width           =   1245
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   525
      Left            =   375
      TabIndex        =   3
      Top             =   6435
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   525
      Left            =   10740
      TabIndex        =   2
      Top             =   6435
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   12000
      Begin GridEX20.GridEX gexLista 
         Height          =   5895
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   10398
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmStockServTenido.frx":0000
         Column(2)       =   "frmStockServTenido.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmStockServTenido.frx":016C
         FormatStyle(2)  =   "frmStockServTenido.frx":02A4
         FormatStyle(3)  =   "frmStockServTenido.frx":0354
         FormatStyle(4)  =   "frmStockServTenido.frx":0408
         FormatStyle(5)  =   "frmStockServTenido.frx":04E0
         FormatStyle(6)  =   "frmStockServTenido.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmStockServTenido.frx":0678
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2940
      Top             =   6510
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmStockServTenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String

Public Sub CARGA_GRID()
   
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC SM_STOCKS_EN_SERVICIO_TENIDO"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    Call CONFIGURAR_GRID
    
    If gexLista.RowCount > 0 Then
        'HabilitaMant Me.FunctButt1, "IMPRIMIR"
        
    Else
        'HabilitaMant Me.FunctButt1, "GENERAR/REVERTIR/IMPRIMIR/SALIR"
    End If

End Sub

Public Sub REPORTE()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open App.Path & "\RptCfOrdProAvios.xlt"
    oo.Workbooks.Open vRuta & "\RptStockServTenido.xlt"
    oo.Visible = True
    
    oo.Run "REPORTE", cConnect
    
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de OrdPro de Avios " & Err.Description, vbCritical, "Impresion"
End Sub

Private Sub cmdImprimir_Click()
    Call Me.REPORTE
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVerDetalle_Click()
    Load frmDetalleServTenido
    frmDetalleServTenido.sSer_OrdCom = gexLista.Value(gexLista.Columns("SER_ORDCOMP").Index)
    frmDetalleServTenido.sCod_OrdCom = gexLista.Value(gexLista.Columns("COD_ORDCOMP").Index)
    frmDetalleServTenido.sCod_Tela = gexLista.Value(gexLista.Columns("COD_TELA").Index)
    frmDetalleServTenido.sDes_Tela = gexLista.Value(gexLista.Columns("DESCRIPCION").Index)
    frmDetalleServTenido.sCod_Combinacion = gexLista.Value(gexLista.Columns("COD_COMB").Index)
    frmDetalleServTenido.CARGA_GRID
    frmDetalleServTenido.Show vbModal
    Set frmDetalleServTenido = Nothing
End Sub

Private Sub Form_Load()
'LoadConnectEmpresa ""
'LoadConnectSeguridad ""
'vemp = "01"
'vper = "0001"
'vusu = "sistemas"
'InitMessages

    Call Me.CARGA_GRID
End Sub

Public Sub CONFIGURAR_GRID()

    Me.gexLista.Columns("SER_ORDCOMP").Visible = False
    Me.gexLista.Columns("COD_ORDCOMP").Visible = False
    Me.gexLista.Columns("COD_COMB").Visible = True
    Me.gexLista.Columns("COD_COMB").Width = 400
    Me.gexLista.Columns("NOMBRE_COMB").Visible = True
    Me.gexLista.Columns("NOMBRE_COMB").Width = 900
        

    Me.gexLista.Columns("GRUPO").Caption = "Grupo"
    Me.gexLista.Columns("GRUPO").Width = 900
    Me.gexLista.Columns("O/C").Caption = "O/C"
    Me.gexLista.Columns("O/C").Width = 1100
    Me.gexLista.Columns("PROVEEDOR").Caption = "Proveedor"
    Me.gexLista.Columns("PROVEEDOR").Width = 2000
    Me.gexLista.Columns("COD_TELA").Caption = "Cod. Tela"
    Me.gexLista.Columns("COD_TELA").Width = 900
    Me.gexLista.Columns("DESCRIPCION").Caption = "Tela"
    Me.gexLista.Columns("DESCRIPCION").Width = 2000
    Me.gexLista.Columns("ENVIADO").Caption = "Enviado"
    Me.gexLista.Columns("ENVIADO").Width = 1000
    Me.gexLista.Columns("INGRESADO").Caption = "Recibido"
    Me.gexLista.Columns("INGRESADO").Width = 1000
    Me.gexLista.Columns("SALDO").Caption = "Saldo"
    Me.gexLista.Columns("SALDO").Width = 1000
    Me.gexLista.Columns("ROLLOS_ENVIADOS").Caption = "Rollos Enviados"
    Me.gexLista.Columns("ROLLOS_ENVIADOS").Width = 1000
    Me.gexLista.Columns("ROLLOS_RECIBIDOS").Caption = "Rollos Recib."
    Me.gexLista.Columns("ROLLOS_RECIBIDOS").Width = 1000
    Me.gexLista.Columns("SALDO_ROLLOS").Caption = "Saldo Rollos"
    Me.gexLista.Columns("SALDO_ROLLOS").Width = 1000
    Me.gexLista.Columns("ORDENES").Caption = "Ordenes"
    Me.gexLista.Columns("ORDENES").Width = 2200

End Sub

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, ByVal iFixsCols As Integer, ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If
    
    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub
