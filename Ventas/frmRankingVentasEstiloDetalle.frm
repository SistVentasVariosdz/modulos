VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRankingVentasEstiloDetalle 
   Caption         =   "Detalle"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   4800
      TabIndex        =   1
      Top             =   4860
      Width           =   1155
   End
   Begin GridEX20.GridEX grxListado 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   8467
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
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      GridLines       =   2
      BackColorBkg    =   15531775
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRankingVentasEstiloDetalle.frx":0000
      Column(2)       =   "frmRankingVentasEstiloDetalle.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRankingVentasEstiloDetalle.frx":016C
      FormatStyle(2)  =   "frmRankingVentasEstiloDetalle.frx":02A4
      FormatStyle(3)  =   "frmRankingVentasEstiloDetalle.frx":0354
      FormatStyle(4)  =   "frmRankingVentasEstiloDetalle.frx":0408
      FormatStyle(5)  =   "frmRankingVentasEstiloDetalle.frx":04E0
      FormatStyle(6)  =   "frmRankingVentasEstiloDetalle.frx":0598
      FormatStyle(7)  =   "frmRankingVentasEstiloDetalle.frx":0678
      FormatStyle(8)  =   "frmRankingVentasEstiloDetalle.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmRankingVentasEstiloDetalle.frx":07D4
   End
End
Attribute VB_Name = "frmRankingVentasEstiloDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public f1 As Date
Public f2 As Date
Public Cod_EstCli As String
Public Cod_OrdPro As String
Dim strSQL As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION_ESTILO '" & f1 & "', '" & f2 & "', '8', '', '', '', '','" & Cod_EstCli & "','" & Cod_OrdPro & "'"
    Screen.MousePointer = vbHourglass
    Set grxListado.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Screen.MousePointer = vbDefault
End Sub
