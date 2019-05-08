VERSION 5.00
Begin VB.Form frmMenuPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Consulta por Rango"
      Height          =   660
      Left            =   2400
      TabIndex        =   21
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Reporte Partidas Programadas"
      Height          =   645
      Left            =   2520
      TabIndex        =   20
      Top             =   5160
      Width           =   1485
   End
   Begin VB.CommandButton Command24 
      Caption         =   "KARDEX"
      Height          =   600
      Left            =   2520
      TabIndex        =   19
      Top             =   4440
      Width           =   1515
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Asignar Accesos CF_Almacen"
      Height          =   600
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Width           =   1635
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Asignar Accesos LG_Almacen"
      Height          =   600
      Left            =   2520
      TabIndex        =   17
      Top             =   3000
      Width           =   1635
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Reporte Ingresos X Comprobante"
      Height          =   525
      Left            =   2595
      TabIndex        =   16
      Top             =   870
      Width           =   1500
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Stocks Telas Valorizados Mensuales"
      Height          =   720
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
      Width           =   1605
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Ver Saldos de Stock"
      Height          =   615
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Width           =   1620
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Control Produccion Mensual"
      Height          =   735
      Left            =   2565
      TabIndex        =   13
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Transportista"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "STOCKS POR ALMACEN"
      Height          =   510
      Left            =   510
      TabIndex        =   11
      Top             =   5610
      Width           =   1800
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Req. Vs Real Avios"
      Height          =   510
      Left            =   495
      TabIndex        =   10
      Top             =   5055
      Width           =   1800
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Stock Familias"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4515
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Tipos de Movimiento"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   3975
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Almacen"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3450
      Width           =   1800
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Datos Tecnicos"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2895
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   555
      Left            =   2400
      TabIndex        =   5
      Top             =   6840
      Width           =   1785
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Movimientos Stocks"
      Height          =   525
      Left            =   480
      TabIndex        =   4
      Top             =   2325
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kardex Tela Acabada"
      Height          =   500
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1785
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Movimiento Almacen"
      Height          =   500
      Left            =   510
      TabIndex        =   2
      Top             =   1275
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kardex Hilo Crudo"
      Height          =   500
      Left            =   510
      TabIndex        =   1
      Top             =   750
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kardex Tela Cruda"
      Height          =   500
      Left            =   510
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    FrmKardexTelCru.Show 1
End Sub

Private Sub Command10_Click()
    FrmMantAlmacen.Show 1
End Sub

Private Sub Command11_Click()
    FrmMantTipMov.Show 1
End Sub

Private Sub Command12_Click()
    FrmStockFam.Show vbModal
    
End Sub

'Private Sub Command13_Click()
    'frmMovStocksGuias.Show 1
'End Sub

Private Sub Command14_Click()
    frmReqVsRealAvios.Show 1
End Sub

Private Sub Command15_Click()
    FrmRep.Show vbModal
    
End Sub

Private Sub Command16_Click()
    frmManTransportistas.Show 1
End Sub



Private Sub Command18_Click()
    frmControlProdMensual.Show 1
End Sub

'Private Sub Command19_Click()
'    frmStocksSaldos.Show vbModal
'End Sub

Private Sub Command2_Click()
    FrmKardexHilCru.Show 1
End Sub

Private Sub Command20_Click()
    frmRptStkTelas.Show 1
End Sub

Private Sub Command21_Click()
    frmRptIngxComprobante.Show 1
End Sub

Private Sub Command22_Click()
    frmAccLG_SEGALM.Show vbModal
End Sub

Private Sub Command23_Click()
    frmAccCF_SEGALM.Show vbModal
End Sub

Private Sub Command24_Click()
    Load FrmKardex
    FrmKardex.Show vbModal
    Set FrmKardex = Nothing
End Sub

Private Sub Command25_Click()
FrmPartidasProgramadas.Show 1
End Sub

Private Sub Command26_Click()

End Sub

'Private Sub Command26_Click()
'    frmReclamosAviosProd.Show vbModal
'    Set frmReclamosAviosProd = Nothing
'End Sub

Private Sub Command27_Click()
    frmMovStocks.Show 1
End Sub

Private Sub Command3_Click()
    FrmMovAlmacen.Show 1
End Sub

Private Sub Command4_Click()
    FrmKardexTelaca.Show 1
End Sub

Private Sub Command5_Click()
    frmMovStocks.Show 1
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
    frmDatosTecnicos.Show 1
End Sub

Private Sub Command8_Click()
Frm_StockCriticos.Show 1
End Sub

Private Sub Form_Load()
'cConnect = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=Textilesjoc;Data Source=192.168.1.10"

cConnect = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=facontex;Data Source=192.168.1.10"
cSEGURIDAD = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=Seguridad;Data Source=192.168.1.10"
    vper = "0001"
    vemp = "01"
    vemp1 = "01"
    vusu = "sistemas"
    vRuta = App.Path
End Sub
