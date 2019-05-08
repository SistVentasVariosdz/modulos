VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form FrmShowRepLog 
   Caption         =   "OPCIONES REPORTES LOGISTICOS"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdResumenEncajado 
      Caption         =   "Resumen Encajado"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Control de Importaciones"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compras por Proveedor"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Status de Avios"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1485
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1485
   End
   Begin VB.CommandButton CmdReporteOCitem 
      Caption         =   "Orden de compra por Item"
      Height          =   540
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1485
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   255
      Top             =   3120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmShowRepLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReporteOCitem_Click()
    FrmReporteOCompraItems.Show vbModal
End Sub

Private Sub CmdResumenEncajado_Click()
FrmResumenEncajado.Show vbModal
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
FrmStatusAvios.Show vbModal
End Sub

Private Sub Command2_Click()
FrmComprasxProveedor.Show vbModal
End Sub

Private Sub Command3_Click()
FrmControlImportaciones.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub
