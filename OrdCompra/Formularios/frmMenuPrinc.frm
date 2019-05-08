VERSION 5.00
Begin VB.Form frmMenuPrinc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Stocks Serv Tenido"
      Height          =   465
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1485
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   465
      Left            =   255
      TabIndex        =   1
      Top             =   1290
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Orden de Compra"
      Height          =   465
      Left            =   285
      TabIndex        =   0
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "frmMenuPrinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'frmManItems.Show 1
frmOrdComp.Show 1
End Sub
Private Sub Command2_Click()
'frmManTelas.Show 1
frmClaOrdComp.Show 1
End Sub
Private Sub Command3_Click()
    e.Show 1
End Sub
Private Sub Command6_Click()
   FrmDscto.Show
End Sub
Private Sub Command7_Click()
Unload Me
End Sub
Private Sub Command9_Click()
    Load frmStockServTenido
    frmStockServTenido.Show vbModal
    Set frmStockServTenido = Nothing
End Sub

Private Sub Form_Load()
LoadConnectEmpresa ""
LoadConnectSeguridad ""
vemp = "01"
vper = "0001"
vusu = "gtelas"
'vusu = "evillar"
vRuta = App.Path
InitMessages
End Sub
