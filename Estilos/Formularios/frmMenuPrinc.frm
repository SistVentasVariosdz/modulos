VERSION 5.00
Begin VB.Form frmMenuPrinc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   870
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   3735
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   615
      Left            =   2160
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Items2"
      Height          =   615
      Left            =   2280
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   855
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   735
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Telas Enviadas a Desarrollo de Comercial"
      Height          =   705
      Left            =   2040
      TabIndex        =   14
      Top             =   120
      Width           =   1605
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Rapport"
      Height          =   450
      Left            =   270
      TabIndex        =   13
      Top             =   7995
      Width           =   1485
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Grupo Tallas"
      Height          =   495
      Left            =   300
      TabIndex        =   12
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "GrupoPro"
      Height          =   450
      Left            =   285
      TabIndex        =   11
      Top             =   6705
      Width           =   1485
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Procesos"
      Height          =   480
      Left            =   300
      TabIndex        =   10
      Top             =   5475
      Width           =   1485
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Estilo Cliente"
      Height          =   495
      Left            =   285
      TabIndex        =   9
      Top             =   4920
      Width           =   1485
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Tela"
      Height          =   495
      Left            =   285
      TabIndex        =   8
      Top             =   4230
      Width           =   1485
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Item"
      Height          =   495
      Left            =   285
      TabIndex        =   7
      Top             =   3600
      Width           =   1485
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   465
      Left            =   285
      TabIndex        =   6
      Top             =   6120
      Width           =   1485
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tipo Comp"
      Height          =   465
      Left            =   285
      TabIndex        =   5
      Top             =   3015
      Width           =   1485
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Mat. Prima"
      Height          =   465
      Left            =   285
      TabIndex        =   4
      Top             =   2445
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Comp. Estilo"
      Height          =   465
      Left            =   285
      TabIndex        =   3
      Top             =   1875
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pza Estilo"
      Height          =   465
      Left            =   285
      TabIndex        =   2
      Top             =   1305
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Talla"
      Height          =   465
      Left            =   285
      TabIndex        =   1
      Top             =   735
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hilados"
      Height          =   465
      Left            =   240
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
frmMantHilTel.Show
End Sub

Private Sub Command10_Click()
frmEstCliTem.Show
End Sub

Private Sub Command11_Click()
    frmProcesos.Show 1
End Sub

Private Sub Command12_Click()
    frmGrupoPro.Show
End Sub

Private Sub Command13_Click()
frmMantESTalla.Show
End Sub

Private Sub Command14_Click()
    Load frmShowTX_Rapport
    frmShowTX_Rapport.Show vbModal
    Set frmShowTX_Rapport = Nothing
End Sub

Private Sub Command15_Click()
frmMuestraTelasEnviadasDesarrollo.Show
End Sub

Private Sub Command16_Click()
frmMantItemServicios.Show
End Sub

Private Sub Command17_Click()
frmCambiaGrafico.Codigo_item = "AC000001"
frmCambiaGrafico.StrImagen1_Origen = "C:\Estilos\ImagenesDesarrollo\JE000032.jpg"
frmCambiaGrafico.Show
End Sub

Private Sub Command18_Click()
frmMantItemServicios.Show
End Sub

Private Sub Command19_Click()
frmMantFamItem.Show 1
End Sub

Private Sub Command2_Click()
frmMantESTalla.Show
End Sub
Private Sub Command3_Click()
frmMantPzaEst.Show
End Sub
Private Sub Command4_Click()
frmMantCompEst.Show
End Sub
Private Sub Command5_Click()
frmMantMatPri.Show
End Sub
Private Sub Command6_Click()
frmMantTipComp.Show
End Sub
Private Sub Command7_Click()
Unload Me
End Sub
Private Sub Command8_Click()
frmManItems.Show
End Sub

Private Sub Command9_Click()
frmManTelas.Show
End Sub

Private Sub Form_Load()

cCONNECT = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=textilesjoc;Data Source=192.168.1.10"
cSEGURIDAD = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=Seguridad;Data Source=192.168.1.10"

vusu = "SISTEMAS"
vper = "0001"
vemp = "01"
vemp1 = vemp
vRuta = App.Path
InitMessages
iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

End Sub
