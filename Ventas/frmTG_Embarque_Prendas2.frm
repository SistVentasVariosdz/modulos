VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTG_Embarque_Prendas 
   Caption         =   "Detalle Embarque Prendas"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNP 
      Caption         =   "NP"
      Height          =   945
      Left            =   60
      TabIndex        =   26
      Top             =   15
      Width           =   6825
      Begin VB.TextBox txtCod_OrdPro 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   795
         MaxLength       =   5
         TabIndex        =   0
         Top             =   495
         Width           =   915
      End
      Begin VB.TextBox txtDes_OrdPro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1785
         TabIndex        =   29
         Top             =   495
         Width           =   4815
      End
      Begin VB.TextBox txtNom_Fabrica 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1335
         TabIndex        =   28
         Top             =   165
         Width           =   2490
      End
      Begin VB.TextBox txtCod_Fabrica 
         Height          =   285
         Left            =   795
         TabIndex        =   27
         Top             =   165
         Width           =   480
      End
      Begin VB.Label lblOrdPro 
         AutoSize        =   -1  'True
         Caption         =   "N/P:"
         Height          =   195
         Left            =   165
         TabIndex        =   31
         Top             =   540
         Width           =   345
      End
      Begin VB.Label lblFabrica 
         Caption         =   "Fábrica"
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame fraProgramado 
      Caption         =   "Detalle Programado"
      Height          =   2970
      Left            =   60
      TabIndex        =   19
      Top             =   990
      Width           =   3030
      Begin VB.TextBox txtCubicaje_Prog 
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Tag             =   "SET"
         Top             =   2505
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto_Prog 
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Tag             =   "SET"
         Top             =   2055
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto_Prog 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Tag             =   "SET"
         Top             =   1605
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Cajas_Prog 
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Tag             =   "SET"
         Top             =   1155
         Width           =   1200
      End
      Begin VB.TextBox txtPre_Unitario 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Tag             =   "SET"
         Top             =   705
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Prendas_Prog 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Tag             =   "SET"
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label lblCubicaje_Prog 
         Caption         =   "Cubicaje Prog"
         Height          =   405
         Left            =   135
         TabIndex        =   25
         Tag             =   "CUBICAJE_PROG"
         Top             =   2505
         Width           =   1500
      End
      Begin VB.Label lblPeso_Neto_Prog 
         Caption         =   "Peso Neto Prog"
         Height          =   480
         Left            =   135
         TabIndex        =   24
         Tag             =   "PESO_NETO_PROG"
         Top             =   2055
         Width           =   1500
      End
      Begin VB.Label lblPeso_Bruto_Prog 
         Caption         =   "Peso Bruto"
         Height          =   480
         Left            =   135
         TabIndex        =   23
         Tag             =   "PESO_BRUTO_PROG"
         Top             =   1605
         Width           =   1500
      End
      Begin VB.Label lblNum_Cajas_Prog 
         Caption         =   "Num Cajas "
         Height          =   480
         Left            =   135
         TabIndex        =   22
         Tag             =   "NUM_CAJAS_PROG"
         Top             =   1155
         Width           =   1500
      End
      Begin VB.Label lblPre_Unitario 
         Caption         =   "Precio Unitario"
         Height          =   480
         Left            =   135
         TabIndex        =   21
         Tag             =   "PRE_UNITARIO"
         Top             =   705
         Width           =   1500
      End
      Begin VB.Label lblNum_Prendas_Prog 
         Caption         =   "Prendas "
         Height          =   480
         Left            =   135
         TabIndex        =   20
         Tag             =   "NUM_PRENDAS_PROG"
         Top             =   255
         Width           =   1500
      End
   End
   Begin VB.Frame fraReal 
      Caption         =   "Detalle Real"
      Enabled         =   0   'False
      Height          =   2970
      Left            =   3765
      TabIndex        =   8
      Top             =   990
      Width           =   3105
      Begin VB.TextBox txtCubicaje 
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Tag             =   "SET"
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto 
         Height          =   300
         Left            =   1680
         TabIndex        =   16
         Tag             =   "SET"
         Top             =   2070
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto 
         Height          =   300
         Left            =   1680
         TabIndex        =   14
         Tag             =   "SET"
         Top             =   1620
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Cajas 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Tag             =   "SET"
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Prendas 
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Tag             =   "SET"
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblCubicaje 
         Caption         =   "Cubicaje"
         Height          =   390
         Left            =   135
         TabIndex        =   17
         Tag             =   "CUBICAJE"
         Top             =   2535
         Width           =   1500
      End
      Begin VB.Label lblPeso_Neto 
         Caption         =   "Peso Neto"
         Height          =   480
         Left            =   135
         TabIndex        =   15
         Tag             =   "PESO_NETO"
         Top             =   2070
         Width           =   1500
      End
      Begin VB.Label lblPeso_Bruto 
         Caption         =   "Peso Bruto"
         Height          =   480
         Left            =   135
         TabIndex        =   13
         Tag             =   "PESO_BRUTO"
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label lblNum_Cajas 
         Caption         =   "Num. Cajas"
         Height          =   480
         Left            =   135
         TabIndex        =   11
         Tag             =   "NUM_CAJAS"
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label lblNum_Prendas 
         Caption         =   "Num_Prendas"
         Height          =   480
         Left            =   135
         TabIndex        =   9
         Tag             =   "NUM_PRENDAS"
         Top             =   240
         Width           =   1500
      End
   End
   Begin FunctionsButtons.FunctButt FunctOKCancel 
      Height          =   510
      Left            =   2265
      TabIndex        =   7
      Top             =   4200
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~8~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Line Line1 
      X1              =   5385
      X2              =   6375
      Y1              =   4530
      Y2              =   4530
   End
End
Attribute VB_Name = "frmTG_Embarque_Prendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sTituliAbrOP  As String

Private Sub Form_Load()
    Dim sSQL As String
    VerificaFabrica txtCod_Fabrica, txtNom_Fabrica
    sTituliAbrOP = DevuelveCampo("select Titulo_Abr_Orden from TG_Control", cCONNECT)
    lblOrdPro.Caption = sTituliAbrOP
    
End Sub
Private Sub VerificaFabrica(ByRef objFabrica As TextBox, ByRef objNombreFabrica As TextBox)
On Error GoTo errorx
    Dim sSQL As String
    Dim iRet As String
    
    sSQL = "SELECT count(*) FROM TG_Fabrica "
    iRet = DevuelveCampo(sSQL, cCONNECT)
    If iRet = 1 Then
        sSQL = "SELECT Cod_Fabrica FROM TG_Fabrica "
        objFabrica.Text = DevuelveCampo(sSQL, cCONNECT)
        
        sSQL = "SELECT Nom_Fabrica FROM TG_Fabrica "
        objNombreFabrica.Text = DevuelveCampo(sSQL, cCONNECT)
        objFabrica.Enabled = False
        objNombreFabrica.Enabled = False
        
    End If
Exit Sub
errorx:
    errores Err.Number
    
End Sub

