VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCambioFecFinProduccion 
   Caption         =   "Cambio fecha fin de produccion"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   750
      TabIndex        =   10
      Tag             =   "&Accept"
      Top             =   2415
      Width           =   1470
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   2700
      TabIndex        =   9
      Tag             =   "&Cancel"
      Top             =   2400
      Width           =   1485
   End
   Begin VB.Frame fraActualiza 
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   1530
      Width           =   4830
      Begin MSComCtl2.DTPicker dtpFecFinProd 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   270
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   40001
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fec Fin Producc :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Tag             =   "New PO :"
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txtPO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         MaxLength       =   20
         TabIndex        =   3
         Top             =   285
         Width           =   2940
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         MaxLength       =   20
         TabIndex        =   2
         Top             =   645
         Width           =   2940
      End
      Begin VB.TextBox txtEstilo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1005
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Tag             =   "PO :"
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Tag             =   "Client :"
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estilo :"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Tag             =   "Style :"
         Top             =   1035
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCambioFecFinProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cod_Cliente As String
Public Cod_PurOrd  As String
Public Cod_LotPurOrd As String
Public Cod_EstCli As String
Dim strSql As String
Sub CAMBIO()
On Error GoTo AceptaErr
    strSql = "EXEC SM_TG_LotEstUpdateFec_DespachoActual '" & Cod_Cliente & "','" & Cod_PurOrd & "','" & Cod_LotPurOrd & "','" & Cod_EstCli & "','" & dtpFecFinProd.value & "','" & vusu & "'"
    Call ExecuteCommandSQL(cCONNECT, strSql)
    MsgBox "El cambio fue exitoso", vbInformation, "PO Change"
    Unload Me
    Exit Sub
AceptaErr:
    ErrorHandler Err, "Error en Cambio"
End Sub

Private Sub cmdAceptar_Click()
    Call CAMBIO
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


