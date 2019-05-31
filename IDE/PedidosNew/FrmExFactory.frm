VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExFactory 
   Caption         =   "Ex Factory Reprogramada"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraActualiza 
      Height          =   1080
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4470
      Begin MSComCtl2.DTPicker dtpFec 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58327041
         CurrentDate     =   40001
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Despacho Original Reprogramada :"
         Height          =   435
         Left            =   480
         TabIndex        =   4
         Tag             =   "New PO :"
         Top             =   300
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   2700
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   1230
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   750
      TabIndex        =   0
      Tag             =   "&Accept"
      Top             =   1245
      Width           =   1470
   End
End
Attribute VB_Name = "FrmExFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public COD_CLIENTE   As String

Public cod_purord    As String

Public cod_lotpurord As String

Public cod_estcli    As String

Dim strSql           As String

Sub CAMBIO()

    On Error GoTo AceptaErr

    strSql = "EXEC TG_LOTEST_ACTUALIZA_Fec_DespachoOri_Reprogramada '" & COD_CLIENTE & "','" & cod_purord & "','" & cod_lotpurord & "','" & cod_estcli & "','" & dtpFec.value & "','" & vusu & "'"
    Call ExecuteCommandSQL(cCONNECT, strSql)
    MsgBox "Se grabó correctamente", vbInformation, "PO Change"
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

