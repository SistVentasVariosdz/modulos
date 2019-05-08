VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAsignaNroDespacho 
   Caption         =   "Asignar Nro Despacho -PackOne"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNro_Despacho 
      Height          =   360
      Left            =   1410
      TabIndex        =   2
      Top             =   225
      Width           =   660
   End
   Begin FunctionsButtons.FunctButt funTemCli 
      Height          =   510
      Left            =   285
      TabIndex        =   0
      Top             =   885
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAsignaNroDespacho.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label1 
      Caption         =   "Número Despacho PackOne"
      Height          =   390
      Left            =   255
      TabIndex        =   1
      Top             =   195
      Width           =   885
   End
End
Attribute VB_Name = "frmAsignaNroDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sCod_ClientE As String
Public sCod_PurOrD As String
Public sCod_LotPurOrD As String
Public sCod_EstCli As String

Private Sub funTemCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        Grabar
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub Grabar()
On Error GoTo errx
Dim sSql As String

sSql = "EXEC SM_GRABA_NRO_DESPACHO_LOTEST '$','$','$','$', $"
sSql = VBsprintf(sSql, sCod_ClientE, sCod_PurOrD, sCod_LotPurOrD, sCod_EstCli, Val(txtNro_Despacho.Text))
ExecuteCommandSQL cCONNECT, sSql

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
Unload Me
Exit Sub
errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Public Sub CargaNroDespachoActual()
On Error GoTo errx
Dim sSql As String

sSql = "EXEC SM_CONSULTA_NRO_DESPACHO_LOTEST '$','$','$','$'"
sSql = VBsprintf(sSql, sCod_ClientE, sCod_PurOrD, sCod_LotPurOrD, sCod_EstCli)

txtNro_Despacho.Text = DevuelveCampo(sSql, cCONNECT)
Exit Sub
errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

