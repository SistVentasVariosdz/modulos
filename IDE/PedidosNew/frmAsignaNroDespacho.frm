VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
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
      Caption         =   "N�mero Despacho PackOne"
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

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Private Sub funTemCli_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"
            Grabar

        Case "CANCELAR"
            Unload Me
    End Select

End Sub

Private Sub Grabar()

    On Error GoTo errx

    Dim sSQl As String

    sSQl = "EXEC SM_GRABA_NRO_DESPACHO_LOTEST '$','$','$','$', $"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, Val(txtNro_Despacho.Text))
    ExecuteCommandSQL cCONNECT, sSQl

    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Unload Me

    Exit Sub

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Public Sub CargaNroDespachoActual()

    On Error GoTo errx

    Dim sSQl As String

    sSQl = "EXEC SM_CONSULTA_NRO_DESPACHO_LOTEST '$','$','$','$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)

    txtNro_Despacho.Text = DevuelveCampo(sSQl, cCONNECT)

    Exit Sub

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

