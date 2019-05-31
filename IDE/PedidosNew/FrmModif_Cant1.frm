VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmModif_Cant1 
   Caption         =   "Modificar Cantidad"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmModif_Cant1.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox TxtCantidad 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FrmModif_Cant1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public oParent        As Object

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Public scod_colcli    As String

Public sCod_Talla     As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "MODIFICAR"
            MODIFICAR

        Case "CANCELAR"
            Unload Me
    End Select

End Sub

Private Sub MODIFICAR()

    On Error GoTo DeleteErr

    Dim strSql As String

    strSql = "TG_PURORD_MOD_CANT_REQUERIDA_LOTCOLTAL '" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "','" & scod_colcli & "','" & sCod_Talla & "','" & TxtCantidad & "'"
 
    ExecuteCommandSQL cCONNECT, strSql
    ''oParent.cargar
    MsgBox "Se Modificó Correctamante"
    Unload Me

    Exit Sub

DeleteErr:
    errores Err.Number
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then FunctButt1.SetFocus
End Sub
