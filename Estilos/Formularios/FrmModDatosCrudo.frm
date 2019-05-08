VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmModDatosCrudo 
   Caption         =   "Modificacion Datos Crudo"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmModDatosCrudo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.TextBox TxtAnchoCrudo 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtGramajeCRudo 
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Crudo"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gramaje Crudo"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmModDatosCrudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public sCod_Tela As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call MODIFICA_DATOS_CRUDO
Case "CANCELAR"
    Unload Me
End Select
End Sub



Sub MODIFICA_DATOS_CRUDO()
On Error GoTo errModifica

If Trim(TxtGramajeCRudo.Text) = "" Then
    TxtGramajeCRudo.Text = "0"
End If

If Trim(TxtAnchoCrudo.Text) = "" Then
    TxtAnchoCrudo.Text = "0"
End If

strSQL = "up_modifica_datos_crudo_tela '" & sCod_Tela & "'," & CDbl(TxtGramajeCRudo.Text) & "," & CDbl(TxtAnchoCrudo.Text)
ExecuteSQL cCONNECT, strSQL
Unload Me
Exit Sub
errModifica:
    ErrorHandler Err, "Modificacion"

End Sub

Private Sub TxtAnchoCrudo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtAnchoCrudo, KeyAscii, True)
End If
End Sub

Private Sub TxtGramajeCRudo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtGramajeCRudo, KeyAscii, True)
End If
End Sub
