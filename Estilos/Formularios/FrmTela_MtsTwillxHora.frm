VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmTela_MtsTwillxHora 
   Caption         =   "Mts. Twill por Hora"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   585
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmTela_MtsTwillxHora.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1070
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.TextBox TxtTwill 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mts. Twill x Hora"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   680
         Width           =   1170
      End
      Begin VB.Label LblDes_Tela 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   4155
      End
      Begin VB.Label LblCod_Tela 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   315
      End
   End
End
Attribute VB_Name = "FrmTela_MtsTwillxHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public sCod_Tela As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "MODIFICAR"
    Call Cambia_Twill
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub Cambia_Twill()
On Error GoTo errCambia_Twill
strSQL = "up_cambia_tx_tela_Twill_x_hora '" & sCod_Tela & "'," & IIf(Val(TxtTwill.Text) = 0, 0, CDbl(TxtTwill.Text))
ExecuteSQL cCONNECT, strSQL

MsgBox "Cambios Realizados correctamente"
Unload Me
Exit Sub
errCambia_Twill:
    ErrorHandler Err, "Cambia Twill"
End Sub

Private Sub TxtTwill_GotFocus()
SelectionText TxtTwill
End Sub

Private Sub TxtTwill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtTwill, KeyAscii, True, 2)
End If
End Sub
