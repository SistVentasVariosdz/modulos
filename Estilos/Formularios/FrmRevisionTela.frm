VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmRevisionTela 
   Caption         =   "Revision Tela"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.TextBox TxtObservaciones 
         Height          =   1005
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox TxtCod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmRevisionTela.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmRevisionTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public vCod_Tela As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "GRABAR"
    Call Grabar
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub Grabar()
On Error GoTo errGrabar

strSQL = "tx_up_man_tx_telas_revisadas '" & vCod_Tela & "','" & vusu & "','" & Trim(TxtObservaciones.Text) & "'"
ExecuteSQL cCONNECT, strSQL
Unload Me
Exit Sub
errGrabar:
    MsgBox Err.Description, vbCritical, "Grabar"
End Sub

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub
