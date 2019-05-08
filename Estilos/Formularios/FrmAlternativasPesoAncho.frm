VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmAlternativasPesoAncho 
   Caption         =   "Alternativas Peso / Ancho"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAlternativasPesoAncho.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame FraDatos 
      Height          =   1455
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7335
      Begin VB.TextBox TxtAncho 
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtGramaje 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox TxtAlternativa 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox TxtCod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gramaje"
         Height          =   195
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alternativa"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   315
      End
   End
End
Attribute VB_Name = "FrmAlternativasPesoAncho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Tela As String, vAccion As String, oParent As Object
Dim strSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "CANCELAR"
    Unload Me
End Select
End Sub

Private Sub TxtAncho_GotFocus()
SelectionText TxtAncho
End Sub

Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtAncho, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtGramaje_GotFocus()
SelectionText TxtGramaje
End Sub

Private Sub TxtGramaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtGramaje, KeyAscii, False)
End If
End Sub

Sub Grabar()
On Error GoTo errGrabar

If Trim(TxtGramaje.Text) = "" Then
    TxtGramaje.Text = 0
End If
If Trim(TxtAncho.Text) = "" Then
    TxtAncho.Text = 0
End If

strSQL = "tx_up_man_tx_tela_alternativas '" & vAccion & "','" & vCod_Tela & "'," & _
            Val(TxtAlternativa.Text) & ",'" & Trim(TxtDescripcion.Text) & "'," & _
            CDbl(TxtGramaje.Text) & "," & CDbl(TxtAncho.Text)
            
ExecuteCommandSQL cCONNECT, strSQL

oParent.CARGA_GRID
Unload Me

Exit Sub
errGrabar:
    MsgBox Err.Description, vbCritical, "Grabar"
End Sub
