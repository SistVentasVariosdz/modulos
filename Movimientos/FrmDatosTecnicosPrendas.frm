VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmDatosTecnicosPrendas 
   Caption         =   "Datos Tecnicos Prendas"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Caption         =   "Datos Técnicos Prenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   8355
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   6435
      End
      Begin VB.TextBox TxtGramaje 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtEncogLargo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtEncogAncho 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtRevirado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Gramaje:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encogimiento Largo:"
         Height          =   195
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Encogimiento Ancho:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Revirado:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8355
      Begin VB.TextBox TxtColor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4320
         TabIndex        =   19
         Top             =   555
         Width           =   3855
      End
      Begin VB.TextBox TxtComb 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   720
         TabIndex        =   17
         Top             =   555
         Width           =   3015
      End
      Begin VB.TextBox txtPartida 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   720
         TabIndex        =   0
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox TxtTela 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4320
         TabIndex        =   1
         Top             =   200
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Color:"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Comb.:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Partida :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Tela :"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDatosTecnicosPrendas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   2760
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDatosTecnicosPrendas.frx":009F
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmDatosTecnicosPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vCod_TipOrdTra As String
Public vCod_OrdTra As String
Public vCod_Tela As String
Public vDes_tela As String
Public vCod_Comb As String
Public vCod_Color As String
Public vPartida As String

Public Tipo As String
Dim strSQL As String

Public vOk As Boolean

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACTUALIZAR"
    Call Grabar
Case "CANCELAR"
    vOk = False
    Unload Me
End Select
End Sub

Sub Grabar()
On Error GoTo errGrabar
If Trim(TxtEncogAncho.Text) = "" Then
    TxtEncogAncho.Text = "0"
End If
If Trim(TxtEncogLargo.Text) = "" Then
    TxtEncogAncho.Text = "0"
End If
If Trim(TxtGramaje.Text) = "" Then
    TxtGramaje.Text = "0"
End If
If Trim(TxtRevirado.Text) = "" Then
    TxtRevirado.Text = "0"
End If
vOk = False

strSQL = "Up_Actualiza_Datos_Tecnicos_Tx_OrdTra_Telas '" & vCod_TipOrdTra & "','" & vCod_OrdTra & "','" & vCod_Tela & "','" & _
            vCod_Comb & "','" & vCod_Color & "'," & CDbl(TxtEncogAncho.Text) & "," & CDbl(TxtEncogLargo.Text) & "," & CDbl(TxtRevirado.Text) & "," & _
            Trim(TxtGramaje.Text) & ",'" & Trim(txtObservaciones.Text) & "'"
ExecuteSQL cConnect, strSQL
vOk = True

Unload Me
Exit Sub
errGrabar:
    vOk = False
    ErrorHandler err, "Grabar"
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Unload Me
End Sub

Private Sub TxtEncogAncho_GotFocus()
SelectionText TxtEncogAncho
End Sub

Private Sub TxtEncogAncho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtEncogLargo_GotFocus()
SelectionText TxtEncogLargo
End Sub

Private Sub TxtEncogLargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtGramaje_GotFocus()
SelectionText TxtGramaje
End Sub

Private Sub TxtGramaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtObservaciones_GotFocus()
SelectionText txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRevirado_GotFocus()
SelectionText TxtRevirado
End Sub

Private Sub TxtRevirado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
