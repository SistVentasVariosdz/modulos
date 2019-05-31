VERSION 5.00
Begin VB.Form FrmMensaje 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   210
      Width           =   5265
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   5355
   End
End
Attribute VB_Name = "FrmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Titulo As String

Public Sub Refresca()
    Label1.Caption = Titulo
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
End Sub

