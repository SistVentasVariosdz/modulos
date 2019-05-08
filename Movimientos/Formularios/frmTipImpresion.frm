VERSION 5.00
Begin VB.Form frmTipImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Guia"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   795
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   2685
      Begin VB.OptionButton Option2 
         Caption         =   "Remision"
         Height          =   210
         Left            =   1455
         TabIndex        =   2
         Top             =   210
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Recojo"
         Height          =   240
         Left            =   210
         TabIndex        =   1
         Top             =   195
         Value           =   -1  'True
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmTipImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sTipImpresion As String
Public oParent As Object

Private Sub Command1_Click()
    If Option1.Value Then
        sTipImpresion = "RECOJO"
    Else
        sTipImpresion = "REMISION"
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    sTipImpresion = "RECOJO"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oParent.sTipImpresion = Me.sTipImpresion
End Sub
