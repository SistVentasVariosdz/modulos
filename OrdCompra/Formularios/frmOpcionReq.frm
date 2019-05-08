VERSION 5.00
Begin VB.Form frmOpcionReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Inserción :"
   ClientHeight    =   2400
   ClientLeft      =   6015
   ClientTop       =   4650
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   465
      Left            =   2640
      TabIndex        =   2
      Top             =   1875
      Width           =   1125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   465
      Left            =   405
      TabIndex        =   1
      Top             =   1875
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Inserción :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   4125
      Begin VB.OptionButton Option3 
         Caption         =   "Añadir Requerimientos sobre Secuencia Actual"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Value           =   -1  'True
         Width           =   3690
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Añadir Requerimientos a los existentes"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   900
         Width           =   3210
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Crear nuevos Items de Requerimientos"
         Height          =   345
         Left            =   375
         TabIndex        =   3
         Top             =   375
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmOpcionReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmmaster As Object

Private Sub cmdAceptar_Click()
    If Option1.Value = True Then
        frmmaster.varAccion = "C"   'Crea nueva linea
    ElseIf Option2.Value = True Then
        frmmaster.varAccion = "A"   'Añade a las ya existentes
    Else
        frmmaster.varAccion = "I"   'Añade a la secuencia actual
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    frmmaster.varAccion = ""
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmaster = Nothing
End Sub
