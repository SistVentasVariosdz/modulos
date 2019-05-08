VERSION 5.00
Begin VB.Form frmSelReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Tag             =   "Select"
   Begin VB.OptionButton optTotal 
      Caption         =   "Total"
      Height          =   195
      Left            =   2565
      TabIndex        =   7
      Tag             =   "Total"
      Top             =   135
      Width           =   870
   End
   Begin VB.OptionButton optAgrupado 
      Caption         =   "Agrupado"
      Height          =   195
      Left            =   1275
      TabIndex        =   6
      Tag             =   "Group"
      Top             =   120
      Width           =   990
   End
   Begin VB.OptionButton optSimple 
      Caption         =   "Simple"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Tag             =   "Simple"
      Top             =   120
      Value           =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1815
      TabIndex        =   4
      Tag             =   "&Cancel"
      Top             =   1530
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   585
      TabIndex        =   3
      Tag             =   "&OK"
      Top             =   1530
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      TabIndex        =   0
      Tag             =   "Type"
      Top             =   420
      Width           =   3315
      Begin VB.OptionButton optImportes 
         Caption         =   "Detalle de Importes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   885
         TabIndex        =   2
         Tag             =   "Amount"
         Top             =   645
         Width           =   1770
      End
      Begin VB.OptionButton optPrendas 
         Caption         =   "Detalle de Prendas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   885
         TabIndex        =   1
         Tag             =   "Garment"
         Top             =   360
         Value           =   -1  'True
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmSelReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Private Sub cmdAceptar_Click()

If Me.optSimple.value = True Then
    oParent.Tipo_RepAcum = "S"
Else
    If Me.optAgrupado.value = True Then
        oParent.Tipo_RepAcum = "G"
    Else
        oParent.Tipo_RepAcum = "SI"
    End If
End If

If optPrendas.value = True Then
    oParent.Tipo_Rep = "C"
Else
    oParent.Tipo_Rep = "I"
End If
Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Call FormSet(Me)
End Sub

Private Sub optAgrupado_Click()
    Frame1.Enabled = True
End Sub

Private Sub optSimple_Click()
    Frame1.Enabled = True
End Sub

Private Sub optTotal_Click()
    optPrendas.value = True
    Frame1.Enabled = False
End Sub
