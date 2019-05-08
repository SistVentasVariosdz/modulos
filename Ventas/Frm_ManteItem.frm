VERSION 5.00
Begin VB.Form Frm_ManteItem 
   Caption         =   "Mantenimiento de Items"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
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
      Height          =   480
      Left            =   2760
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Height          =   480
      Left            =   3997
      TabIndex        =   2
      Tag             =   "&Cancel"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8340
      Begin VB.TextBox txtUnida_Medida 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   6405
      End
      Begin VB.Label Label2 
         Caption         =   "Unidad Medida :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   930
      End
   End
End
Attribute VB_Name = "Frm_ManteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strSQL As String

Dim sAccion As String

Private Sub cmdAceptar_Click()
    Call Grabar
    Frm_DetallaeItems.txtCod_Producto.Text = DevuelveCampo("SELECT COD_ITEM FROM LG_ITEM WHERE DES_ITEM='" & txtDescripcion.Text & "' AND Cod_UniMed='" & txtUnida_Medida.Text & "'", cCONNECT)
    Frm_DetallaeItems.txtDescripcion.Text = DevuelveCampo("SELECT DES_ITEM   FROM LG_ITEM WHERE DES_ITEM='" & txtDescripcion.Text & "' AND Cod_UniMed='" & txtUnida_Medida.Text & "'", cCONNECT)
    Frm_DetallaeItems.sunidad = DevuelveCampo("SELECT Cod_UniMed  FROM LG_ITEM WHERE DES_ITEM='" & txtDescripcion.Text & "' AND Cod_UniMed='" & txtUnida_Medida.Text & "'", cCONNECT)
    Unload Me
End Sub

Sub Grabar()
Dim sRows As Integer
On Error GoTo hand

strSQL = "EXEC LG_ITEM_CREA_ITEM_DIVERSO '" & Trim(txtDescripcion.Text) & "','" & UCase(Trim(txtUnida_Medida.Text)) & "'"
    
Call ExecuteSQL(cCONNECT, strSQL)
          
Exit Sub
hand:
    ErrorHandler err, "Mantenimiento de Item"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtUnida_Medida.SetFocus

End Sub

Private Sub txtUnida_Medida_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdAceptar.SetFocus
End Sub
