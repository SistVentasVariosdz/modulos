VERSION 5.00
Begin VB.Form frmAsignaOCorte 
   Caption         =   "Adiciona Orden de Corte"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   3060
      Begin VB.TextBox txtCo_CodOrdPro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden Corte:"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   1665
      TabIndex        =   4
      Top             =   930
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   195
      TabIndex        =   3
      Top             =   930
      Width           =   1320
   End
End
Attribute VB_Name = "frmAsignaOCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Public varCOD_ALMACEN As String
Public varNUM_MOVSTK As String

Private Sub cmCancelar_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    Me.txtCo_CodOrdPro.Text = Right("00000" & Trim(txtCo_CodOrdPro.Text), 5)
    Strsql = "SELECT COUNT(*) FROM CO_ORDPRO WHERE Co_CodOrdPro = '" & Trim(Me.txtCo_CodOrdPro.Text) & "'"
    If DevuelveCampo(Strsql, cConnect) = 0 Then
        MsgBox "La Orden de Corte Ingresada no existe. Sirvase verificar", vbInformation, "Mensaje"
        Exit Sub
    Else
        Call SALVAR_DATOS
        Unload Me
    End If
End Sub


Private Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim i As Integer
    On Error GoTo Salvar_DatosErr
    Dim Strsql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

            Strsql = "EXEC UP_ADICIONA_TELAS_CORTE_EN_MOVIMIENTO '" & _
            Me.txtCo_CodOrdPro.Text & "','" & _
            Me.varCOD_ALMACEN & "','" & _
            Me.varNUM_MOVSTK & "'"
        
            Con.Execute Strsql

        Con.CommitTrans
        MsgBox "La asignación se realizó con éxito", vbInformation, "Mensaje"
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub txtCo_CodOrdPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCo_CodOrdPro.Text = Right("00000" & Trim(txtCo_CodOrdPro.Text), 5)
        Me.cmdAceptar.SetFocus
    End If
End Sub
