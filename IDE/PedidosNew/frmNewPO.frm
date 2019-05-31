VERSION 5.00
Begin VB.Form frmNewPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrese PO Destino"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "New PO"
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
      Height          =   465
      Left            =   3330
      TabIndex        =   6
      Tag             =   "Cancel"
      Top             =   1590
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
      Height          =   465
      Left            =   2100
      TabIndex        =   5
      Tag             =   "&OK"
      Top             =   1590
      Width           =   1185
   End
   Begin VB.TextBox txtPO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   4
      Top             =   975
      Width           =   3120
   End
   Begin VB.TextBox txtIdCliente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3330
      TabIndex        =   1
      Top             =   570
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtNomCliente 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1365
      TabIndex        =   0
      Top             =   555
      Width           =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "PO Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   2
      Tag             =   "New PO"
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   3
      Tag             =   "Client"
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmNewPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public oParent As Object

Private Sub cmdAceptar_Click()

    If Not ValidaPO() Then

        Exit Sub

    End If

    oParent.sPONew = txtPO.Text
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
End Sub

Function ValidaPO() As Boolean

    Dim sQuery As String

    ValidaPO = True

    On Error GoTo ValidaPOErr

    sQuery = "SELECT count(*) FROM TG_PurOrd WHERE cod_cliente = '" & txtIdCliente.Text & "' AND cod_purord = '" & txtPO.Text & "'"

    If DevuelveCampo(sQuery, cCONNECT) > 0 Then
        MsgBox "El PO que ingreso ya se ha registrado." & vbCr & "Ingrese nuevo PO", vbInformation, "Copiar PO"
        ValidaPO = False
    End If

    Exit Function

ValidaPOErr:
    ErrorHandler Err, "ValidaPO"
    ValidaPO = False
End Function
