VERSION 5.00
Begin VB.Form frmCambioEstilo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Código de Estilo"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   2880
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1080
      TabIndex        =   10
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4935
      Begin VB.TextBox txtNuevo_Codigo 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   2100
         Width           =   1305
      End
      Begin VB.TextBox txtDes_EstCli 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Top             =   1170
         Width           =   3315
      End
      Begin VB.TextBox txtCod_EstCli 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtAbr_Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtDes_Cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   4
         Top             =   300
         Width           =   2445
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   4560
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Nuevo Código:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   2100
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Estilo Cliente:"
         Height          =   165
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Label LblMsg 
      Caption         =   "Este proceso puede tardar unos minutos. Espere por favor..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   5175
   End
End
Attribute VB_Name = "frmCambioEstilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vmarCodCliente, vmarAbrcliente, vmarNomCliente, vmarCodEstCli, vmarDesEstCli, vmarCodTem As String

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
If txtNuevo_Codigo.Text = "" Then
    MsgBox "Ingrese el Nuevo Código", vbCritical, Me.Caption
    Exit Sub
End If
LblMsg.Visible = True
    ActualizaCodigo
LblMsg.Visible = False
'    Set Rs = New ADODB.Recordset
'    Rs.ActiveConnection = cCONNECT
'    Rs.CursorType = adOpenStatic
'    Rs.CursorLocation = adUseClient
'    Rs.LockType = adLockReadOnly
'
'    Rs.Open "EXEC UP_Cambio_EstCliente_Tem '" & vmarCodCliente & "','" & vmarCodTem & "','" & vmarCodEstCli & "','" & txtNuevo_Codigo.Text & "'"
'    Set Rs = Nothing
    Unload Me
End Sub
Private Sub ActualizaCodigo()
Dim StrSQL As String
    Dim con As New ADODB.Connection
    On Error GoTo Actualizar_DatosErr

    Screen.MousePointer = vbHourglass
    
    con.ConnectionString = cCONNECT
    con.CommandTimeout = 900
    con.Open
    con.BeginTrans
    

    StrSQL = "EXEC UP_Cambio_EstCliente_Tem '" & vmarCodCliente & "','" & vmarCodTem & "','" & vmarCodEstCli & "','" & txtNuevo_Codigo.Text & "'"

    con.Execute StrSQL

    con.CommitTrans
    
    Screen.MousePointer = vbNormal
    Exit Sub
Actualizar_DatosErr:
    con.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Number & ", " & Err.Description, vbCritical, Me.Caption

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtAbr_Cliente.Text = vmarAbrcliente
txtDes_Cliente.Text = vmarNomCliente
txtCod_EstCli.Text = vmarCodEstCli
txtDes_EstCli.Text = vmarDesEstCli
End Sub


