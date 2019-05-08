VERSION 5.00
Begin VB.Form frmCambioColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de còdigo de color"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   4755
      Begin VB.TextBox txtCod_Color 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   12
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox txtNom_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         TabIndex        =   8
         Top             =   510
         Width           =   2685
      End
      Begin VB.TextBox txtDes_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         TabIndex        =   7
         Top             =   210
         Width           =   2685
      End
      Begin VB.TextBox txtCod_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   6
         Top             =   510
         Width           =   795
      End
      Begin VB.TextBox txtAbr_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txtNuevo_Codigo 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "Nuevo Código:"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   4560
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1020
      TabIndex        =   1
      Top             =   2490
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   2580
      TabIndex        =   3
      Top             =   2490
      Width           =   1335
   End
End
Attribute VB_Name = "frmCambioColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Cliente, vCod_TemCli, vCod_ColCli As String

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
If txtNuevo_Codigo.Text = "" Then
    MsgBox "Ingrese el Nuevo Código", vbCritical, Me.Caption
    Exit Sub
End If
    ActualizaCodigo

    Unload Me
End Sub
Private Sub ActualizaCodigo()
Dim StrSQL As String
    Dim Con As New ADODB.Connection
    On Error GoTo Actualizar_DatosErr

    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans

        StrSQL = "EXEC UP_Cambio_EstCliCol_Tem '" & vCod_Cliente & "','" & vCod_TemCli & "','" & vCod_ColCli & "','" & txtNuevo_Codigo.Text & "'"

        Con.Execute StrSQL

    Con.CommitTrans
    Exit Sub
Actualizar_DatosErr:
    Con.RollbackTrans
    MsgBox Err.Number & ", " & Err.Description, vbCritical, Me.Caption

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Rs As New ADODB.Recordset
Dim StrSQL As String
    
    Rs.Open "select abr_cliente,nom_cliente from tg_cliente where cod_cliente='" & vCod_Cliente & "'", cCONNECT, adOpenStatic
    If Rs.RecordCount > 0 Then
        txtAbr_Cliente = Rs.Fields("abr_cliente").Value
        txtDes_Cliente = Rs.Fields("nom_cliente").Value
    End If
    
    StrSQL = "select nom_temcli from tg_temcli where cod_cliente='" & vCod_Cliente & "' and cod_temcli='" & vCod_TemCli & "'"
    txtCod_TemCli = vCod_TemCli
    txtNom_TemCli = DevuelveCampo(StrSQL, cCONNECT)
    txtCod_Color.Text = vCod_ColCli
End Sub
