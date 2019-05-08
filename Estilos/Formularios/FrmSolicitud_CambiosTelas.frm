VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmSolicitud_CambiosTelas 
   Caption         =   "Solicitud Modificación Telas"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2205
      TabIndex        =   5
      Top             =   1365
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~1~0~3~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2940
         TabIndex        =   8
         Top             =   210
         Width           =   3795
      End
      Begin VB.TextBox TxtCod_Tela 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1785
         TabIndex        =   7
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton CmdUsuarios 
         Caption         =   "..."
         Height          =   330
         Left            =   2940
         TabIndex        =   4
         Top             =   735
         Width           =   435
      End
      Begin VB.TextBox TxtDes_Usuario 
         Height          =   330
         Left            =   3465
         TabIndex        =   3
         Top             =   735
         Width           =   3270
      End
      Begin VB.TextBox TxtCod_Usuario 
         Height          =   330
         Left            =   1785
         TabIndex        =   2
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Autorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   840
         Width           =   1620
      End
   End
End
Attribute VB_Name = "FrmSolicitud_CambiosTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo As String
Public Descripcion As String

Private Sub CmdUsuarios_Click()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me
oTipo.sQuery = "SELECT Cod_Usuario as Codigo, Nom_Usuario as Descripcion FROM SEG_USUARIOS ORDER BY Cod_Usuario"
oTipo.Cargar_DatosSeguridad
oTipo.Show 1
If Codigo <> "" Then
    TxtCod_Usuario.Text = Codigo
    TxtDes_Usuario.Text = Descripcion
    Codigo = ""
    Descripcion = ""
    FunctButt1.SetFocus
End If
Set oTipo = Nothing
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If VALIDA = False Then Exit Sub
    Call APRUEBA_SOLICITUD
Case "CANCELAR"
Unload Me
End Select
End Sub

Sub APRUEBA_SOLICITUD()
On Error GoTo ErrSalvarDatos
strSQL = "TX_APRUEBA_SOLICITUDCAMBIOS '" & TxtCod_Tela & "','" & vusu & "','" & ComputerName & "','" & _
            TxtCod_Usuario & "'"
ExecuteCommandSQL cCONNECT, strSQL
Unload Me
Exit Sub
ErrSalvarDatos:
    ErrorHandler Err, "APRUEBA_SOLICITUD"
End Sub

Function VALIDA() As Boolean
If TxtCod_Usuario = "" Then
    MsgBox "Ingrese Usuarios Autorizado"
    VALIDA = False
    Exit Function
End If
VALIDA = True
End Function

Private Sub TxtCod_Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdUsuarios_Click
End If
End Sub
