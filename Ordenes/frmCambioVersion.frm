VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmCambioVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Version plan ventas"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAccion 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   30
      TabIndex        =   9
      Top             =   1245
      Width           =   5850
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   248
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~ASIGNA~True~True~&Asigna~0~0~1~~0~False~False~&Asigna~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame fraPublica 
      Caption         =   "Datos para Cambio de Versión"
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5850
      Begin VB.TextBox txtcod_estpro 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDes_estpro 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   3435
      End
      Begin VB.TextBox txtCod_Version 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdBusVersion 
         Caption         =   "..."
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   300
      End
      Begin VB.TextBox txtDes_Version 
         Height          =   285
         Left            =   2330
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estilo :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Versión"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmCambioVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCod_Cliente As String
Public sCod_PurOrd As String
Public sCod_LotPurOrd As String
Public sCod_EstCli As String
Public sCod_EstPro As String
Public sDes_EstPro As String
Public codigo, Descripcion As String

Dim StrSql As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtcod_estpro.Text = sCod_EstPro
    txtDes_estpro.Text = sDes_EstPro
    txtCod_Version.Text = codigo
            StrSql = "SELECT Des_Version FROM ES_ESTPROVER WHERE Cod_EstPro = '" & sCod_EstPro & "' AND Cod_Version='" & Trim(txtCod_Version.Text) & "'"
            txtDes_Version.Text = DevuelveCampo(StrSql, cCONNECT)
    'txtCod_Version.SetFocus
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
            If Trim(txtCod_Version.Text) = "" Then
                MsgBox "Ingrese la version", vbInformation, "Aviso"
                txtCod_Version.SetFocus
                Exit Sub
            End If
            On Error GoTo errsalvar
            StrSql = "EXEC TG_UP_MAN_TG_LOTESTPRO_VERSION_VENTAS '" & Trim(sCod_Cliente) & "','" & Trim(sCod_PurOrd) & "','" & Trim(sCod_LotPurOrd) & "' ,'" & Trim(sCod_EstCli) & "','" & Trim(sCod_EstPro) & "','" & Trim(txtCod_Version.Text) & "'"
            Call ExecuteCommandSQL(cCONNECT, StrSql)
            Unload Me
        Exit Sub
errsalvar:
            ErrorHandler Err, "SALVAR_DATOS"
            Unload Me
End Sub

Private Sub txtCod_Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Version.Text) = "" Then
            Call BUSCA_VERSION
        Else
            StrSql = "SELECT Des_Version FROM ES_ESTPROVER WHERE Cod_EstPro = '" & sCod_EstPro & "' AND Cod_Version='" & Trim(txtCod_Version.Text) & "'"
            txtDes_Version.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
    End If
End Sub

Private Sub txtDes_Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtDes_Version.Text)) < 5 Then
            Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
            Exit Sub
        Else
            StrSql = "SELECT Cod_Version FROM ES_ESTPROVER WHERE Cod_EstPro = '" & sCod_EstPro & "' AND Des_Version LIKE '" & Trim(txtDes_Version.Text) & "%'"
            txtCod_Version.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
    End If
    
End Sub
Private Sub cmdBusVersion_Click()
    Call BUSCA_VERSION
End Sub
Public Sub BUSCA_VERSION()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Cod_Version AS Código, Des_Version as Descripción FROM ES_ESTPROVER WHERE Cod_EstPro = '" & sCod_EstPro & "'"
    oTipo.Cargar_Datos
    oTipo.DGridlista.Columns("Descripción").Width = 2654.929
    oTipo.Show 1
    If codigo <> "" Then
        txtCod_Version.Text = codigo
        txtDes_Version.Text = Descripcion
        FunctButt1.SetFocus
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub
