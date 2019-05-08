VERSION 5.00
Begin VB.Form frmAddGrupoPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicion de Grupo de Producción"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGrupoPro 
      Caption         =   "Adicionar Grupo de Producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   5115
      Begin VB.TextBox txtAbr_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCod_GrupoPro 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtDes_GrupoPro 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdCancelarGP 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   2820
         TabIndex        =   9
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarGP 
         Caption         =   "A&ceptar"
         Height          =   495
         Left            =   1335
         TabIndex        =   8
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox txtNom_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2175
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdBuscaCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   290
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   650
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAddGrupoPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Public oParent As Object
Public Codigo, Descripcion As String

Private Sub ANADE_GRUPOPRO()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        Strsql = "SELECT COD_CLIENTE FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        Strsql = "EXEC UP_MAN_GRUPOPRO '" & _
        "I" & "','" & _
        DevuelveCampo(Strsql, cCONNECT) & "','" & _
        Trim(txtCod_GrupoPro.Text) & "','" & _
        Trim(txtDes_GrupoPro.Text) & "'"
  
    Con.Execute Strsql
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = MESSAGECODE.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
  
    Set Con = Nothing
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al añadir el Color de Estilo Propio", vbCritical)
End Sub


Private Function VALIDA_ANADE_GRUPOPRO() As Boolean
    VALIDA_ANADE_GRUPOPRO = True
    If txtAbr_Cliente.Text = "" Then
        Call MsgBox("Sirvase seleccionar un Cliente", vbInformation)
        txtAbr_Cliente.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    If txtCod_GrupoPro.Text = "" Then
        Call MsgBox("El código de grupo no puede estar vacio. Sirvase verificar", vbInformation)
        txtCod_GrupoPro.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    If txtDes_GrupoPro.Text = "" Then
        Call MsgBox("La descripción del grupo no puede estar vacia. Sirvase verificar", vbInformation)
        txtDes_GrupoPro.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    Strsql = "SELECT COUNT(COD_CLIENTE) FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    If DevuelveCampo(Strsql, cCONNECT) = 0 Then
        Call MsgBox("El cliente ingresado no se encuentra registrado. Sirvase verificar", vbInformation)
        txtAbr_Cliente.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
End Function

Private Sub cmdBuscaCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.RecordSet
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If Codigo <> "" Then
        txtAbr_Cliente.Text = Codigo
        txtNom_Cliente.Text = Descripcion
        txtCod_GrupoPro.SetFocus
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdCancelarGP_Click()
    Unload Me
End Sub

Private Sub cmdAceptarGP_Click()
    If VALIDA_ANADE_GRUPOPRO Then
        Call ANADE_GRUPOPRO
        
        oParent.txtCod_GrupoPro.Text = Me.txtCod_GrupoPro.Text
        oParent.txtDes_GrupoPro.Text = Me.txtDes_GrupoPro.Text
        
        Unload Me
        
        'Call CARGA_GRUPOPRO
        'stipo = ""
        'fraGrupoPro.Visible = False
    End If
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBuscaCliente_Click
        Else
            Strsql = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente ='" & Trim(txtAbr_Cliente.Text) & "'"
            txtNom_Cliente.Text = DevuelveCampo(Strsql, cCONNECT)
            txtCod_GrupoPro.SetFocus
        End If
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente.Text) > 4 Then
            Strsql = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtAbr_Cliente.Text = DevuelveCampo(Strsql, cCONNECT)
            If Trim(txtAbr_Cliente.Text) <> "" Then
                txtAbr_Cliente_KeyPress (13)
            End If
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
    End If
End Sub
