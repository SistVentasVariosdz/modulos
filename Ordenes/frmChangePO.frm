VERSION 5.00
Begin VB.Form frmChangePO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de PO"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChangePO.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Chage of PO"
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   2685
      TabIndex        =   10
      Tag             =   "&Cancel"
      Top             =   2355
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   735
      TabIndex        =   9
      Tag             =   "&Accept"
      Top             =   2370
      Width           =   1470
   End
   Begin VB.Frame fraActualiza 
      Height          =   720
      Left            =   30
      TabIndex        =   7
      Top             =   1470
      Width           =   5070
      Begin VB.TextBox txtNewPO 
         Height          =   315
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   8
         Top             =   225
         Width           =   2940
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nueva PO :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Tag             =   "New PO :"
         Top             =   300
         Width           =   840
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1500
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   5085
      Begin VB.TextBox txtEstilo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1005
         Width           =   2955
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   5
         Top             =   645
         Width           =   2940
      End
      Begin VB.TextBox txtPO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   4
         Top             =   285
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estilo :"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Tag             =   "Style :"
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Tag             =   "Client :"
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Tag             =   "PO :"
         Top             =   345
         Width           =   315
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
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   5175
   End
End
Attribute VB_Name = "frmChangePO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Public varCod_Cliente, varCod_EstCli, varCod_LotPurOrd, varCod_TemCli As String
Public bNivelPO As Boolean
Dim Rs_Lista As New ADODB.Recordset

Sub CAMBIO()
Dim oCone As ADODB.Connection
On Error GoTo AceptaErr
    
    Set oCone = New ADODB.Connection
    
    oCone.CursorLocation = adUseClient
    oCone.Open cCONNECT
    oCone.CommandTimeout = 4000
           
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = oCone
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    
    
    'Esta cadena es para cambiar el Po existente por uno nuevo
    If bNivelPO Then
        Strsql = "EXEC SP_CAMBIOS_PO_GENUPD_ALL_LOTES '" & varCod_Cliente & "','" & txtPO.Text & "','" & txtNewPO.Text & "'"
    Else
        Strsql = "EXEC SP_CAMBIOS_PO_GENUPD '" & varCod_Cliente & "','" & txtPO.Text & "','" & varCod_LotPurOrd & "','" & varCod_EstCli & "','" & txtNewPO.Text & "'"
    End If
    Rs_Lista.Open Strsql
    MsgBox "El cambio de PO fue exitoso", vbInformation, "PO Change"
    Unload Me
    Exit Sub
AceptaErr:
    ErrorHandler Err, "Error en Cambio"
End Sub

Private Function EXISTE_PO() As Boolean
    EXISTE_PO = False
    Strsql = "SELECT COUNT(*) FROM TG_PURORD WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_PurOrd = '" & Trim(txtNewPO.Text) & "'"
    If DevuelveCampo(Strsql, cCONNECT) > 0 Then
        EXISTE_PO = True
    End If
End Function

Private Sub cmdAceptar_Click()
    Dim opcion As Integer
    
    If Trim(txtNewPO.Text) = "" Then
        MsgBox "La PO no puede estar vacia. Sirvase verificar", vbInformation, "Change of PO"
        txtNewPO.Text = ""
        txtNewPO.SetFocus
        Exit Sub
    End If
    
    If EXISTE_PO Then
        opcion = MsgBox("La PO ya se encuentra registrada. ¿Desea incorporarla?", vbInformation + vbYesNo, "Change of PO")
        If opcion = vbYes Then
'            Strsql = "SELECT Cod_TemCli FROM TG_PURORD WHERE Cod_Cliente = '" & Me.varCod_Cliente & "' AND Cod_PurOrd = '" & Trim(Me.txtNewPO.Text) & "'"
'            If Trim(Me.varCod_TemCli) <> Trim(DevuelveCampo(Strsql, cCONNECT)) Then
'                MsgBox "No se puede copiar a una PO que pertenece a otra temporada. Sirvase verificar", vbInformation, "Mensaje"
'                Exit Sub
'            End If
            LblMsg.Visible = True
            'Aqui ejecutamos el store de cambio
            Call CAMBIO
            LblMsg.Visible = False
        End If
    Else
    
'        If Trim(Me.varCod_TemCli) <> Trim(DevuelveCampo(Strsql, cCONNECT)) Then
'            MsgBox "No se puede copiar a una PO que pertenece a otra temporada. Sirvase verificar", vbInformation, "Mensaje"
'            Exit Sub
'        End If
    
        'Aqui ejecutamos el store de cambio
        LblMsg.Visible = True
        Call CAMBIO
        LblMsg.Visible = False
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtNewPO.Text = ""
End Sub
