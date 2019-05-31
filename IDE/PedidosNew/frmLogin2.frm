VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2535
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5280
   FillColor       =   &H80000011&
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1497.762
   ScaleMode       =   0  'User
   ScaleWidth      =   4957.634
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin MSDataListLib.DataCombo DCboEmpresas 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483637
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000005&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1140
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   1695
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3120
      TabIndex        =   5
      Tag             =   "Cancel"
      Top             =   1680
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Tag             =   "Company :"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "User Name  :"
      Top             =   135
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Tag             =   "Password :"
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    'LoginSucceeded = False
    Unload Me
End Sub
Private Function Valida_usuario(ByVal xusr As Variant, ByVal xpas As Variant, ByVal xemp As Variant)
If Len(xemp) = 0 Then
  SQUERY = "SELECT count(*) from seg_usuarios  WHERE COD_USUARIO='" & xusr & "'"
Else
  SQUERY = "SELECT cod_perfil from seg_empusuper a,seg_usuarios b WHERE  a.cod_empresa='" & xemp & "' and a.cod_usuario=b.cod_usuario and b.COD_USUARIO='" & xusr & "'"
End If
If xpas = "" Then
   xcondi = " and (password is null or password='')"
Else
   xcondi = "  and password='" & xpas & "'"
End If
SQUERY = SQUERY & xcondi
Set RS1 = New ADODB.RecordSet
RS1.ActiveConnection = conn
RS1.CursorType = adOpenStatic
RS1.Open SQUERY
If Not (RS1.BOF And RS1.EOF) Then
  If RS1(0) > 0 Then
     Valida_usuario = RS1(0)
  Else
     Valida_usuario = ""
  End If
End If
Set RS1 = Nothing
End Function
Private Sub cmdOK_Click()
vusr = txtUserName
vpas = txtPassword
If DCboEmpresas.Enabled = False Then
  vu = Valida_usuario(vusr, vpas, "")
  If Len(vu) > 0 Then
   scarga = Carga_Empresas()
   If scarga Then
    DCboEmpresas.Enabled = True
    DCboEmpresas.BackColor = &H80000005
   Else
    MsgBox "Usuario no registrado en Empresa", , "Inicio de sesión"
    txtPassword.SetFocus
   End If
  Else
    MsgBox "Usuario o clave no Validos", , "Inicio de sesión"
    txtUserName.SetFocus
  End If
Else
  vu = Valida_usuario(vusr, vpas, DCboEmpresas.BoundText)
  If Len(vu) > 0 Then
      With MDIPrincipal
      ' With MdiPrueba
        .PUsuario = vusr
        .PClave = vpas
        .PEmpresa = DCboEmpresas.BoundText
        .NEmpresa = DCboEmpresas.Text
        .perfil = vu
        Datos_Empresa DCboEmpresas.BoundText
       End With
    Unload Me
    MDIPrincipal.Show
  Else
    MsgBox "La contraseña o el usuario no son válidos o no registrado en Empresa. Vuelva a intentarlo", , "Inicio de sesión"
    txtPassword.SetFocus
  End If
End If
End Sub

Private Function Carga_Empresas()
    SQUERY = "SELECT A.COD_EMPRESA AS CODIGO,B.DES_EMPRESA AS NOMBRE,RUTA_LOGO,NUM_RUC,DIRECCION,DSN,RUTA0 FROM SEG_EMPUSUPER A,SEG_EMPRESAS B WHERE A.COD_EMPRESA=B.COD_EMPRESA AND A.COD_USUARIO='" & txtUserName & "'"
    Set mRs = New ADODB.RecordSet
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open SQUERY
    icount = mRs.RecordCount
    icodini = "00"
    If icount > 0 Then
       icodini = mRs(0)
       Set DCboEmpresas.RowSource = mRs
       DCboEmpresas.ListField = "NOMBRE"
       DCboEmpresas.BoundColumn = "CODIGO"
       DCboEmpresas.BoundText = icodini
       Carga_Empresas = True
    Else
       Carga_Empresas = False
    End If
    Set mRs = Nothing
End Function
Private Function Datos_Empresa(ByVal codemp As Variant)
    SQUERY = "SELECT ISNULL(RUTA_LOGO,'') AS RUTA_LOGO,ISNULL(NUM_RUC,'') AS NUM_RUC,ISNULL(DIRECCION,'') AS DIRECCION,ISNULL(DSN,'') AS DSN,ISNULL(RUTA0,'') AS RUTA0 , ISNULL(DSNSEGURIDAD,'') AS DSNSEGURIDAD FROM SEG_EMPRESAS  WHERE COD_EMPRESA='" & codemp & "'"
    Set mRs = New ADODB.RecordSet
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open SQUERY
    icount = mRs.RecordCount
    If icount > 0 Then
     Ruta_Logo_Empresa = mRs(0)
     Num_Ruc_Empresa = mRs(1)
     Direccion_Empresa = mRs(2)
     DSN_Empresa = mRs(3)
     Ruta0_Empresa = mRs(4)
     DSN_Seguridad = mRs(5)
    End If
    Set mRs = Nothing
    Fecha_Hora_Conexion = Now()
End Function

Private Sub Form_Load()
LoadConnectEmpresa ""
IdiomaEtiquetas1 Me
End Sub
