VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Pedidos"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "Ò.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Log On"
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   4545
      TabIndex        =   5
      Tag             =   "&Cancel"
      Top             =   2340
      Width           =   1245
   End
   Begin MSDataListLib.DataCombo DCboEmpresas 
      Height          =   315
      Left            =   3855
      TabIndex        =   3
      Top             =   1800
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000005&
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   3240
      TabIndex        =   4
      Tag             =   "&Ok"
      Top             =   2340
      Width           =   1245
   End
   Begin VB.TextBox TxtUserName 
      Height          =   315
      Left            =   3855
      TabIndex        =   0
      Top             =   975
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3855
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1395
      Width           =   2325
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2385
      TabIndex        =   9
      Tag             =   "Manufacturing Management"
      Top             =   120
      Width           =   3945
   End
   Begin VB.Line Line1 
      X1              =   2340
      X2              =   6315
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   60
      Picture         =   "Ò.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   210
      Picture         =   "Ò.frx":1E46
      Stretch         =   -1  'True
      Top             =   510
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   2385
      TabIndex        =   8
      Tag             =   "Copyright Release 7.7.0"
      Top             =   480
      Width           =   3930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   255
      Left            =   2505
      TabIndex        =   7
      Tag             =   "Company"
      Top             =   1845
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   225
      Left            =   2505
      TabIndex        =   6
      Tag             =   "Password"
      Top             =   1425
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   2505
      TabIndex        =   1
      Tag             =   "User"
      Top             =   1020
      Width           =   1245
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Option Explicit
Public LoginSucceeded As Boolean
Public bOk As Boolean

Private Function Valida_usuario(ByVal xusr As Variant, ByVal xpas As Variant, ByVal xemp As Variant)

'conn = "Provider=sqloledb;Server=HIALPESA4;Database=HIALPESA;Integrated Security=SSPI;"
'conn = "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"


If Len(xemp) = 0 Then
  SQuery = "SELECT count(*) from seg_usuarios  WHERE COD_USUARIO='" & xusr & "'"
Else
  SQuery = "SELECT cod_perfil from seg_empusuper a,seg_usuarios b WHERE  a.cod_empresa='" & xemp & "' and a.cod_usuario=b.cod_usuario and b.COD_USUARIO='" & xusr & "'"
End If
If xpas = "" Then
   xcondi = " and (password is null or password='')"
Else
   xcondi = "  and password='" & xpas & "'"
End If
SQuery = SQuery & xcondi
Set RS1 = New ADODB.Recordset
RS1.ActiveConnection = conn
'RS1.ActiveConnection = sconnect
RS1.CursorType = adOpenStatic
RS1.Open SQuery
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
Static ICONT As Byte
vusr = TxtUserName
vpas = txtPassword
If ICONT = 3 Then End
If DCboEmpresas.Enabled = False Then
  vu = Valida_usuario(vusr, vpas, "")
  If Len(vu) > 0 Then
   scarga = Carga_Empresas()
   If scarga Then
    DCboEmpresas.Enabled = True
    DCboEmpresas.BackColor = &H80000005
   Else
    ICONT = ICONT + 1
    MsgBox "Usuario no registrado en Empresa", , "Inicio de sesión"
    txtPassword.SetFocus
   End If
  Else
    ICONT = ICONT + 1
    MsgBox "Usuario o clave no Validos", vbInformation, "Inicio de sesión"
    TxtUserName = ""
    TxtUserName.SetFocus
  End If
Else
  vu = Valida_usuario(vusr, vpas, DCboEmpresas.BoundText)
  
  If Len(vu) > 0 Then
      Call RegistrarAcceso
      With MDIPrincipal
      ' With MdiPrueba
        .pUsuario = vusr
        .PClave = vpas
        .pEmpresa = DCboEmpresas.BoundText
        .NEmpresa = DCboEmpresas.Text
        .perfil = vu
         bOk = Datos_Empresa(DCboEmpresas.BoundText)
        
       End With
    Unload Me
    If bOk Then
        MDIPrincipal.Show
    End If
  Else
    ICONT = ICONT + 1
    MsgBox "La contraseña o el usuario no son válidos o no registrado en Empresa. Vuelva a intentarlo", vbInformation, "Inicio de sesión"
    txtPassword.SetFocus
  End If
End If
End Sub
Private Sub RegistrarAcceso()
Dim sSQl As String
    sSQl = "EXEC Seg_RegistarAccesoUsuario '$','$'"
    sSQl = VBsprintf(sSQl, TxtUserName.Text, ComputerName)
    ExecuteCommandSQL sconnect, sSQl
End Sub
Private Function Carga_Empresas()
    SQuery = "SELECT A.COD_EMPRESA AS CODIGO,B.DES_EMPRESA AS NOMBRE,RUTA_LOGO,NUM_RUC,DIRECCION,DSN,RUTA0 FROM SEG_EMPUSUPER A,SEG_EMPRESAS B WHERE A.COD_EMPRESA=B.COD_EMPRESA AND A.COD_USUARIO='" & TxtUserName & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open SQuery
    iCount = mRs.RecordCount
    icodini = "00"
    If iCount > 0 Then
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
Private Function Datos_Empresa(ByVal codemp As Variant) As Boolean
Dim serror As String, strSql As String
        On Error GoTo errsalvar
    SQuery = "SELECT ISNULL(RUTA_LOGO,'') AS RUTA_LOGO,ISNULL(NUM_RUC,'') AS NUM_RUC,ISNULL(DIRECCION,'') AS DIRECCION,ISNULL(DSN,'') AS DSN,ISNULL(RUTA0,'') AS RUTA0 , ISNULL(DSNSEGURIDAD,'') AS DSNSEGURIDAD FROM SEG_EMPRESAS  WHERE COD_EMPRESA='" & codemp & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open SQuery
    iCount = mRs.RecordCount
    If iCount > 0 Then
     Ruta_Logo_Empresa = mRs(0)
     Num_Ruc_Empresa = mRs(1)
     Direccion_Empresa = mRs(2)
     DSN_Empresa = mRs(3)
     'Para pruebas
     'DSN_Empresa = "Provider=sqloledb;Server=" & GetSetting("Visuales", "Settings", "Server") & "Database=LIVES;UID=sa;pwd=;"
     'DSN_Empresa = "Provider=sqloledb;Server=ECARDENAS\ECN;Database=HIALPESAX;Integrated Security=SSPI;"
     'DSN_Empresa = "Provider=sqloledb;Server=VRIOS2;Database=HIALPESAX;Integrated Security=SSPI;"
     'DSN_Empresa = "Provider=sqloledb;Server=BESCALANTE\SQL2005;Database=HIALPESA;Integrated Security=SSPI;"
     
     'DSN_Empresa = "Provider=sqloledb;Server=wflores2;Database=inka;Integrated Security=SSPI;"
     
     '''''DSN_Empresa = "Provider=sqloledb;Server=CESARATOCHE2\SQL2005;Database=INKADESIGNS;Integrated Security=SSPI;"
     
     '''''''''DSN_Empresa = "Provider=sqloledb;Server=HIALPESA4;Database=HIALPESA;Integrated Security=SSPI;"
     
     '''''DSN_Empresa = "Provider=sqloledb;Server=CATOCHE\SQL2008R2;Database=HIALPESA;Integrated Security=SSPI;"
     
             
     Ruta0_Empresa = mRs(4)
     DSN_Seguridad = mRs(5)
     cCONNECT = DSN_Empresa

            strSql = "EXEC SEG_MAN_BITACORA_ACCESO '" & vusu & "','" & ComputerName & "','" & vemp & "'"
            Call ExecuteCommandSQL(DSN_Seguridad, strSql)
'            serror = DevuelveCampo(strSql, DSN_Seguridad)
'            If serror = "N" Then
'                MsgBox "COMPUTADORA BLOQUEDA POR ACCESO INAPROPIADO. USUARIO " & vusu & " GENERO BLOQUEO PARA ESTA PC."
'               End
'            End If

    End If
    Set mRs = Nothing
    Fecha_Hora_Conexion = Now()
    Datos_Empresa = True
    Exit Function
errsalvar:
            ErrorHandler Err, "SALVAR_DATOS"
            Unload Me
End Function

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
'LoadConnectEmpresa ""
Dim sBuffer As String, Ret As Long
sBuffer = String(256, 0)
Ret = Len(sBuffer)

'''conn = "Provider=sqloledb;Server=CESARATOCHE2\SQL2005;Database=INKADESIGNS;Integrated Security=SSPI;"


''''''''conn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=INKADESIGNS;Data Source=CESARATOCHE2\SQL2005"
'''''''''''''''''''conn.ConnectionString = "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"
'conn = "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"
'conn.Open "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"
'sconnect = "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"

IdiomaEtiquetas1 Me
Label4.Caption = "Copyright Release " & App.Major & "." & App.Minor & "." & App.Revision

'If GetUserNameEx(NameSamCompatible, sBuffer, Ret) <> 0 Then
TxtUserName = ComputerName
' Mid(sBuffer, InStr(sBuffer, "\") + 1, 20)
'End If
End Sub


Private Sub Timer1_Timer()
Static Estado
If Estado = Empty Then
    Image1.Visible = True
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Estado = 2
ElseIf Estado = 2 Then
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = False
    Image4.Visible = False
    Estado = 3

ElseIf Estado = 3 Then
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = True
    Image4.Visible = False
    Estado = 4

ElseIf Estado = 4 Then
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = True
    Estado = Empty
End If

End Sub






Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub TxtUserName_GotFocus()
    SelectionText TxtUserName
End Sub

Private Sub TxtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
    'cmdOK_Click
End If
End Sub

