VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{7B0D986D-3A03-4634-828F-D16994E0941A}#3.0#0"; "ECNVB6WINCTRL.ocx"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de gestión textil"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLogin.frx":6852
   ScaleHeight     =   3105
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Log On"
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6435
      Begin ECNVB6WINCTRL.ucLabel lblSistema 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   765
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         Caption         =   "SISTEMA DE GESTIÓN TEXTIL"
         ForeColor       =   0
         BackColor       =   16777215
         ShadowColor     =   6710886
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   1050
         Left            =   5070
         Picture         =   "FrmLogin.frx":36B13
         Top             =   0
         Width           =   1320
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   120
         Picture         =   "FrmLogin.frx":37D34
         Top             =   0
         Width           =   1905
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6380
         Y1              =   1110
         Y2              =   1110
      End
   End
   Begin MSDataListLib.DataCombo DCboEmpresas 
      Height          =   315
      Left            =   4050
      TabIndex        =   3
      Top             =   2130
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtUserName 
      Height          =   315
      Left            =   4065
      TabIndex        =   0
      Top             =   1260
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4065
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2325
   End
   Begin ECNVB6WINCTRL.ucButton_02 cmdOK 
      Height          =   435
      Left            =   3660
      TabIndex        =   9
      ToolTipText     =   "Hacer clic para ingresar al sistema."
      Top             =   2610
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   767
      Icon            =   "FrmLogin.frx":38A2C
      Style           =   5
      Caption         =   "    &Ingresar"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   "Hacer clic para ingresar al sistema."
      ToolTipTitle    =   "Ingresar al sistema"
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin ECNVB6WINCTRL.ucButton_02 Command1 
      Height          =   435
      Left            =   5085
      TabIndex        =   10
      ToolTipText     =   "Clic para cancelar acción y cerrar la ventana"
      Top             =   2610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      Icon            =   "FrmLogin.frx":38FC6
      Style           =   5
      Caption         =   "    &Cancelar"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   "Clic para cancelar acción y cerrar la ventana"
      ToolTipTitle    =   "Cancelar"
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   600
      Picture         =   "FrmLogin.frx":39560
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Tag             =   "Copyright Release 7.7.0"
      Top             =   2730
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   195
      Left            =   2580
      TabIndex        =   5
      Tag             =   "Company"
      Top             =   2190
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   2580
      TabIndex        =   4
      Tag             =   "Password"
      Top             =   1740
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   2580
      TabIndex        =   1
      Tag             =   "User"
      Top             =   1320
      Width           =   540
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Public LoginSucceeded As Boolean

Public bOk            As Boolean

Private Function Valida_usuario(ByVal xusr As Variant, _
                                ByVal xpas As Variant, _
                                ByVal xemp As Variant)

    'conn = "Provider=sqloledb;Server=HIALPESA4;Database=HIALPESA;Integrated Security=SSPI;"
    'conn = "Provider=sqloledb;Server=HIALPESA4;Database=SEGURIDAD;Integrated Security=SSPI;"

    If Len(xemp) = 0 Then
        sQuery = "SELECT count(*) from seg_usuarios  WHERE COD_USUARIO='" & xusr & "'"
    Else
        sQuery = "SELECT cod_perfil from seg_empusuper a,seg_usuarios b WHERE  a.cod_empresa='" & xemp & "' and a.cod_usuario=b.cod_usuario and b.COD_USUARIO='" & xusr & "'"
    End If

    If xpas = "" Then
        xcondi = " and (password is null or password='')"
    Else
        xcondi = "  and password='" & xpas & "'"
    End If

    sQuery = sQuery & xcondi
    Set RS1 = New ADODB.Recordset
    RS1.ActiveConnection = conn
    'RS1.ActiveConnection = sconnect
    RS1.CursorType = adOpenStatic
    RS1.Open sQuery

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
    sQuery = "SELECT A.COD_EMPRESA AS CODIGO,B.DES_EMPRESA AS NOMBRE,RUTA_LOGO,NUM_RUC,DIRECCION,DSN,RUTA0 FROM SEG_EMPUSUPER A,SEG_EMPRESAS B WHERE A.COD_EMPRESA=B.COD_EMPRESA AND A.COD_USUARIO='" & TxtUserName & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open sQuery
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

    sQuery = "SELECT ISNULL(RUTA_LOGO,'') AS RUTA_LOGO,ISNULL(NUM_RUC,'') AS NUM_RUC,ISNULL(DIRECCION,'') AS DIRECCION,ISNULL(DSN,'') AS DSN,ISNULL(RUTA0,'') AS RUTA0 , ISNULL(DSNSEGURIDAD,'') AS DSNSEGURIDAD FROM SEG_EMPRESAS  WHERE COD_EMPRESA='" & codemp & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open sQuery
    iCount = mRs.RecordCount

    If iCount > 0 Then
        Ruta_Logo_Empresa = mRs(0)
        Num_Ruc_Empresa = mRs(1)
        Direccion_Empresa = mRs(2)
''''        DSN_Empresa = mRs(3)
        'Para pruebas
        'DSN_Empresa = "Provider=sqloledb;Server=" & GetSetting("Visuales", "Settings", "Server") & "Database=LIVES;UID=sa;pwd=;"
        'DSN_Empresa = "Provider=sqloledb;Server=ECARDENAS\ECN;Database=HIALPESAX;Integrated Security=SSPI;"
        'DSN_Empresa = "Provider=sqloledb;Server=VRIOS2;Database=HIALPESAX;Integrated Security=SSPI;"
        'DSN_Empresa = "Provider=sqloledb;Server=BESCALANTE\SQL2005;Database=HIALPESA;Integrated Security=SSPI;"
     
        'DSN_Empresa = "Provider=sqloledb;Server=wflores2;Database=inka;Integrated Security=SSPI;"
     
        '''''DSN_Empresa = "Provider=sqloledb;Server=CESARATOCHE2\SQL2005;Database=INKADESIGNS;Integrated Security=SSPI;"
     
        '''''DSN_Empresa = "Provider=sqloledb;Server=HIALPESA4;Database=HIALPESA;Integrated Security=SSPI;"
     
        '''''DSN_Empresa = "Provider=sqloledb;Server=CATOCHE\SQL2008R2;Database=HIALPESA;Integrated Security=SSPI;"
     
        '''''''''''DSN_Empresa = "Provider=sqloledb;Server=CATOCHEPC\SQL2008R2;Database=HIALPESA_27052014;Integrated Security=SSPI;"
     
        ''''''''''''''''DSN_Empresa = "Provider=sqloledb;Server=SERVER_DEV;Database=HIALPESA;Integrated Security=SSPI;"
             
        Ruta0_Empresa = mRs(4)
''''        DSN_Seguridad = mRs(5)
''        cCONNECT = DSN_Empresa

        cCONNECT = G_CONEXION_SQL
        DSN_Seguridad = G_CONEXION_SQL_SEG
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

   ''' sconnect = "Provider=sqloledb;Server=SERVER_DEV;Database=SEGURIDAD;Integrated Security=SSPI;"
    'conn = "Provider=sqloledb;Server=SERVER_DEV;Database=SEGURIDAD;Integrated Security=SSPI;"
    IdiomaEtiquetas1 Me
    Label4.Caption = "Copyright Release " & App.Major & "." & App.Minor & "." & App.Revision

    'If GetUserNameEx(NameSamCompatible, sBuffer, Ret) <> 0 Then
    'TxtUserName = ComputerName

    ' Ruta0_Empresa = mRs(4)
    ' DSN_Seguridad = mRs(5)
    ' cCONNECT = DSN_Empresa
    strSql = "EXEC SEG_USUARIO_ESTACION_SELECT " & vbNewLine
    strSql = strSql & "@Cod_Estacion='" & ComputerName & "'" & vbNewLine

    Dim rsPc As New ADODB.Recordset

    ' MsgBox conn.ConnectionString
    '  Set rsPc = CargarRecordSetDesconectado(strSQL, conn.ConnectionString)
    ' MsgBox "XXX"
           
    rsPc.ActiveConnection = conn
    'RS1.ActiveConnection = sconnect
    rsPc.CursorType = adOpenStatic
    rsPc.Open strSql

    If Not (rsPc.BOF And rsPc.EOF) Then
        rsPc.MoveFirst
        TxtUserName = rsPc.Fields("COD_USUARIO").value
    Else
        TxtUserName = ComputerName
    End If

    ' Mid(sBuffer, InStr(sBuffer, "\") + 1, 20)
    'End If
End Sub

Private Sub Timer1_Timer()

    Static estado

    If estado = Empty Then
        Image1.Visible = True
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        estado = 2
    ElseIf estado = 2 Then
        Image1.Visible = False
        Image2.Visible = True
        Image3.Visible = False
        Image4.Visible = False
        estado = 3

    ElseIf estado = 3 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = True
        Image4.Visible = False
        estado = 4

    ElseIf estado = 4 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = True
        estado = Empty
    End If

End Sub

Private Sub txtPassword_GotFocus()
    Call focoControl(txtPassword, True)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub txtPassword_LostFocus()
Call focoControl(txtPassword, False)
End Sub

Private Sub TxtUserName_GotFocus()
    Call focoControl(TxtUserName, True)
    SelectionText TxtUserName
End Sub

Private Sub TxtUserName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        'cmdOK_Click
    End If

End Sub

Private Sub TxtUserName_LostFocus()
Call focoControl(TxtUserName, False)
End Sub
