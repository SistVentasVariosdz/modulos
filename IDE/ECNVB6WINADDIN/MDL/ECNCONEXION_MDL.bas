Attribute VB_Name = "mdl003_ECNCONEXION"
Option Explicit

Global GO_003_RUTA_ECN_APP As String
Global GO_003_FORM_PARENT As Object
Global GO_003_PU_SW_LOAD As Boolean
Global GO_003_MODO_WIN As G_ENU_ModoWIN_CONEXION
Global GO_003_ENU_OPC_WIN_RESULT As GE_WIN_RESULT

Public PU_003_ECNLIB01_FUNSUB As ECNVB6LIB.ECNLIB01_FUNSUB
Public PU_003_ECNLIB03_WINEVE As ECNVB6LIB.ECNLIB03_WINEVE
Public PU_003_ECNLIB04_EFFECTS As ECNVB6LIB.ECNLIB04_EFFECTS

Public Const PU_003_Ks_ECNOVL_NOMFILE As String = "ECNOVL.com"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_SERVIDOR As String = "SERVER"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_USER As String = "[SQL_USER]"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_UPWD As String = "[SQL_UPWD]"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_MAIN As String = "[BDN_MAIN]"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_IMAG As String = "[BDN_IMAG]"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_CONE As String = "[SQL_CONE]"
Public Const PU_003_Ks_ECNOVL_LINEA_ECNOVL_XLS_CONE As String = "[XLS_SERV]"

Public Const PU_003_Ks_ICO_CONNECT_YES As Integer = 1
Public Const PU_003_Ks_ICO_CONNECT_NOT As Integer = 2


Public PU_003_RUTA_ECN_OVL As String

Global GO_003_SQL_USER As String
Global GO_003_SQL_UPWD As String
Global GO_003_BD_MAIN As String
Global GO_003_BD_IMAG As String
Global GO_003_CONEXION_SQL As String
Global GO_003_CONEXION_XLS As String

Public PU_003_SERVER As String
Public PU_003_USUARIO As String
Public PU_003_CLAVE As String
Public PU_003_BD As String
Public PU_003_BDI As String

Public PU_003_CONNECT As String
Public PU_003_SW_Conexion As Boolean

Public Sub PU_003_CrearECNOVL(ByVal sSQLServ As String, _
                       ByVal sSQLUser As String, _
                       ByVal sSQLUpwd As String, _
                       ByVal sBDNMain As String, _
                       ByVal sBDNImag As String)
    On Error GoTo SALTO_ERROR
        
    If Len(Trim(GO_003_RUTA_ECN_APP)) = 0 Then
        MsgBox "No se ha definido la ruta de la aplicacion en la clase, la clave [PU_003_RUTA_ECN_OVL]", vbCritical, "mdl003_ECNCONEXION - PU_003_CrearECNOVL"
        Exit Sub
    End If
    
    
    Dim sCadena As String
    Dim sSQLCadCon As String
    Dim sXLSCadCon As String
        
    sSQLCadCon = "Provider=SQLOLEDB;" _
               & "Data Source=" & sSQLServ & ";" _
               & "Initial Catalog=" & sBDNMain & ";" _
               & "Uid=" & sSQLUser & ";" _
               & "Pwd=" & sSQLUpwd & ";"
               
    sXLSCadCon = "ODBC;" _
               & "DRIVER=SQL Server;" _
               & "SERVER=" & sSQLServ & ";" _
               & "UID=" & sSQLUser & ";" _
               & "PWD=" & sSQLUpwd & ";" _
               & "DATABASE=" & sBDNMain
    
    Open PU_003_RUTA_ECN_OVL For Output As #1
    
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(PU_003_Ks_ECNOVL_LINEA_ECNOVL_SERVIDOR))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_USER & "=" & sSQLUser
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_UPWD & "=" & sSQLUpwd
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_MAIN & "=" & sBDNMain
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_IMAG & "=" & sBDNImag
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_CONE & "=" & sSQLCadCon
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_XLS_CONE & "=" & sXLSCadCon
    Print #1, PU_003_ECNLIB01_FUNSUB.Cript2(Trim(sCadena))
    
    Close #1
    Exit Sub
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, "mdl003_ECNCONEXION - CrearECNOVL"
End Sub

Public Function PU_003_LeerECNOVL() As Boolean
    PU_003_LeerECNOVL = True
    If Len(Trim(GO_003_RUTA_ECN_APP)) = 0 Then
        MsgBox "No se ha definido la ruta de la aplicacion en la clase, la clave [PU_003_RUTA_ECN_OVL]", vbCritical, "mdl003_ECNCONEXION - PU_003_CrearECNOVL"
        PU_003_LeerECNOVL = False
        Exit Function
    End If
    
    On Error GoTo SALTO_ERROR
    
    Dim sCadena As String
    Dim sLinea As String
    Dim sTipo  As String
    Dim N      As Long
    Dim lPosCadFnd    As Long
        
    N = FreeFile
    GO_003_SQL_USER = ""
    Open PU_003_RUTA_ECN_OVL For Input As N
    Do While Not EOF(N)
        Line Input #N, sLinea
        sCadena = PU_003_ECNLIB01_FUNSUB.DCript2(sLinea)
        If sCadena = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SERVIDOR Then
            sTipo = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SERVIDOR
        End If
        lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_USER): If lPosCadFnd > 0 Then GO_003_SQL_USER = Mid(sCadena, 12)
        lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_UPWD): If lPosCadFnd > 0 Then GO_003_SQL_UPWD = Mid(sCadena, 12)
        lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_MAIN): If lPosCadFnd > 0 Then GO_003_BD_MAIN = Mid(sCadena, 12)
        lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_BDN_IMAG): If lPosCadFnd > 0 Then GO_003_BD_IMAG = Mid(sCadena, 12)
        If sTipo = PU_003_Ks_ECNOVL_LINEA_ECNOVL_SERVIDOR Then
            lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_SQL_CONE): If lPosCadFnd > 0 Then GO_003_CONEXION_SQL = Mid(sCadena, 12)
            lPosCadFnd = InStr(sCadena, PU_003_Ks_ECNOVL_LINEA_ECNOVL_XLS_CONE): If lPosCadFnd > 0 Then GO_003_CONEXION_XLS = Mid(sCadena, 12)
        End If
    Loop
    Close N
    Exit Function
    
SALTO_ERROR:
    PU_003_LeerECNOVL = False
    If Err.Number = 53 Then
        MsgBox "No se podrá conectar a la BD de SQL", vbCritical, "mdl003_ECNCONEXION - LeerECNOVL"
        Unload frm003_ECNCONEXION
    Else
        Resume Next
    End If
End Function


Public Sub PU_003_ADO()
    On Error GoTo Err_Conexion:
    
    Screen.MousePointer = vbHourglass
    
    With frm003_ECNCONEXION
        PU_003_SERVER = Trim(.txtServidor)
        PU_003_USUARIO = Trim(.txtUsuario)
        PU_003_CLAVE = Trim(.txtClave)
    End With
    
    If PU_003_SERVER = Empty _
    Or PU_003_USUARIO = Empty Then
        Screen.MousePointer = vbCustom
        PU_003_SW_Conexion = False
        'frm003_ECNCONEXION.lblProceso.Caption = Empty
        Exit Sub
    End If
    
    Dim oCn As ADODB.Connection
    
    Set oCn = New ADODB.Connection
    PU_003_CONNECT = "Provider=SQLOLEDB.1;" _
                   & "Password=" & PU_003_CLAVE & ";" _
                   & "Persist Security Info=True;" _
                   & "User ID=" & PU_003_USUARIO & ";" _
                   & "Data Source=" & PU_003_SERVER
    oCn.Open PU_003_CONNECT
    oCn.Close
    
    Dim bvData() As Byte
    With frm003_ECNCONEXION
        bvData = .imlConexion.GetStream(PU_003_Ks_ICO_CONNECT_YES)
        
        .lblProceso.Caption = "Conexion establecida...!!"
        .icnConnect.LoadImageFromStream bvData
    End With
    Set oCn = Nothing
    PU_003_SW_Conexion = True
    Screen.MousePointer = vbCustom
    Exit Sub
    
Err_Conexion:
    Screen.MousePointer = vbCustom
    PU_003_SW_Conexion = False
    frm003_ECNCONEXION.lblProceso.Caption = Empty
    If GO_003_PU_SW_LOAD = False Then MsgBox "No se encontro el servidor", vbInformation, "mdl003_ECNCONEXION - PU_003_ADO"
    If Not oCn Is Nothing Then
        If oCn.State = 1 Then
            oCn.Close
        End If
        Set oCn = Nothing
    End If
    Err.Clear
End Sub

Public Sub PU_003_GetBD()
    If PU_003_SW_Conexion = False Then Exit Sub
    Dim oECNSQLHELP As New ECNVB6LIB.ECNSQLHELP
    Dim oRs As New ADODB.Recordset
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT CODIGO = DBID,"
    xSQL = xSQL & vbNewLine & "       NOMBRE = NAME"
    xSQL = xSQL & vbNewLine & "FROM MASTER.DBO.SYSDATABASES"
    xSQL = xSQL & vbNewLine & "ORDER BY NAME"
    
    oECNSQLHELP.CADENA_CONEXION = PU_003_CONNECT
    Set oRs = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, False)
    If oRs.BOF And oRs.EOF Then
        MsgBox "No hay base de datos a seleccionar", vbInformation, "mdl003_ECNCONEXION - PU_003_GetBD"
    Else
        With frm003_ECNCONEXION
            With .dcBDI
                Set .RowSource = oRs
                .ListField = "NOMBRE"
                .BoundColumn = "NOMBRE" '"CODIGO"
            End With
            With .dcBD
                Set .RowSource = oRs
                .ListField = "NOMBRE"
                .BoundColumn = "NOMBRE" '"CODIGO"
                If .Visible = True Then
                    If GO_003_PU_SW_LOAD = False Then .SetFocus
                End If
            End With
        End With
    End If
    Set oRs = Nothing
End Sub

Public Sub PU_003_Click_End_dcBD(Optional ByVal Area As Integer = 2)
    If Area <> 2 Then Exit Sub
    PU_003_BD = frm003_ECNCONEXION.dcBD.Text
    If Len(Trim(PU_003_BD)) > 0 Then
        frm003_ECNCONEXION.btnConectar.Caption = "&Aceptar"
        PU_003_CONNECT = "Provider=SQLOLEDB.1;Password=" & PU_003_CLAVE & ";Persist Security Info=True;User ID=" & PU_003_USUARIO & ";Initial Catalog=" & PU_003_BD & ";Data Source=" & PU_003_SERVER
    End If
End Sub
