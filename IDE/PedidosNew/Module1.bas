Attribute VB_Name = "Module1"
Declare Function GetcomputerName _
        Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                  nSize As Long) As Long
'Declare Function GetUserNameEx Lib "secur32.dll" Alias _
'"GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, _
'ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

Global sconnect As String

Global conn     As New ADODB.Connection

Enum EXTENDED_NAME_FORMAT

    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10

End Enum

Public Function ConectarBD(usuario, password, servidor, base) As Integer
    ConectarBD = 0

    On Error GoTo procesaerror

    sconnect = "Provider=MSDASQL.1;Password= " & password & ";User ID = " & usuario & ";Data Source= " & servidor & ";Initial Catalog=" & base
    'SConnect = "Provider=MSDASQL.1;Persist Security Info=False;Password=;User ID=sa;Data Source=servidor;Initial Catalog=seguridad"
    'Set conn = CreateObject("ADODB.Connection")
    conn.Open sconnect
     
    ConectarBD = 1

    Exit Function

procesaerror:
    ConectarBD = 0
End Function

Public Sub Main()

    Dim Network            As WshNetwork
    Dim sModoAutenticacion As String
    Dim sServerName        As String

    Set Network = New WshNetwork

    'sModoAutenticacion = GetSetting("Visuales", "Settings", "AutenticacionMode")
    'sServerName = GetSetting("Visuales", "Settings", "Server")


'    xservidor = "LocalServer"
'    xbase = "SEGURIDAD"
'    xusr = "sa"
'    xpas = ""

     '*********************************************************************************************************
    ' ME CONECTO A LA BD
    '*********************************************************************************************************
RETORNO:
    Dim oECNCONEXION As New cls003_ECNCONEXION
    oECNCONEXION.WINAPPPATH = App.Path
    If Len(Dir(oECNCONEXION.RUTA_ECN_OVL)) <> 0 Then
        Select Case oECNCONEXION.LeerECNOVL
            Case False
                End
            Case True
                With oECNCONEXION
                    G_SQL_USER = .ECNOVL_SQL_USER
                    G_SQL_UPWD = .ECNOVL_SQL_UPWD
                    G_BD_MAIN = .ECNOVL_BD_MAIN
                    G_BD_IMAG = .ECNOVL_BD_IMAG
                    G_BD_SEG = .ECNOVL_BD_SEG
                    G_CONEXION_SQL = .ECNOVL_CONEXION_SQL
                    G_CONEXION_XLS = .ECNOVL_CONEXION_XLS
                    G_CONEXION_SQL_SEG = .ECNOVL_CONEXION_SQL_SEG
                End With
        End Select
    Else
        Dim blSW_Accept As Boolean
        
        With oECNCONEXION
            .WINAPPPATH = App.Path
            .MODO_WIN = WINC_Main
            .ShowPrompt
            blSW_Accept = True
            If .WIN_RESULT <> WD_ACCEPT Then _
                blSW_Accept = False
        End With
        Set oECNCONEXION = Nothing
        If blSW_Accept = False Then End
        GoTo RETORNO
    End If
    '*********************************************************************************************************
   If ConectarBD1(G_CONEXION_SQL_SEG) Then
   
        cCONNECT = G_CONEXION_SQL
        cSEGURIDAD = G_CONEXION_SQL_SEG
        DSN_Empresa = G_CONEXION_SQL
        DSN_Seguridad = G_CONEXION_SQL_SEG
  '  If ConectarBD1(G_SQL_USER, G_SQL_UPWD, xservidor, G_BD_MAIN, G_CONEXION_SQL) Then
      '  Load FrmLogin
        FrmLogin.Show
    Else
        Call MsgBox("No se ha podido realizar la conexiòn", vbInformation, MDIPrincipal.Caption)
    End If

End Sub

Private Sub SeteaReg()
    DeleteSetting "Visuales", "Settings", "Server"
    SaveSetting "Visuales", "Settings", "Server", "HIALPESA4"
    'MsgBox "Nuevo Servidor SQL HIALPESA1 Instalado", vbInformation
End Sub

'Public Function ConectarBD1(usuario, password, servidor, base) As Integer
Public Function ConectarBD1(strConexion As String) As Integer
    ConectarBD1 = 0

    On Error GoTo procesaerror

    Dim sServer            As String

    Dim sModoAutenticacion As String

    Dim sUserName          As String

    Dim sPassword          As String
     
    sServer = GetSetting("Visuales", "Settings", "Server")
    sModoAutenticacion = GetSetting("Visuales", "Settings", "AutenticacionMode")

    If UCase(sModoAutenticacion) = "SQL" Then
        sUserName = GetSetting("Visuales", "Settings", "UserName")
        sPassword = GetSetting("Visuales", "Settings", "Password")
        sconnect = strConexion
    '    sconnect = "Provider=SQLOLEDB;User ID=" & RTrim(sUserName) & ";Password=" & RTrim(sPassword) & ";Server=" & sServer & ";Database=SEGURIDAD;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
    Else
        sconnect = strConexion
  
        '  MsgBox "Z"
   '     sconnect = "Provider=SQLOLEDB;Integrated Security=SSPI;Server=" & sServer & ";Database=SEGURIDAD;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
    
    End If
     
    conn.ConnectionString = sconnect
    
    conn.Open sconnect

    ConectarBD1 = 1

    Exit Function

procesaerror:
    ConectarBD1 = 0
End Function

Public Function get_botones(ByVal f As Form, _
                            ByVal Vcod_perfil As Variant, _
                            ByVal vcod_empresa As Variant, _
                            ByVal fname As Variant)
    Set RS1 = New ADODB.Recordset
    sQuery = "Sp_funciones3 '" & Vcod_perfil & "','" & vcod_empresa & "','" & fname & "'"
    RS1.ActiveConnection = conn
    RS1.CursorType = adOpenStatic
    RS1.Open sQuery

    If Not (RS1.BOF And RS1.EOF) Then

        For j = 1 To RS1.RecordCount
            Boton_Enabled RS1!nom_corto, f
            RS1.MoveNext
        Next j

    End If

End Function

Private Sub Boton_Enabled(ByVal sName As Variant, ByVal f As Form)

    Dim ctl As Control

    For Each ctl In f.Controls

        If TypeOf ctl Is Button Then
            If LTrim(RTrim(UCase(sName))) = LTrim(RTrim(UCase(ctl.Name))) Then
                ctl.Enabled = True

                Exit For

            End If
        End If

    Next ctl

End Sub

Public Sub DActivaControles(ByVal fform As Form, _
                            ByVal TipOpe As Variant, _
                            ByVal Scontroles As String)

    If TipOpe = "A" Then
        xEnabled = True
        xbackColor = &H80000005
    Else
        xEnabled = False
        xbackColor = &H8000000B
    End If

    Dim ctl As Control

    For Each ctl In fform.Controls

        If InStr(UCase(Scontroles), UCase(ctl.Name)) > 0 Then
            If InStr("V/I", TipOpe) > 0 Then
                If TipOpe = "V" Then
                    ctl.Visible = True
                Else
                    ctl.Visible = False
                End If

            Else
                ctl.Enabled = xEnabled

                If UCase(Mid(ctl.Name, 1, 3)) <> "CMD" Then
                    ctl.BackColor = xbackColor
                End If
            End If
        End If

    Next ctl

End Sub

Public Sub Limpia_Campos(ByVal fform As Form, ByVal Scontroles As String)

    Dim ctl As Control

    For Each ctl In fform.Controls

        If InStr(UCase(Scontroles), UCase(ctl.Name)) > 0 Then
            ctl.Text = ""
        End If

    Next ctl

End Sub

Public Function Maximo(ByVal stabla As String, _
                       ByVal sCampo As String, _
                       ByVal scondi As String, _
                       ByVal conn As ADODB.Connection, _
                       ByVal stipo As String, _
                       ByVal ilargo As Integer)
    Set RS1 = New ADODB.Recordset
    RS1.ActiveConnection = cn
    RS1.CursorType = adOpenStatic

    If scondi = "" Then
        scondi = "1<2"
    End If

    If stipo = "S" Then
        sQuery = "select len(" & sCampo & ")" & ",max(" & sCampo & ") from " & stabla & " where " & scondi & " group by len(" & sCampo & ")"
        RS1.Open sQuery

        If Not RS1.EOF Then
            a = RS1(1) + 1
            b = RS1(0)
            a = Ceros(a, b, "0")
        Else
            a = Ceros("1", ilargo, "0")
        End If

    Else
        sQuery = "select max(" & sCampo & ") from " & stabla & " where " & scondi
        RS1.Open sQuery
        a = RS1(1)

        If IsNull(a) Then
            a = 1
        End If
    End If

    Maximo = a
    Set RS1 = Nothing
End Function

Public Function Ceros(ByVal scadena As String, _
                      ByVal iLen As Integer, _
                      ByVal schar As String)
    Ceros = scadena

    If iLen < 2 Then Exit Function

    For i = 1 To iLen - 1
        Ceros = schar & Ceros
    Next i

End Function

Public Sub Carga_Categorias(ByVal fform As Form, _
                            ByVal Datag As Object, _
                            ByRef rs As ADODB.Recordset)
    sQuery = "SELECT COD_MOTATR AS CODIGO,DES_MOTATR AS DESCRIPCION FROM TG_MOTATR"
    'Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorType = adOpenStatic
    rs.Open sQuery

    With fform
        Set Datag.DataSource = rs
        Set .txtIdCategoria.DataSource = rs
        .txtIdCategoria.DataField = "CODIGO"
        Set .txtNombre.DataSource = rs
        .txtNombre.DataField = "DESCRIPCION"
    End With

    'Set rs = Nothing
End Sub

Public Sub ReCarga_Categoria(ByVal fform As Form, _
                             ByVal Datag As Object, _
                             ByRef rs As ADODB.Recordset)

    With fform
        Set Datag.DataSource = rs
        Set .txtIdCategoria.DataSource = rs
        .txtIdCategoria.DataField = "CODIGO"
        Set .txtNombre.DataSource = rs
        .txtNombre.DataField = "DESCRIPCION"
    End With

    'Set rs = Nothing
End Sub

Public Function Borra_Categoria(ByVal f As Form, ByVal cn As ADODB.Connection)
    Borra_Categoria = False

    Dim RS1 As New ADODB.Recordset

    RS1.ActiveConnection = conn

    With f
        sQuery1 = "select count(*) from articulos where codcat = '" & .txtIdCategoria & "'"
        sQuery = " Delete from  Categorias where codcat = '" & .txtIdCategoria & "'"
        RS1.Open sQuery1
        a = LTrim(RTrim(RS1(0)))

        If Not IsNull(a) Then
            If a = 0 Then
                cn.Execute sQuery
                Borra_Categoria = True
            End If

        Else
            cn.Execute sQuery
            Borra_Categoria = True
        End If

    End With

    Set RS1 = Nothing
End Function

Public Function Valida_Inserta_Categorias(ByVal f As Form)
    Valida_Inserta_Categorias = True

    With f

        If Len(.txtNombre) = 0 Then
            xVAr = " Nombre Categoria "
            MsgBox ("El Campo " & xVAr & " està Vacìo")
            Valida_Inserta_Categorias = False
        End If

    End With

End Function

Public Function Graba_Categoria(ByVal f As Form, _
                                ByVal cn As ADODB.Connection, _
                                ByVal stipomov As String)

    'On Error GoTo Trata_error
    'cn.BeginTrans
    With f

        If stipomov = "I" Then
            sQuery = " Insert into TG_MOTATR (" & " " & "COD_MOTATR,DES_MOTATR)  VALUES (" & " " & .txtIdCategoria & "," & "'" & .txtNombre & "')"
            cn.Execute sQuery
        Else
            sQuery = " Update TG_MOTATR set  " & " " & "DES_MOTATR ='" & .txtNombre & "' " & " where " & "COD_MOTATR='" & .txtIdCategoria & "'"
            '" " & "CODCAT='" & .dbcboCategoria.BoundText
            cn.Execute sQuery
        End If

        'cn.CommitTrans
    End With

    'Trata_error:
    ' cn.RollbackTrans
    ' MsgBox ("No se grabaron los cambios")
End Function

Public Function Borra_Motatr(ByVal f As Form, ByVal cn As ADODB.Connection)

    With f
        sQuery = " Delete from  TG_Motatr where cod_motatr = '" & .txtIdArticulo & "'"
    End With

    conn.Execute sQuery
End Function

Sub Avanza(ByVal Tecla As Integer)

    Select Case Tecla

        Case 13, 40: SendKeys "{TAB}", True

        Case 38: SendKeys "+{TAB}", True
    End Select

End Sub

Public Function Ancho_Columnas(ByVal fform As Form, _
                               ByVal dcontainer As Object, _
                               ByVal scadena As String)
    'identificador de fin de cadena
    scadena = scadena & "_"
    xPos = 1
    xPos1 = 1
    i = 0

    If TypeOf dcontainer Is MSHFlexGrid Then
        xcol = dcontainer.Cols
    Else
        xcol = dcontainer.Columns.count
    End If

    Dim a As Integer

    While InStr(xPos1, scadena, ",") > 0 And i <= xcol

        xPos1 = InStr(xPos, scadena, ",") + 1

        If TypeOf dcontainer Is MSHFlexGrid Then
            dcontainer.ColWidth(i) = (CInt(Mid(scadena, xPos, xPos1 - xPos - 1)) * 100) + 0
        Else
            dcontainer.Columns(i).Width = (CInt(Mid(scadena, xPos, xPos1 - xPos - 1)) * 100) + 0
        End If
   
        xPos = xPos1
        i = i + 1

    Wend

    'ultimo campo
    If xcol > i Then
        If TypeOf dcontainer Is MSHFlexGrid Then
            dcontainer.ColWidth(i) = (CInt(Mid(scadena, xPos1, CInt(InStr(scadena, "_")) - xPos1)) * 100) + 0
        Else
            dcontainer.Columns(i).Width = (CInt(Mid(scadena, xPos1, CInt(InStr(scadena, "_")) - xPos1)) * 100) + 0
        End If
    End If

End Function

