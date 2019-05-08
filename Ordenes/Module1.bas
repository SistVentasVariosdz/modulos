Attribute VB_Name = "Module1"
Declare Function GetcomputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
'Declare Function GetUserNameEx Lib "secur32.dll" Alias _
'"GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, _
'ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

Global sconnect As String
Global conn As New ADODB.Connection

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
    Dim Network As WshNetwork
    Dim sModoAutenticacion As String
    Dim sServerName As String
    
    If App.PrevInstance = True Then
        Call MsgBox("No se puede ejecutar la aplicacion mas de una vez", vbInformation, "") ',Me.Caption)
        End
    End If
    
    Set Network = New WshNetwork
    
If UCase(Network.UserDomain) = "textilesjoc.com" Then
    SeteaReg
End If
    
If sServerName = "" Then
 sServerName = ComputerName
End If

sModoAutenticacion = GetSetting("Visuales", "Settings", "AutenticacionMode")
sServerName = GetSetting("Visuales", "Settings", "Server")
    
If sServerName = "" Then
 sServerName = ComputerName
End If
    
'    If UCase(sModoAutenticacion) = "SQL" Then
'        If UCase(sServerName) <> "SERVIDOR" And (sServerName) <> "SERVERDYC" And UCase(sServerName) <> "PROBOOKHP" Then
'            MsgBox "Aplicación no Registrada (2)!!!"
'    '''ERU :        End
'        End If
'    Else
'        If UCase(Network.UserDomain) <> "HIALPESA" And _
'           UCase(Network.UserDomain) <> "LIVES" And _
'           UCase(Network.UserDomain) <> "SUMIT" And _
'           UCase(Network.UserDomain) <> "ONLY_STAR" And _
'           RTrim(UCase(Network.UserDomain)) <> "PRECOTEX" And _
'           RTrim(UCase(Network.UserDomain)) <> "" And _
'           RTrim(UCase(Network.UserDomain)) <> "GENESYSDATA" And _
'           RTrim(UCase(Network.UserDomain)) <> "GENESISDATA" And _
'           RTrim(UCase(Network.UserDomain)) <> "GRUPO_TRABAJO" And _
'           RTrim(UCase(Network.UserDomain)) <> "SBAPERU" Then
'            MsgBox "Aplicación no Registrada (1)!!!"
'        End If
'    End If
'
    
     xservidor = "PC_ALMACEN" 'sServerName
     xbase = "SEGURIDAD"
     xusr = "soporte"
     xpas = "soporte"
     If ConectarBD1(xusr, xpas, xservidor, xbase) Then
        FrmLogin.Show
     Else
        MsgBox ("No se ha podido realizar la conexiòn"), vbInformation
     End If
End Sub

Private Sub SeteaReg()
    DeleteSetting "Visuales", "Settings", "SERVER"
    SaveSetting "Visuales", "Settings", "SERVER", "SERVERDATA"
End Sub


Public Function ConectarBD1(usuario, password, servidor, base) As Integer
    ConectarBD1 = 0
    On Error GoTo procesaerror
    Dim sServer As String
    Dim sModoAutenticacion  As String
    Dim sUserName As String
    Dim sPassword As String

    sServer = GetSetting("Visuales", "Settings", "Server")
    sModoAutenticacion = GetSetting("Visuales", "Settings", "AutenticacionMode")
    If UCase(sModoAutenticacion) = "SQL" Then
       sUserName = GetSetting("Visuales", "Settings", "UserName")
       sPassword = GetSetting("Visuales", "Settings", "Password")
       sconnect = "Provider=SQLOLEDB;User ID=" & RTrim(sUserName) & ";Password=" & RTrim(sPassword) & ";Server=" & sServer & ";Database=SEGURIDAD;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
    Else
       'sconnect = "Provider=SQLOLEDB;Integrated Security=SSPI;Server=" & sServer & ";Database=SEGURIDAD;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
       sconnect = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & usuario & ";Initial Catalog=" & base & ";Data Source=" & servidor & ""
    End If
    conn.ConnectionString = sconnect
    conn.Open sconnect
    ConectarBD1 = 1
Exit Function

procesaerror:
  ConectarBD1 = 0
End Function


Public Function get_botones(ByVal f As Form, ByVal Vcod_perfil As Variant, ByVal vcod_empresa As Variant, ByVal fname As Variant)
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

Private Sub Boton_Enabled(ByVal sname As Variant, ByVal f As Form)
Dim ctl As Control
For Each ctl In f.Controls
        If TypeOf ctl Is Button Then
          If LTrim(RTrim(UCase(sname))) = LTrim(RTrim(UCase(ctl.Name))) Then
              ctl.Enabled = True
              Exit For
          End If
        End If
  Next ctl
End Sub



Public Sub DActivaControles(ByVal fform As Form, ByVal TipOpe As Variant, ByVal Scontroles As String)
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
ByVal sCampo As String, ByVal scondi As String, ByVal conn As ADODB.Connection, ByVal stipo As String, ByVal ilargo As Integer)
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
    A = RS1(1) + 1
    B = RS1(0)
    A = Ceros(A, B, "0")
    Else
    A = Ceros("1", ilargo, "0")
    End If
Else
    sQuery = "select max(" & sCampo & ") from " & stabla & " where " & scondi
    RS1.Open sQuery
    A = RS1(1)
    If IsNull(A) Then
    A = 1
    End If
End If
Maximo = A
Set RS1 = Nothing
End Function
Public Function Ceros(ByVal scadena As String, ByVal iLen As Integer, ByVal schar As String)
Ceros = scadena
If iLen < 2 Then Exit Function
For i = 1 To iLen - 1
Ceros = schar & Ceros
Next i
End Function


Public Sub Carga_Categorias(ByVal fform As Form, ByVal Datag As Object, ByRef rs As ADODB.Recordset)
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
Public Sub ReCarga_Categoria(ByVal fform As Form, ByVal Datag As Object, ByRef rs As ADODB.Recordset)
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
    A = LTrim(RTrim(RS1(0)))
    If Not IsNull(A) Then
      If A = 0 Then
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

Public Function Graba_Categoria(ByVal f As Form, ByVal cn As ADODB.Connection, ByVal stipomov As String)
'On Error GoTo Trata_error
'cn.BeginTrans
With f
If stipomov = "I" Then
    sQuery = " Insert into TG_MOTATR (" & _
    " " & "COD_MOTATR,DES_MOTATR)  VALUES (" & _
    " " & .txtIdCategoria & "," & _
    "'" & .txtNombre & "')"
    cn.Execute sQuery
 Else
    sQuery = " Update TG_MOTATR set  " & _
    " " & "DES_MOTATR ='" & .txtNombre & "' " & _
    " where " & "COD_MOTATR='" & .txtIdCategoria & "'"
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

Public Function Ancho_Columnas(ByVal fform As Form, ByVal dcontainer As Object, ByVal scadena As String)
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
Dim A As Integer
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



