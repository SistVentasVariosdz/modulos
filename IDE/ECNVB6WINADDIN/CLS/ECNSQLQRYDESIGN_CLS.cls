VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls002_ECNSQLQRYDESIGN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum GE_TIPO_TABLA_CLAVE
    TC_0NU = 0 '=> NULO
    TC_1UK = 1
    TC_2FK = 2
    TC_3PK = 3
End Enum

Public Enum GE_TIPO_TABLA
    TT_TABLA = 1
    TT_VISTA = 2
End Enum

Private Sub Class_Initialize()
    Call PU_002_AperturarRSdeTablas
End Sub

Public Property Get CADENA_CONEXION() As String
    CADENA_CONEXION = GO_002_CONEXION_SQL
End Property

Public Property Let CADENA_CONEXION(ByVal xCadenaConexion As String)
    GO_002_CONEXION_SQL = xCadenaConexion
End Property

Public Property Get RUTA_INI_PARAM_WIN() As String
    RUTA_INI_PARAM_WIN = GO_002_RUTA_INI_PARAM_WIN
End Property

Public Property Let RUTA_INI_PARAM_WIN(ByVal xRutaIni As String)
    GO_002_RUTA_INI_PARAM_WIN = xRutaIni
End Property

Public Property Get SW_LOAD_DESIGN() As Boolean
    SW_LOAD_DESIGN = GO_002_SW_LOAD_DESIGN
End Property

Public Property Let SW_LOAD_DESIGN(ByVal blSW As Boolean)
    GO_002_SW_LOAD_DESIGN = blSW
End Property

Public Property Let TABLAS(ByVal xRS As ADODB.Recordset)
    Set GO_002_RS_TABLAS = xRS
End Property

Public Property Get TABLAS() As ADODB.Recordset
    Set TABLAS = GO_002_RS_TABLAS
End Property

Public Property Get TABLAS_CAMPO_TIPO() As String
    TABLAS_CAMPO_TIPO = GO_002_Ks_TABLAS_CAMPO_TIPO
End Property

Public Property Get TABLAS_CAMPO_CODIGO() As String
    TABLAS_CAMPO_CODIGO = GO_002_Ks_TABLAS_CAMPO_CODIGO
End Property

Public Property Get TABLAS_CAMPO_DESCRI() As String
    TABLAS_CAMPO_DESCRI = GO_002_Ks_TABLAS_CAMPO_DESCRI
End Property

Public Property Get WIN_RESULT() As GE_WIN_RESULT
    WIN_RESULT = GO_002_ENU_OPC_WIN_RESULT
End Property

Public Function GetForm() As Object
    GetForm = frm002_ECNSQLQRYDESIGN
End Function

Public Property Let WINAPPPATH(ByVal xRutaAppPath As String)
    GO_002_RUTA_APP_WIN = xRutaAppPath
End Property

Public Property Get WINAPPPATH() As String
    WINAPPPATH = GO_002_RUTA_APP_WIN
End Property

Public Sub Show()
    If Len(Trim(CADENA_CONEXION)) = 0 Then
        MsgBox "No ha establecido la cadena de conexión...", vbCritical, "ECN : clsECNSQLQRYDESIGN"
        Exit Sub
    End If
    Load frm002_ECNSQLQRYDESIGN
    With frm002_ECNSQLQRYDESIGN
        .Show
    End With
    Set frm002_ECNSQLQRYDESIGN = Nothing
End Sub

Public Sub ShowPrompt()
    If Len(Trim(CADENA_CONEXION)) = 0 Then
        MsgBox "No ha establecido la cadena de conexión...", vbCritical, "ECN : clsECNSQLQRYDESIGN"
        Exit Sub
    End If
    Load frm002_ECNSQLQRYDESIGN
    With frm002_ECNSQLQRYDESIGN
        .Show 1
    End With
    Set frm002_ECNSQLQRYDESIGN = Nothing
End Sub

Private Sub Class_Terminate()
    Set GO_002_RS_TABLAS = Nothing
End Sub
