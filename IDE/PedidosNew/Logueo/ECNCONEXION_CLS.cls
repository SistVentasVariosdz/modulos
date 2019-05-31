VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls003_ECNCONEXION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum G_ENU_ModoWIN_CONEXION
    WINC_Null = 0
    WINC_Main = 1 '==> Se usa cuando la ventana CONEXION debe ejecutarse en el inicio del sistema
    WinC_Prompt = 2 '==> Se usa como un Prompt, es decir, como ventana de ayuda para traer datos
End Enum

Public Enum GE_WIN_RESULT
    WD_NULL = 0 '=> ESTADO INICIAL DE LA VENTANA
    WD_ACCEPT = 1 '=> CUANDO HACE CLIC EN EL BOTON ACEPTAR  DE LA VENTANA
    WD_CANCEL = 2 '=> CUANDO HACE CLIC EN EL BOTON CANCELAR DE LA VENTANA
    WD_CLOSEW = 3 '=> CUANDO HACE CLIC EN EL BOTON CERRAR   DE LA VENTANA
End Enum


Private Sub Class_Initialize()
    Set PU_003_ECNLIB01_FUNSUB = New ECNVB6LIB.ECNLIB01_FUNSUB
    Set PU_003_ECNLIB03_WINEVE = New ECNVB6LIB.ECNLIB03_WINEVE
    Set PU_003_ECNLIB04_EFFECTS = New ECNVB6LIB.ECNLIB04_EFFECTS
End Sub

Public Property Get RUTA_ECN_OVL() As String
    RUTA_ECN_OVL = PU_003_RUTA_ECN_OVL
End Property

Public Property Get WIN_RESULT() As GE_WIN_RESULT
    WIN_RESULT = GO_003_ENU_OPC_WIN_RESULT
End Property

Public Function GetForm() As Object
    Set GetForm = frm003_ECNCONEXION
End Function

Public Property Let WINAPPPATH(ByVal xRutaAppPath As String)
    GO_003_RUTA_ECN_APP = xRutaAppPath
    PU_003_RUTA_ECN_OVL = GO_003_RUTA_ECN_APP & "\" & PU_003_Ks_ECNOVL_NOMFILE
End Property

Public Property Get WINAPPPATH() As String
    WINAPPPATH = GO_003_RUTA_ECN_APP
End Property

Public Property Let FORM_PARENT(ByVal xForm As Object)
    GO_003_FORM_PARENT = xForm
End Property

Public Property Get FORM_PARENT() As Object
    Set FORM_PARENT = GO_003_FORM_PARENT
End Property

Public Property Let SW_LOAD(ByVal xSW As Boolean)
    GO_003_PU_SW_LOAD = xSW
End Property

Public Property Get SW_LOAD() As Boolean
    SW_LOAD = GO_003_PU_SW_LOAD
End Property

Public Property Let MODO_WIN(ByVal xOpc As G_ENU_ModoWIN_CONEXION)
    GO_003_MODO_WIN = xOpc
End Property

Public Property Get MODO_WIN() As G_ENU_ModoWIN_CONEXION
    MODO_WIN = GO_003_MODO_WIN
End Property

Public Sub ADO()
    Call PU_003_ADO
End Sub

Public Sub BD()
    Call PU_003_GetBD
End Sub

Public Sub ClicInCboBD()
    Call PU_003_Click_End_dcBD(2)
End Sub

Public Function LeerECNOVL() As Boolean
    LeerECNOVL = PU_003_LeerECNOVL
End Function

Property Get ECNOVL_SQL_USER() As String
    ECNOVL_SQL_USER = GO_003_SQL_USER
End Property

Property Get ECNOVL_SQL_UPWD() As String
    ECNOVL_SQL_UPWD = GO_003_SQL_UPWD
End Property

Property Get ECNOVL_BD_MAIN() As String
    ECNOVL_BD_MAIN = GO_003_BD_MAIN
End Property

Property Get ECNOVL_BD_IMAG() As String
    ECNOVL_BD_IMAG = GO_003_BD_IMAG
End Property

Property Get ECNOVL_BD_SEG() As String
    ECNOVL_BD_SEG = GO_003_BD_SEG
End Property
Property Get ECNOVL_CONEXION_SQL_SEG() As String
    ECNOVL_CONEXION_SQL_SEG = GO_003_CONEXION_SQL_SEG
End Property
Property Get ECNOVL_CONEXION_SQL() As String
    ECNOVL_CONEXION_SQL = GO_003_CONEXION_SQL
End Property

Property Get ECNOVL_CONEXION_XLS() As String
    ECNOVL_CONEXION_XLS = GO_003_CONEXION_XLS
End Property

Public Sub Show()
  '  Load frm003_ECNCONEXION
    With frm003_ECNCONEXION
        .Show
    End With
   ' Set frm003_ECNCONEXION = Nothing
End Sub

Public Sub ShowPrompt()
   ' Load frm003_ECNCONEXION
    With frm003_ECNCONEXION
        .Show 1
    End With
   ' Set frm003_ECNCONEXION = Nothing
End Sub

Private Sub Class_Terminate()
    Set PU_003_ECNLIB01_FUNSUB = Nothing
    Set PU_003_ECNLIB03_WINEVE = Nothing
End Sub
