VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls001_ECNSQLUSPLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum GE_WIN_RESULT
    WD_NULL = 0 '=> ESTADO INICIAL DE LA VENTANA
    WD_ACCEPT = 1 '=> CUANDO HACE CLIC EN EL BOTON ACEPTAR  DE LA VENTANA
    WD_CANCEL = 2 '=> CUANDO HACE CLIC EN EL BOTON CANCELAR DE LA VENTANA
    WD_CLOSEW = 3 '=> CUANDO HACE CLIC EN EL BOTON CERRAR   DE LA VENTANA
End Enum

Public Property Get PREFIJO_USPSQL() As String
    PREFIJO_USPSQL = GO_001_PREFIJO_USPSQL
End Property

Public Property Let PREFIJO_USPSQL(ByVal xPrefijoUSPSQL As String)
    GO_001_PREFIJO_USPSQL = xPrefijoUSPSQL
End Property

Public Property Get SW_PREFIJO_USPSQL() As Boolean
    SW_PREFIJO_USPSQL = GO_001_SW_PREFIJO_SQL
End Property

Public Property Let SW_PREFIJO_USPSQL(ByVal xUsarPrefijoUSPSQL As Boolean)
    GO_001_SW_PREFIJO_SQL = xUsarPrefijoUSPSQL
End Property

Public Property Get CADENA_CONEXION() As String
    CADENA_CONEXION = GO_001_CONEXION_SQL
End Property

Public Property Let CADENA_CONEXION(ByVal xCadenaConexion As String)
    GO_001_CONEXION_SQL = xCadenaConexion
End Property

Public Property Get USPSQL_COD_SEL() As String
    USPSQL_COD_SEL = GO_001_USPSQL_SEL_NOM
End Property

Public Property Get USPSQL_NOM_SEL() As String
    USPSQL_NOM_SEL = GO_001_USPSQL_SEL_NOM
End Property

Public Property Get USPSQL_PARAM() As ADODB.Recordset
    Set USPSQL_PARAM = GO_001_LST_PARAMETROS
End Property

Public Property Get WIN_RESULT() As GE_WIN_RESULT
    WIN_RESULT = GO_001_ENU_OPC_WIN_RESULT
End Property

Private Sub Class_Initialize()
    GO_001_PREFIJO_USPSQL = "USP"
End Sub

Public Sub Show()
    If Len(Trim(CADENA_CONEXION)) = 0 Then
        MsgBox "No ha establecido la cadena de conexión...", vbCritical, "ECN : clsECNSQLUSPLIST"
        Exit Sub
    End If
    Load frm001_ECNSQLUSPLIST
    With frm001_ECNSQLUSPLIST
        .Show
    End With
    Set frm001_ECNSQLUSPLIST = Nothing
End Sub

Public Sub ShowPrompt()
    If Len(Trim(CADENA_CONEXION)) = 0 Then
        MsgBox "No ha establecido la cadena de conexión...", vbCritical, "ECN : clsECNSQLUSPLIST"
        Exit Sub
    End If
    Load frm001_ECNSQLUSPLIST
    With frm001_ECNSQLUSPLIST
        .Show 1
    End With
    Set frm001_ECNSQLUSPLIST = Nothing
End Sub

Public Function ObtenerParametrosDeUSPSQL(ByVal sCodUSPSQL As String, _
                                          Optional ByVal blSW_ShowWinIn As Boolean = False, _
                                          Optional ByVal blSW_WinDialog As Boolean = True, _
                                          Optional ByVal sNomUSPSQL As String = "") As ADODB.Recordset
                            
    Dim oECNSQLHELP As ECNVB6LIB.ECNSQLHELP

    Set oECNSQLHELP = New ECNVB6LIB.ECNSQLHELP
    oECNSQLHELP.CADENA_CONEXION = GO_001_CONEXION_SQL
    Set ObtenerParametrosDeUSPSQL = PU_001_ObtenerParametros(oECNSQLHELP, sCodUSPSQL)
    If blSW_ShowWinIn = True Then
        Load frm001_ECNSQLUSPLIST_PARAM
        With frm001_ECNSQLUSPLIST_PARAM
            .PU_USPSQL_ID = sCodUSPSQL
            If Len(Trim(sNomUSPSQL)) > 0 Then
                .PU_USPSQL_NM = sNomUSPSQL
            End If
            Set .PU_RS_PARAMS = ObtenerParametrosDeUSPSQL
            Call .PU_CargarInfoParam
            If blSW_WinDialog Then .Show 1 Else .Show
        End With
        Set frm001_ECNSQLUSPLIST_PARAM = Nothing
    End If
End Function
