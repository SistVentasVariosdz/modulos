Attribute VB_Name = "mdl001_ECNSQLUSPLIST"
Option Explicit

Global GO_001_PREFIJO_USPSQL As String
Global GO_001_SW_PREFIJO_SQL As Boolean
Global GO_001_CONEXION_SQL As String
Global GO_001_USPSQL_SEL_COD As String
Global GO_001_USPSQL_SEL_NOM  As String
Global GO_001_LST_PARAMETROS As ADODB.Recordset
Global GO_001_ENU_OPC_WIN_RESULT As GE_WIN_RESULT

Public Function PU_001_ObtenerParametros(ByVal objECNSQLHELP As ECNVB6LIB.ECNSQLHELP, _
                                         ByVal sCodUSPSQL As String) As ADODB.Recordset
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT PARAMETER_ID,"
    xSQL = xSQL & vbNewLine & "       NAME = UPPER(A.NAME),"
    xSQL = xSQL & vbNewLine & "       A.SYSTEM_TYPE_ID,"
    xSQL = xSQL & vbNewLine & "       SYSTEM_TYPE_NM = B.NAME,"
    xSQL = xSQL & vbNewLine & "       A.USER_TYPE_ID,"
    xSQL = xSQL & vbNewLine & "       USER_TYPE_ID_NM = ISNULL(C.NAME,''),"
    xSQL = xSQL & vbNewLine & "       A.MAX_LENGTH,"
    xSQL = xSQL & vbNewLine & "       A.PRECISION,"
    xSQL = xSQL & vbNewLine & "       A.SCALE,"
    xSQL = xSQL & vbNewLine & "       IS_OUTPUT         = LTRIM(RTRIM(STR(IS_OUTPUT))),"
    xSQL = xSQL & vbNewLine & "       IS_READONLY       = LTRIM(RTRIM(STR(A.IS_READONLY))),"
    xSQL = xSQL & vbNewLine & "       IS_XML_DOCUMENT   = LTRIM(RTRIM(STR(A.IS_XML_DOCUMENT))),"
    xSQL = xSQL & vbNewLine & "       HAS_DEFAULT_VALUE = LTRIM(RTRIM(STR(A.HAS_DEFAULT_VALUE))),"
    xSQL = xSQL & vbNewLine & "       DEFAULT_VALUE = ISNULL(A.DEFAULT_VALUE,'<NULL>')"
    xSQL = xSQL & vbNewLine & "FROM SYS.ALL_PARAMETERS AS A"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.TYPES    AS B"
    xSQL = xSQL & vbNewLine & "     ON A.SYSTEM_TYPE_ID = B.SYSTEM_TYPE_ID"
    xSQL = xSQL & vbNewLine & "LEFT JOIN SYS.TYPES    AS C"
    xSQL = xSQL & vbNewLine & "     ON A.USER_TYPE_ID = C.USER_TYPE_ID"
    xSQL = xSQL & vbNewLine & "WHERE OBJECT_ID = " & sCodUSPSQL
    xSQL = xSQL & vbNewLine & "ORDER BY PARAMETER_ID"

    Set PU_001_ObtenerParametros = objECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, False)
End Function
