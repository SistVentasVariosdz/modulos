Attribute VB_Name = "mdl002_ECNSQLQRYDESIGN"
Option Explicit

Global Const GO_002_Ks_TABLAS_CAMPO_TIPO As String = "TIP_TABLA"
Global Const GO_002_Ks_TABLAS_CAMPO_CODIGO As String = "COD_TABLA"
Global Const GO_002_Ks_TABLAS_CAMPO_DESCRI As String = "DES_TABLA"

Global Const GO_002_Ks_TIPO_DE_ORDENAMIENTO_ASC As String = "ASC"
Global Const GO_002_Ks_TIPO_DE_ORDENAMIENTO_DES As String = "DESC"

Global Const GO_002_Ks_TIPO_DE_CAMPO_FD As String = "FIELD"
Global Const GO_002_Ks_TIPO_DE_CAMPO_TX As String = "TEXT"
Global Const GO_002_Ks_TIPO_DE_CAMPO_FX As String = "FX"
Global Const GO_002_Ks_TIPO_DE_CAMPO_AD As String = "ADDED"

Global Const GO_002_Ks_TIPO_DE_JOIN_FR As String = "FROM"
Global Const GO_002_Ks_TIPO_DE_JOIN_LF As String = "LEFT JOIN"
Global Const GO_002_Ks_TIPO_DE_JOIN_IN As String = "INNER JOIN"
Global Const GO_002_Ks_TIPO_DE_JOIN_RI As String = "RIGHT JOIN"
Global Const GO_002_Ks_TIPO_DE_JOIN_UN As String = "RIGHT JOIN"

Global Const GO_002_RES_STR_ID_Ki_AVI_EARTH As Integer = 101


Global GO_002_CONEXION_SQL As String
Global GO_002_SW_LOAD_DESIGN As Boolean
Global GO_002_ENU_OPC_WIN_RESULT As GE_WIN_RESULT
Global GO_002_RS_TABLAS As ADODB.Recordset
Global GO_002_RUTA_INI_PARAM_WIN As String
Global GO_002_RUTA_APP_WIN As String

Public Function PU_002_DevuelveRutaAviRes(ByVal sNomAviResID As String)
    PU_002_DevuelveRutaAviRes = GO_002_RUTA_APP_WIN & "\RES\AVI\" & sNomAviResID & ".avi"
End Function

Public Sub PU_002_AperturarRSdeTablas()
    On Error Resume Next
    Set GO_002_RS_TABLAS = Nothing
    Set GO_002_RS_TABLAS = New ADODB.Recordset
    With GO_002_RS_TABLAS
        .CursorLocation = adUseClient
        .Fields.Append GO_002_Ks_TABLAS_CAMPO_TIPO, adInteger, 4
        .Fields.Append GO_002_Ks_TABLAS_CAMPO_CODIGO, adVarChar, 50
        .Fields.Append GO_002_Ks_TABLAS_CAMPO_DESCRI, adVarChar, 100
        .Open
    End With
End Sub

Public Function PU_002_CargarTablas() As ADODB.Recordset
    Dim oECNSQLHELP  As New ECNVB6LIB.ECNSQLHELP
    Dim xSQL As String
    
    oECNSQLHELP.CADENA_CONEXION = GO_002_CONEXION_SQL
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT OBJECT_ID = LTRIM(RTRIM(STR(OBJECT_ID))),"
    xSQL = xSQL & vbNewLine & "       NAME,"
    xSQL = xSQL & vbNewLine & "       CREATE_DATE = CONVERT(CHAR(10),CREATE_DATE,103) + SPACE(1) + CONVERT(CHAR(8),CREATE_DATE,108),"
    xSQL = xSQL & vbNewLine & "       MODIFY_DATE = CONVERT(CHAR(10),MODIFY_DATE,103) + SPACE(1) + CONVERT(CHAR(8),MODIFY_DATE,108)"
    xSQL = xSQL & vbNewLine & "FROM SYS.TABLES"
    xSQL = xSQL & vbNewLine & "ORDER BY NAME"
    Set PU_002_CargarTablas = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, True)
End Function

Public Function PU_002_CargarVistas() As ADODB.Recordset
    Dim oECNSQLHELP  As New ECNVB6LIB.ECNSQLHELP
    Dim xSQL As String
    
    oECNSQLHELP.CADENA_CONEXION = GO_002_CONEXION_SQL
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT OBJECT_ID = LTRIM(RTRIM(STR(OBJECT_ID))),"
    xSQL = xSQL & vbNewLine & "       NAME,"
    xSQL = xSQL & vbNewLine & "       CREATE_DATE = CONVERT(CHAR(10),CREATE_DATE,103) + SPACE(1) + CONVERT(CHAR(8),CREATE_DATE,108),"
    xSQL = xSQL & vbNewLine & "       MODIFY_DATE = CONVERT(CHAR(10),MODIFY_DATE,103) + SPACE(1) + CONVERT(CHAR(8),MODIFY_DATE,108)"
    xSQL = xSQL & vbNewLine & "FROM SYS.VIEWS"
    xSQL = xSQL & vbNewLine & "ORDER BY NAME"
    Set PU_002_CargarVistas = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, True)
End Function

Public Function PU_002_AgregarTabla(ByVal opcTipTab As GE_TIPO_TABLA, _
                                    ByVal sCodTabla As String, _
                                    ByVal sDesTabla As String) As Boolean
    On Error Resume Next
    PU_002_AgregarTabla = False
    With GO_002_RS_TABLAS
        .AddNew
        .Fields(GO_002_Ks_TABLAS_CAMPO_TIPO) = CInt(opcTipTab)
        .Fields(GO_002_Ks_TABLAS_CAMPO_CODIGO) = sCodTabla
        .Fields(GO_002_Ks_TABLAS_CAMPO_DESCRI) = sDesTabla
        .Update
        .MoveLast
    End With
    PU_002_AgregarTabla = True
End Function

Public Function PU_002_EliminarTablas() As Boolean
    On Error Resume Next
    PU_002_EliminarTablas = False
    With GO_002_RS_TABLAS
        .MoveFirst
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
    End With
    PU_002_EliminarTablas = True
End Function

Public Function PU_002_CargarColumnas(ByVal sCodTabla As String) As ADODB.Recordset
    Set PU_002_CargarColumnas = Nothing
    If Len(Trim(sCodTabla)) = 0 Then Exit Function
    Dim oECNSQLHELP As New ECNVB6LIB.ECNSQLHELP
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT COD_COLUMNA = A.COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "       NOM_COLUMNA = UPPER(LTRIM(RTRIM(A.NAME))),"
    xSQL = xSQL & vbNewLine & "       IS_PK       = CASE WHEN KEY_ORDINAL = 1 THEN 1"
    xSQL = xSQL & vbNewLine & "                          ELSE 0"
    xSQL = xSQL & vbNewLine & "                     END,"
    xSQL = xSQL & vbNewLine & "       IS_UK       = CASE WHEN KEY_ORDINAL = 2 THEN 1"
    xSQL = xSQL & vbNewLine & "                          ELSE 0"
    xSQL = xSQL & vbNewLine & "                     END,"
    xSQL = xSQL & vbNewLine & "       IS_FK       = CASE WHEN C.PARENT_COLUMN_ID IS NULL THEN 0"
    xSQL = xSQL & vbNewLine & "                          ELSE 1"
    xSQL = xSQL & vbNewLine & "                     End"
    xSQL = xSQL & vbNewLine & "FROM SYS.COLUMNS AS A"
    xSQL = xSQL & vbNewLine & "LEFT JOIN SYS.INDEX_COLUMNS AS B"
    xSQL = xSQL & vbNewLine & "    ON A.COLUMN_ID = B.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "   AND A.OBJECT_ID = B.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "   AND B.KEY_ORDINAL IN (1,2)"
    xSQL = xSQL & vbNewLine & "LEFT JOIN SYS.FOREIGN_KEY_COLUMNS AS C"
    xSQL = xSQL & vbNewLine & "    ON C.PARENT_COLUMN_ID = A.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "   AND C.PARENT_OBJECT_ID = A.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "WHERE A.OBJECT_ID =  '" & sCodTabla & "'"

    oECNSQLHELP.CADENA_CONEXION = GO_002_CONEXION_SQL
    Set PU_002_CargarColumnas = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, True)
End Function

Public Function PU_002_ObtenerRelacionesPKFK(ByVal sCodTabla As String) As ADODB.Recordset
    Set PU_002_ObtenerRelacionesPKFK = Nothing
    If Len(Trim(sCodTabla)) = 0 Then Exit Function
    Dim oECNSQLHELP As New ECNVB6LIB.ECNSQLHELP
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT PK_TABLA_COD  = A.REFERENCED_OBJECT_ID,"
    xSQL = xSQL & vbNewLine & "       PK_TABLA_DES  = C.NAME,"
    xSQL = xSQL & vbNewLine & "       PK_COLUMN_COD = A.REFERENCED_COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "       PK_COLUMN_DES = D.NAME,"
    xSQL = xSQL & vbNewLine & "       FK_TABLA_COD  = A.PARENT_OBJECT_ID,"
    xSQL = xSQL & vbNewLine & "       FK_TABLA_DES  = B.NAME,"
    xSQL = xSQL & vbNewLine & "       FK_COLUMN_COD = A.PARENT_COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "       FK_COLUMN_DES = E.NAME"
    xSQL = xSQL & vbNewLine & "FROM SYS.FOREIGN_KEY_COLUMNS AS A"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.TABLES        AS B"
    xSQL = xSQL & vbNewLine & "    ON A.PARENT_OBJECT_ID = B.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.TABLES  AS C"
    xSQL = xSQL & vbNewLine & "    ON A.REFERENCED_OBJECT_ID = C.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.COLUMNS   AS D"
    xSQL = xSQL & vbNewLine & "    ON A.REFERENCED_OBJECT_ID = D.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "   AND A.REFERENCED_COLUMN_ID = D.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.COLUMNS   AS E"
    xSQL = xSQL & vbNewLine & "    ON A.PARENT_OBJECT_ID     = E.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "   AND A.PARENT_COLUMN_ID     = E.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "WHERE A.REFERENCED_OBJECT_ID  = '" & sCodTabla & "'"
    xSQL = xSQL & vbNewLine & "ORDER BY A.REFERENCED_COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "         A.PARENT_COLUMN_ID"
    
    oECNSQLHELP.CADENA_CONEXION = GO_002_CONEXION_SQL
    Set PU_002_ObtenerRelacionesPKFK = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, True)
End Function

Public Function PU_002_ObtenerKeys(ByVal sCodTabla As String) As ADODB.Recordset
    Set PU_002_ObtenerKeys = Nothing
    If Len(Trim(sCodTabla)) = 0 Then Exit Function
    Dim oECNSQLHELP As New ECNVB6LIB.ECNSQLHELP
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT COD_COLUMNA = A.COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "       NOM_COLUMNA = UPPER(LTRIM(RTRIM(B.NAME))),"
    xSQL = xSQL & vbNewLine & "       TIP_COLUMNA = CASE WHEN KEY_ORDINAL = 1 THEN '3PK'"
    xSQL = xSQL & vbNewLine & "                          WHEN KEY_ORDINAL = 2 THEN '1UK'"
    xSQL = xSQL & vbNewLine & "                     End"
    xSQL = xSQL & vbNewLine & "FROM SYS.INDEX_COLUMNS AS A"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.COLUMNS AS B"
    xSQL = xSQL & vbNewLine & "    ON A.COLUMN_ID = B.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "   AND A.object_id = B.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "WHERE A.OBJECT_ID =  '" & sCodTabla & "'"
    xSQL = xSQL & vbNewLine & "AND   KEY_ORDINAL IN (1,2)"
    xSQL = xSQL & vbNewLine & ""
    xSQL = xSQL & vbNewLine & "Union All"
    xSQL = xSQL & vbNewLine & ""
    xSQL = xSQL & vbNewLine & "SELECT COD_COLUMNA = PARENT_COLUMN_ID,"
    xSQL = xSQL & vbNewLine & "       NOM_COLUMNA = UPPER(LTRIM(RTRIM(B.NAME))),"
    xSQL = xSQL & vbNewLine & "       TIP_COLUMNA = '2FK'"
    xSQL = xSQL & vbNewLine & "FROM SYS.FOREIGN_KEY_COLUMNS AS A"
    xSQL = xSQL & vbNewLine & "INNER JOIN SYS.COLUMNS       AS B"
    xSQL = xSQL & vbNewLine & "    ON A.PARENT_COLUMN_ID = B.COLUMN_ID"
    xSQL = xSQL & vbNewLine & "   AND A.PARENT_OBJECT_ID = B.OBJECT_ID"
    xSQL = xSQL & vbNewLine & "WHERE PARENT_OBJECT_ID =  '" & sCodTabla & "'"

    oECNSQLHELP.CADENA_CONEXION = GO_002_CONEXION_SQL
    Set PU_002_ObtenerKeys = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, True)
    PU_002_ObtenerKeys.Sort = "COD_COLUMNA,TIP_COLUMNA"
End Function
