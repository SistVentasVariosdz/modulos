VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNSQLHELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EruIndicacion
    ConParametros = 1
    SinParametros = 0
End Enum

Private strCn As String
Private bUltimo As Byte
Private blSW_ResultOK_En_RetornaRsCad As Boolean
Private blSW_ResultOK_En_RetornaRsSP As Boolean
Private blSW_ResultOK_En_RetornaRsVista As Boolean


Public Property Get CADENA_CONEXION() As String
    CADENA_CONEXION = strCn
End Property

Public Property Let CADENA_CONEXION(ByVal sCadenaConexion As String)
    strCn = sCadenaConexion
End Property

Public Property Get SW_RPTA_RetornaRsCad() As String
    SW_RPTA_RetornaRsCad = blSW_ResultOK_En_RetornaRsCad
End Property

Public Property Let SW_RPTA_RetornaRsCad(ByVal xSW_ResultOK_En_RetornaRsCad As String)
    blSW_ResultOK_En_RetornaRsCad = xSW_ResultOK_En_RetornaRsCad
End Property

Public Property Get SW_RPTA_RetornaRsSP() As String
    SW_RPTA_RetornaRsSP = blSW_ResultOK_En_RetornaRsSP
End Property

Public Property Let SW_RPTA_RetornaRsSP(ByVal xSW_ResultOK_En_RetornaRsSP As String)
    blSW_ResultOK_En_RetornaRsSP = xSW_ResultOK_En_RetornaRsSP
End Property

Public Property Get SW_RPTA_RetornaRsVista() As String
    SW_RPTA_RetornaRsVista = blSW_ResultOK_En_RetornaRsVista
End Property

Public Property Let SW_RPTA_RetornaRsVista(ByVal xSW_ResultOK_En_RetornaRsSP As String)
    blSW_ResultOK_En_RetornaRsVista = xSW_ResultOK_En_RetornaRsSP
End Property

Public Function CreaCmdConCadSQL(ByVal sQrySQL As String, _
                                 ByVal xTipOrden As EruIndicacion, _
                                 ParamArray vParametros() As Variant) As Command
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim oCmd As New Command
    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdText
        .CommandText = sQrySQL
        If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
        Set CreaCmdConCadSQL = oCmd
    End With
    Set oCmd = Nothing
    Screen.MousePointer = vbCustom
End Function

Public Function CreaCmdConSP(ByVal sSP As String, _
                            ByVal xTipOrden As EruIndicacion, _
                            ParamArray vParametros() As Variant) As Command
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim oCmd As New Command
    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdStoredProc
        .CommandText = sSP
        If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
        Set CreaCmdConSP = oCmd
    End With
    Set oCmd = Nothing
    Screen.MousePointer = vbCustom
End Function


Private Sub CreaParametros(oCmd As ADODB.Command, ByVal aParametros As Variant)
    On Error Resume Next
    
    Dim sNombre As String
    Dim Tipo As DataTypeEnum
    Dim Direccion As ParameterDirectionEnum
    Dim lTamaño As Long
    Dim Valor As Variant
    Dim I As Integer
    
    Dim oPrm As Parameter
    
    If Not IsArray(aParametros) Then Exit Sub
    bUltimo = UBound(aParametros)
    For I = 0 To (bUltimo)
        sNombre = aParametros(I)(0)
        Tipo = aParametros(I)(1)
        Direccion = aParametros(I)(2)
        lTamaño = aParametros(I)(3)
        Valor = aParametros(I)(4)
        
        oCmd.Parameters.Append oCmd.CreateParameter(sNombre, Tipo, Direccion, lTamaño, Valor)
    Next I
End Sub

Public Function RetornaRsSP(ByVal xSp As String, _
                            ByVal xTipOrden As EruIndicacion, _
                            ByVal blSW_Desconectado As Boolean, _
                            ParamArray vParametros() As Variant) As ADODB.Recordset
    On Error GoTo SALTO_ERROR
    Screen.MousePointer = vbHourglass
    Dim oCmd As New ADODB.Command
    Dim oRs As New ADODB.Recordset

    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdStoredProc
        .CommandText = xSp
    End With
    
    If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
    oCmd.Execute
    oRs.CursorLocation = adUseClient
    oRs.Open oCmd, , adOpenStatic, adLockOptimistic 'adLockReadOnly
    Set oCmd = Nothing
    If blSW_Desconectado = True Then
        Set RetornaRsSP = DevuelveRecordSetDesconectado(oRs)
    Else
        Set RetornaRsSP = oRs
    End If
    Set oRs = Nothing
    Screen.MousePointer = vbCustom
    Exit Function
    
SALTO_ERROR:
    Screen.MousePointer = vbCustom
    Set oCmd = Nothing
    Set oRs = Nothing
    MsgBox Err.Number & Space(1) & ":" & Space(2) & Err.Description, vbCritical, "ECNSQLDAO"
End Function

Public Function RetornaRsCad(ByVal xCad As String, _
                             ByVal xTipOrden As EruIndicacion, _
                             ByVal blSW_Desconectado As Boolean, _
                             ParamArray vParametros() As Variant) As Recordset
    On Error GoTo SALTO_ERROR
    Screen.MousePointer = vbHourglass
    Dim oCmd As New ADODB.Command
    Dim oRs As New ADODB.Recordset
    
    blSW_ResultOK_En_RetornaRsCad = False
    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdText
        .CommandText = xCad
        .Execute
    End With
    If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
    oRs.CursorLocation = adUseClient
    oRs.Open oCmd, , adOpenStatic, adLockOptimistic
    Set oCmd = Nothing
    If blSW_Desconectado = True Then
        Set RetornaRsCad = DevuelveRecordSetDesconectado(oRs)
    Else
        Set RetornaRsCad = oRs
    End If
    Set oRs = Nothing
    Screen.MousePointer = vbCustom
    blSW_ResultOK_En_RetornaRsCad = True
    Exit Function
    
SALTO_ERROR:
    blSW_ResultOK_En_RetornaRsCad = False
    Screen.MousePointer = vbCustom
    Set oCmd = Nothing
    Set oRs = Nothing
    MsgBox Err.Number & Space(1) & ":" & Space(2) & Err.Description, vbCritical, "ECNSQLDAO"
End Function

Public Function RetornaRsVista(ByVal xVista As String, _
                               ByVal xTipOrden As EruIndicacion, _
                               ByVal blSW_Desconectado As Boolean, _
                               ParamArray vParametros() As Variant) As Recordset
    On Error GoTo SALTO_ERROR
    Screen.MousePointer = vbHourglass
    Dim oCmd As New ADODB.Command
    Dim oRs As New ADODB.Recordset

    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdTable
        .CommandText = xVista
        .Execute
    End With
    If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
    oRs.CursorLocation = adUseClient
    oRs.Open oCmd, , adOpenStatic, adLockOptimistic
    Set oCmd = Nothing
    If blSW_Desconectado = True Then
        Set RetornaRsVista = DevuelveRecordSetDesconectado(oRs)
    Else
        Set RetornaRsVista = oRs
    End If
    Set oRs = Nothing
    Screen.MousePointer = vbCustom
    Exit Function
    
SALTO_ERROR:
    Screen.MousePointer = vbCustom
    Set oCmd = Nothing
    Set oRs = Nothing
    MsgBox Err.Number & Space(1) & ":" & Space(2) & Err.Description, vbCritical, "ECNSQLDAO"
End Function

Public Function EjecutaSp(xSp As String, _
                          xTipOrden As EruIndicacion, _
                          ParamArray vParametros() As Variant) As Boolean
    On Error GoTo SALTO_ERROR
    Screen.MousePointer = vbHourglass
    Dim oCmd As New ADODB.Command
    
    EjecutaSp = False
    With oCmd
        .ActiveConnection = strCn
        .CommandType = adCmdStoredProc
        .CommandText = xSp
    End With
    If xTipOrden = 1 Then Call CreaParametros(oCmd, vParametros)
    oCmd.Execute
    Set oCmd = Nothing
    Screen.MousePointer = vbCustom
    EjecutaSp = True
    Exit Function
    
SALTO_ERROR:
    EjecutaSp = False
    Screen.MousePointer = vbCustom
    Set oCmd = Nothing
    MsgBox Err.Number & Space(1) & ":" & Space(2) & Err.Description, vbCritical, "ECNSQLDAO - EjecutaSp"
End Function

Public Function DevuelveTipoDeDatoSQL(ByVal sCadTipoDeDatoSQL As String) As DataTypeEnum
    Dim vTipoDeDatoSQL As DataTypeEnum
    
    sCadTipoDeDatoSQL = UCase(Trim(sCadTipoDeDatoSQL))
    Select Case sCadTipoDeDatoSQL
        Case "CHAR": vTipoDeDatoSQL = adChar
        Case "NCHAR": vTipoDeDatoSQL = adChar
        Case "VARCHAR": vTipoDeDatoSQL = adVarChar
        Case "NVARCHAR": vTipoDeDatoSQL = adVarChar
        Case "BIT": vTipoDeDatoSQL = adInteger
        Case "INT": vTipoDeDatoSQL = adInteger
        Case "SMALLINT": vTipoDeDatoSQL = adSmallInt
        Case "TINYINT": vTipoDeDatoSQL = adTinyInt
        Case "NUMERIC": vTipoDeDatoSQL = adNumeric
        Case "MONEY": vTipoDeDatoSQL = adCurrency
        Case "BINARY": vTipoDeDatoSQL = adBinary
        Case "IMAGE":: vTipoDeDatoSQL = adBinary
        Case "DATETIME": vTipoDeDatoSQL = adDate
        Case "SMALLDATETIME": vTipoDeDatoSQL = adDate
        Case "DECIMAL": vTipoDeDatoSQL = adDecimal
        Case "XML": vTipoDeDatoSQL = adVarChar
    End Select
    DevuelveTipoDeDatoSQL = vTipoDeDatoSQL
End Function

Public Function DevuelveCampo(ByVal sQuerySQL As String) As Variant
    On Error GoTo SALTO_ERROR
    
    Dim oRs As New ADODB.Recordset
    With oRs
        .CursorLocation = adUseClient
        .Open sQuerySQL, CADENA_CONEXION, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            DevuelveCampo = oRs.Fields(0).Value
        Else
            DevuelveCampo = Empty
        End If
    End With
    Set oRs = Nothing
    Exit Function
    
SALTO_ERROR:
    DevuelveCampo = Empty
    Set oRs = Nothing
    Screen.MousePointer = vbCustom
    MsgBox Err.Number & Space(1) & ":" & Space(2) & Err.Description, vbCritical, "ECNSQLDAO - DevuelveCampo"
End Function

Function DevuelveFechaServidor() As Date
    DevuelveFechaServidor = DevuelveCampo("SELECT FECHA_SQL_SERVER =  GETDATE()")
End Function

Public Function DevuelveRecordSetDesconectado(ByVal oRsOrigen As ADODB.Recordset) As ADODB.Recordset
    Set DevuelveRecordSetDesconectado = Nothing
    If oRsOrigen Is Nothing Then Exit Function
    
    Dim oRsResult As ADODB.Recordset
    Dim oField As Object

    
    Set oRsResult = New ADODB.Recordset
    With oRsResult
        .CursorLocation = adUseClient
        Set .ActiveConnection = Nothing
        For Each oField In oRsOrigen.Fields
            .Fields.Append oField.Name, oField.Type, oField.DefinedSize, adFldIsNullable
            
            .Fields(oField.Name).NumericScale = oRsOrigen.Fields(oField.Name).NumericScale
            .Fields(oField.Name).DefinedSize = oRsOrigen.Fields(oField.Name).DefinedSize
            .Fields(oField.Name).Precision = oRsOrigen.Fields(oField.Name).Precision
        Next
        .Open
    End With
    
    With oRsOrigen
        If .RecordCount Then
            .MoveFirst
            Do While Not .EOF
                oRsResult.AddNew
                For Each oField In oRsOrigen.Fields
                    oRsResult.Fields(oField.Name).Value = FixData(.Fields(oField.Name).Value, _
                                                                  .Fields(oField.Name))
                Next
                oRsResult.Update
                .MoveNext
            Loop
        End If
    End With

    Set DevuelveRecordSetDesconectado = oRsResult
End Function

Private Function FixData(ByRef wTexto As Variant, _
                         ByVal oField As ADODB.Field)
    If IsNull(wTexto) Or Len(Trim(wTexto)) = 0 Then
        Select Case oField.Type
            Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
                wTexto = 0
            Case adBoolean
                wTexto = False
            Case adDate
                wTexto = Empty
            Case adChar, adVarChar
                wTexto = ""
        End Select
    End If
    FixData = wTexto
End Function
