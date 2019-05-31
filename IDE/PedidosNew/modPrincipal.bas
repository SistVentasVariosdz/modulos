Attribute VB_Name = "modPrincipal"
Option Explicit
Option Base 0

Dim sQuery As String

Global Const kFORMAT_TO_PRINT = "dd/mm/yyyy"

Global Const kFORMAT_TO_PRINTSHORT = "dd/mm/yy"

Declare Function DeleteFile _
        Lib "kernel32" _
        Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Const CB_FINDSTRING = &H14C

Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const BlockSize = 100000

Global Ruta_Logo_Empresa   As String

Global Num_Ruc_Empresa     As String

Global Direccion_Empresa   As String

Global DSN_Empresa         As String

Global DSN_Seguridad       As String

Global Ruta0_Empresa       As String

Global Fecha_Hora_Conexion As String

Public vusu                As String

Public vemp                As String

Public vper                As String

Public vemp1               As String

Public vpas                As String

Public vNomFor             As String

Public vRuta               As String

Public iLanguage           As Integer

Public cn                  As ADODB.Connection

Public sDllName            As String

Public oFormObjDLL         As Object

Public objFormDLL          As Object

Public oFormObjDLL2        As Object

Public iRowsGrilla         As Long

' Constantes de la Aplicación
Public Const M_NUEVO       As String = "NUEVO"

Public Const M_EDITAR      As String = "EDITAR"

Public Const M_ELIMINAR    As String = "ELIMINAR"

' Tipo de Dato Message
Type RegMessage

    StringParm As String
    Cancel As Boolean

End Type

' Variable Publica
Public Message As RegMessage

' Registro Detalle Pedido
Type RegDetallePedido

    IdArticulo As String * 8
    Nombre As String * 35
    precio As Currency
    Cantidad As Integer
    Cancel As Boolean
    Accion As String * 6

End Type

Public Enum FIELD

    COL_NAME
    COL_TYPE
    COL_DESCRIPLARGA
    COL_DESCRIPTION
    COL_LENGTH
    COL_MIN
    COL_MAX
    COL_DEFAULT
    COL_DES_CORTA
    COL_DES_ABREVIADA

End Enum

Public Enum DataTypeEnum

    adBigInt = 20
    adBinary = 128
    adBoolean = 11
    adBSTR = 8
    adChar = 129
    adCurrency = 6
    adDate = 7
    adDBDate = 133
    adDBTime = 134
    adDBTimeStamp = 135
    adDecimal = 14
    adDouble = 5
    adEmpty = 0
    adError = 10
    adGUID = 72
    adIDispatch = 9
    adInteger = 3
    adIUnknown = 13
    adLongVarBinary = 205
    adLongVarChar = 201
    adLongVarWChar = 203
    adNumeric = 131
    adSingle = 4
    adSmallInt = 2
    adTinyInt = 16
    adUnsignedBigInt = 21
    adUnsignedInt = 19
    adUnsignedSmallInt = 18
    adUnsignedTinyInt = 17
    adUserDefined = 132
    adVarBinary = 204
    adVarChar = 200
    adVariant = 12
    adVarWChar = 202
    adWChar = 130

End Enum

Public Enum FieldEnum

    IName = 0
    iActualSize = 1
    iAttributes = 2
    iDefinedSize = 3
    iNumericScale = 4
    iOriginalValue = 5
    iPrecision = 6
    iType = 7
    iUnderlyingValue = 8
    iValue = 9
    iMaxEnumField = 9

End Enum

' Variable Detalle Pedido
Public DetPedido              As RegDetallePedido, cServidor As String

Public cCONNECT               As String

Public cSEGURIDAD             As String

Public bCargaConexion         As Boolean

Public Const cCLASS_TG_PURORD As String = "Visuales.clsTG_PurOrd"

Public Const kCHR_BOLD_IN     As String = "E"

Public Const kCHR_BOLD_OUT    As String = "F"

Public Const kKEY_SEPARATOR   As String = "-+-"

Public Enum MESSAGECODE

    kMESSAGE_ERR_NOTEMPTY = 401
    kMESSAGE_ERR_FOUND = 402
    kMESSAGE_ERR_NOTFOUND = 403
    kMESSAGE_ERR_USERCONNECTFAIL = 405
    kMESSAGE_ERR_CODIGO_YA_REGISTRADO = 406
    kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO = 407
    kMESSAGE_ERR_REGISTRO_TIENE_TRANSAC_RELACIONADAS = 408
    kMESSAGE_ERR_FILE_NOT_FOUND = 409
    kMESSAGE_ERR_PROCESS_INSATISFACT = 410
    kMESSAGE_ERR_LOTEST_CLOSED = 456
    kMESSAGE_ERR_STYLE_HAVE_MORE_ESTPRO = 458
    kMESSAGE_INF_PROCESS_SATISFACTO = 102
    kMESSAGE_INF_NO_INIT_SEARCH = 103
    kMESSAGE_INF_DATA_NOTFOUND = 411
    kMESSAGE_INF_FILE_PRINT_OK = 104
    kMESSAGE_INF_NEW_CODIGO = 106
    kMESSAGE_INF_DATA_SAVE = 107
    kMESSAGE_INF_DATA_DELETE = 108
    
    kMESSAGE_WAR_ENABLED_DELETED = 301
    kMESSAGE_WAR_CONFIR_CHANGES = 302
    
    kMESSAGE_ASK_PRINT_FILE = 201
    kMESSAGE_ASK_EXIT_SYSTEM = 202
    kMESSAGE_ASK_PROCESS = 203
    kMESSAGE_ASK_MAILING_FILE = 204
    kMESSAGE_ASK_DELETE_PURORD = 207
    kMESSAGE_ASK_DELETE_LOTEST = 208
    
    kMESSAGE_ERR_VALIDA_COD_CLIENTE = 413
    kMESSAGE_ERR_VALIDA_NOM_CLIENTE = 414
    kMESSAGE_ERR_VALIDA_DES_CLIENTE = 415
    kMESSAGE_ERR_VALIDA_DES_DIVISION = 416
    kMESSAGE_ERR_VALIDA_COD_DIVISION = 417
    kMESSAGE_ERR_VALIDA_DES_COLOR = 418
    kMESSAGE_ERR_VALIDA_COD_COLOR = 419
    kMESSAGE_ERR_VALIDA_COD_ESTCLI = 420
    kMESSAGE_ERR_VALIDA_NOM_ESTCLI = 421
    kMESSAGE_ERR_VALIDA_TIP_ESTCLI = 422
    kMESSAGE_ASK_NUEVO_ESTCLI = 205
    kMESSAGE_ASK_NUEVO_PURORD = 206
    
    kMESSAGE_ERR_VALIDA_COD_COMI = 423
    kMESSAGE_ERR_VALIDA_DES_COMI = 424
    
    kMESSAGE_ERR_VALIDA_COD_DESTINO = 425
    kMESSAGE_ERR_VALIDA_DES_DESTINO = 426
    kMESSAGE_ERR_VALIDA_COD_FABRICA = 427
    kMESSAGE_ERR_VALIDA_ABR_FABRICA = 428
    kMESSAGE_ERR_VALIDA_NOM_FABRICA = 429
    kMESSAGE_ERR_VALIDA_DES_FABRICA = 430
    kMESSAGE_ERR_VALIDA_DIR_FABRICA = 431
    kMESSAGE_ERR_VALIDA_TEL_FABRICA = 432
    
    kMESSAGE_ERR_VALIDA_COD_MONEDA = 433
    kMESSAGE_ERR_VALIDA_DES_MONEDA = 434
    
    kMESSAGE_ERR_VALIDA_COD_ORGANIZ = 435
    kMESSAGE_ERR_VALIDA_DES_ORGANIZ = 436
    
    kMESSAGE_ERR_VALIDA_COD_PAGO = 437
    kMESSAGE_ERR_VALIDA_DES_PAGO = 438
    
    kMESSAGE_ERR_VALIDA_COD_FABRES = 439
    kMESSAGE_ERR_VALIDA_COD_CLIRES = 440
    kMESSAGE_ERR_VALIDA_RES_CLIENTE = 441
    
    kMESSAGE_ERR_VALIDA_COD_TEMP = 442
    kMESSAGE_ERR_VALIDA_DES_TEMP = 443
    
    kMESSAGE_ERR_VALIDA_PORC_CLIENTE = 444
    
    kMESSAGE_ERR_VALIDA_COD_TIPEMB = 445
    kMESSAGE_ERR_VALIDA_DES_TIPEMB = 446
    
    kMESSAGE_ERR_VALIDA_COD_TIPPRE = 447
    kMESSAGE_ERR_VALIDA_DES_TIPPRE = 448
    
    kMESSAGE_ERR_VALIDA_COD_UM = 449
    kMESSAGE_ERR_VALIDA_DES_UM = 450
    
    kMESSAGE_ERR_VALIDA_ANO_MES = 451
    
    kMESSAGE_ERR_VALIDA_SERIE = 452
    
    kMESSAGE_ERR_VALIDA_FACTURA = 453
    kMESSAGE_ERR_ASIGN_STYLE_TEMCLI = 457
    kMESSAGE_ERR_EXIST_CLIENT = 459
    kMESSAGE_ERR_NOT_RIGHT_OPTION = 460
    kMESSAGE_ERR_INVALID_SELECC = 461

End Enum

Public Enum TypeMante

    kMANT_ADICIONAR = 0
    kMANT_MODIFICAR = 1
    kMANT_ELIMINAR = 2
    kMANT_CONSULTAR = 3
    kMANT_BUSCAR = 4
    kMANT_IMPRIMIR = 5
    kMANT_SALIR = 6

End Enum

Public Const kMESSAGE_PROVEEDOR_PASADOS    As String = ""

Public Const kMESSAGE_PROVEEDOR_PENDIENTES As String = ""

Public Sub FormSet(ByRef FormMe As Form)

    Dim oControl     As Control

    Dim oDiccionario As Object

    Dim vbuff        As Variant

    Dim sUserActions As String

    Set oDiccionario = Nothing
    Centrar_form FormMe
    IdiomaEtiquetas FormMe
End Sub

Public Function SeekInTag(ByVal sTag As String, ByVal sValue As String) As Boolean
    SeekInTag = False
    
    Select Case InStr(sTag, sValue)

        Case Is = 0
            SeekInTag = False

        Case Is > 0
            SeekInTag = True

        Case Else
            SeekInTag = True
    End Select
    
End Function

Public Sub LoadRutas(ByVal vnewvalue As Variant)
    
End Sub

Public Sub LoadConnectEmpresa(ByVal vnewvalue As String)
    
    If Not bCargaConexion Then
        'iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
        cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=RGS;UID=sa;pwd=;"
    End If

End Sub

Public Sub LoadConnectSeguridad(ByVal vnewvalue As String)

    If Not bCargaConexion Then
        cSEGURIDAD = "Provider=sqloledb;Server=Servidor;Database=Seguridad;UID=sa;pwd=;"
        'Set B_db = Nothing
        'B_db.Open cSEGURIDAD
    End If

End Sub

Function FixNulos(wtexto As Variant, wTipo As Integer)

    If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then

        Select Case wTipo

            Case 2, 3, 4, 5
                wtexto = 0

            Case 7
                wtexto = Empty '(" Empty 'Format$("", "mm/dd/yyyy")

            Case 8
                wtexto = ""

            Case 11
                wtexto = False
        End Select

    End If

    FixNulos = wtexto
End Function

Public Function SetData(ByVal sData As Variant, Optional iMaxLen As Integer) As String

    Dim iLen As String

    sData = FixNulos(sData, vbstring)

    If IsMissing(iMaxLen) Then
        SetData = sData
    Else
        iLen = Len(sData)
            
        If iLen <= iMaxLen Then
            SetData = Rpad(sData, iMaxLen)
        Else
            SetData = Mid(sData, 1, iMaxLen)
        End If
    End If

End Function

Public Function UnloadForm(ByRef forma As Form)

    Dim oControl As Control

    For Each oControl In forma.Controls

        '        If TypeOf oControl Is TreeFlexGrid.TreeGrid Then
        '            If Not oControl.RefObject Is Nothing Then
        '                Set oControl.RefObject = Nothing
        '            End If
        '            oControl.Term
        '
        '        End If
        '        If TypeOf oControl Is ActionsButton Then
        '            oControl.Term
        '        End If
    Next

End Function

Function GetCommandLine(Optional MaxArgs)

    'Declare variables.
    Dim Argarray()
    
    Dim c, cmdLine, CmdLnLen, InArg, i, NumArgs

    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Make array of the correct size.
    ReDim Argarray(MaxArgs)
    NumArgs = 0: InArg = False
    'Get command line arguments.
    cmdLine = Command()
    CmdLnLen = Len(cmdLine)

    'Go thru command line one character
    'at a time.
    For i = 1 To CmdLnLen
        c = Mid(cmdLine, i, 1)

        'Test for space or tab.
        If (c <> " " And c <> vbTab) Then

            'Neither space nor tab.
            'Test if already in argument.
            If Not InArg Then

                'New argument begins.
                'Test for too many arguments.
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                InArg = True
            End If

            'Concatenate character to current argument.
            Argarray(NumArgs) = Argarray(NumArgs) & c
        Else
            'Found a space or tab.

            'Set InArg flag to False.
            InArg = False
        End If

    Next i

    'Resize array just enough to hold arguments.
    ReDim Preserve Argarray(NumArgs)
    'Return Array in Function name.
    GetCommandLine = Argarray()
End Function

Public Function ToUpper(KeyAscii As Integer) As String

    ToUpper = Asc(UCase(Chr(KeyAscii)))
End Function

Public Function RestoreRowSSDBGrid(ByRef grid As Object, _
                                   ByVal irow As Variant, _
                                   Optional ByVal iRows As Variant)

    If IsMissing(iRows) Then
        iRows = grid.Rows
    End If

    If grid.Rows = iRows Then
        grid.Bookmark = irow
    Else
        grid.Bookmark = 0
        grid.FirstRow = 0
    End If
    
End Function

Public Sub LibraryVBToSSDBGrid(ByRef oData As Object, _
                               ByRef pBuff As Variant, _
                               ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)

    On Error Resume Next

    Dim rsBuff    As LibraryVB.clsRecords

    Dim iContador As Long

    Dim nCols     As Integer

    Dim iVerif    As Integer

    Dim temp      As String

    Dim NVEZ      As Boolean

    Dim X%

    Dim total1    As Long

    Dim y%

    Dim i         As Long

    Dim ic        As Long

    ssDBGrid.FieldSeparator = "~"
    Set rsBuff = New LibraryVB.clsRecords
    Set rsBuff.RefObject = oData

    rsBuff.Buffer = pBuff
    ssDBGrid.Redraw = False
    nCols = rsBuff.count

    ic = ssDBGrid.Cols
 
    If ssDBGrid.Cols < nCols Then

        For i = nCols To ic + 1 Step -1
            ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
            ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
            ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
        Next i

    End If

    For y = 0 To ssDBGrid.Cols - 1

        If ssDBGrid.Columns(y).DataType = 5 Or ssDBGrid.Columns(y).DataType = 6 Or ssDBGrid.Columns(y).DataType = 9 Then
            ssDBGrid.Columns(y).TagVariant = 0
        End If

    Next

    NVEZ = True

    X = 0

    Do While Not rsBuff.EOF
        temp = ""

        For iContador = 0 To nCols - 1

            If NVEZ Then
                '        If Mid(ssDBGrid.Columns(iContador).Caption, 1, 1) = "*" Then
                '            ssDBGrid.Columns(iContador).Caption = oColeccion(rsBuff(iContador + 1).Name).Description ' .DescripCorta
                '
                '            Select Case oColeccion(rsBuff(iContador + 1).Name).TypeField
                '            Case "Alfabético/Alfanumérico"
                '                ssDBGrid.Columns(iContador).DataType = 8
                '            Case "Decimal/Moneda"
                '                ssDBGrid.Columns(iContador).DataType = 5
                '            Case "Fecha"
                '                ssDBGrid.Columns(iContador).DataType = 7
                '            End Select
                '        End If
            End If

            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
            temp = temp & FixNulos(rsBuff(iContador + 1), vbstring)

            If iContador < nCols - 1 Then
                temp = temp & "~"
            End If

            If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
                ssDBGrid.Columns(iContador).DataType = 5
                ssDBGrid.Columns(iContador).Alignment = 1
            End If

            'ssDbgrid.Columns(iContador).DataType = 5
            If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBGrid.TagVariant, vbLong) Then
                If Val(FixNulos(rsBuff(iContador + 1), vbDouble)) > 0 Then
                    ssDBGrid.Columns(iContador).TagVariant = Val(FixNulos(ssDBGrid.Columns(iContador).TagVariant, vbDouble)) + Val(FixNulos(rsBuff(iContador + 1), vbDouble))
                End If
            End If

        Next

        NVEZ = False
        ssDBGrid.AddItem temp
        rsBuff.MoveNext
        X = X + 1
    Loop

    ssDBGrid.AllowDragDrop = True
    ssDBGrid.RowHeight = 300 ' SSDBGrid.RowHeight * 1.25
    ssDBGrid.Refresh

    ssDBGrid.Redraw = True
    Set rsBuff.RefObject = Nothing
    Set rsBuff = Nothing

End Sub

Public Sub SSDBGridTOTALES(ByRef ssDBGrid As Object)  'SSDataWidgets_B.SSDBGrid)

    On Error Resume Next

    Dim iContador As Long

    Dim temp      As String

    Dim X%

    Dim y%

    ssDBGrid.Redraw = False
    temp = ""

    For y = 0 To ssDBGrid.Cols - 1
        temp = temp & FixNulos(ssDBGrid.Columns(y).TagVariant, vbstring)

        If y < ssDBGrid.Cols - 1 Then
            temp = temp & "~"
        End If

    Next

    ssDBGrid.AddItem temp
    ssDBGrid.Refresh

    If ssDBGrid.Rows > 1 Then
        ssDBGrid.Row = 0
    End If

    ssDBGrid.Redraw = True

End Sub

Public Sub SSDBGridSetGrid(ByRef ssDBGrid As Object)

    Dim i As Long

    Dim n As Long
    
    ssDBGrid.col = 0
    ssDBGrid.SplitterPos = 0
    ssDBGrid.SplitterVisible = False
    ssDBGrid.RemoveAll
    ssDBGrid.Refresh
    ssDBGrid.Redraw = False
    n = ssDBGrid.Cols

    If Not IsEmpty(ssDBGrid.TagVariant) Then
        If n > ssDBGrid.TagVariant Then

            For i = n To ssDBGrid.TagVariant + 1 Step -1
                ssDBGrid.Columns.Remove ssDBGrid.Cols - 1
            Next

        End If
    End If

    ssDBGrid.Redraw = True
    ssDBGrid.Refresh
End Sub

Public Sub SSDBGridSetGrid0(ByRef ssDBGrid As Object)
    ssDBGrid.TagVariant = ssDBGrid.Cols
End Sub

Public Function Ancho_Columnas(ByVal fform As Form, _
                               ByVal dcontainer As Object, _
                               ByVal scadena As String)

    Dim xPos  As Integer

    Dim xPos1 As Integer

    Dim i     As Integer

    xPos = 1
    xPos1 = 1
    i = 0

    Dim a As Integer

    While InStr(xPos1, scadena, ",") > 0

        xPos1 = InStr(xPos, scadena, ",") + 1
        dcontainer.Columns(i).Width = (CInt(Mid(scadena, xPos, xPos1 - xPos - 1)) * 100) + 50
        xPos = xPos1
        i = i + 1

    Wend

End Function

Public Sub DActivaControles(ByVal fform As Form, _
                            ByVal TipOpe As Variant, _
                            ByVal Scontroles As String)

    Dim xEnabled   As Boolean

    Dim xbackColor As Variant

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

Public Sub ReCarga_DBCombos(ByVal fform As Form, ByRef rs As ADODB.Recordset)

    With fform
        .dbcboCategoria.ListField = "NOMBRE"
        .dbcboCategoria.BoundColumn = "CODCAT"
        .dbcboCategoria.BoundText = rs!CATEGORIA

        .dbcboUnidad.ListField = "NOMBRE"
        .dbcboUnidad.BoundColumn = "CODUNI"
        .dbcboUnidad.BoundText = rs!UNIDAD
    End With

End Sub

Public Function Maximo(ByVal stabla As String, _
                       ByVal sCampo As String, _
                       ByVal scondi As String, _
                       ByVal conn As ADODB.Connection, _
                       ByVal stipo As String, _
                       ByVal ilargo As Integer)

    Dim RS1    As ADODB.Recordset

    Dim sQuery As String

    Dim a      As Variant

    Dim b      As Variant

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

    Dim i As Long

    Ceros = scadena

    If iLen < 2 Then Exit Function

    For i = 1 To iLen - 1
        Ceros = schar & Ceros
    Next i

End Function

Public Function Maximo1(ByVal stabla As String, _
                        ByVal sCampo As String, _
                        ByVal scondi As String, _
                        ByVal conn As ADODB.Connection, _
                        ByVal scampo1 As String, _
                        ByVal ilargo As Integer)

    Dim RS1    As ADODB.Recordset

    Dim sQuery As String

    Dim a      As String

    Dim b      As String

    Dim c      As String

    Set RS1 = New ADODB.Recordset
    RS1.ActiveConnection = cn
    RS1.CursorType = adOpenStatic

    If scondi = "" Then
        scondi = "1<2"
    End If

    sQuery = "select  max(" & scampo1 & "),max(" & sCampo & ") from " & stabla & " where " & scondi
    RS1.Open sQuery
    a = LTrim(RTrim(RS1(0)))
    b = ilargo - Len(a)
    c = LTrim(RTrim(Str(RS1(1) + 1)))

    If IsNull(a) Then
        a = a & Ceros("1", ilargo, "0")
    Else
        a = a & LTrim(RTrim(Ceros(c, b, "0")))
    End If

    Maximo1 = a
    Set RS1 = Nothing
End Function

Public Sub Carga_Categorias(ByVal fform As Form, _
                            ByVal Datag As Object, _
                            ByRef rs As ADODB.Recordset)

    Dim sQuery As String

    sQuery = "SELECT COD_MOTATR AS CODIGO,DES_MOTATR AS DESCRIPCION FROM TG_MOTATR"
    'Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
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

Function StrZero1(nDato As Variant, nZeros As Integer)

    Dim wdato As String, wAncho As Integer, wDatoOk As String

    Dim tdato As Variant

    Dim i     As Integer

    If TypeName(nDato) = "String" Then
        If nDato = "" Then
            StrZero1 = ""

            Exit Function

        Else
            tdato = Val(nDato)
        End If

    Else
        tdato = nDato
    End If

    wdato = Trim(Str(tdato))
    wAncho = Len(wdato)

    If wAncho < nZeros Then

        For i = 1 To nZeros - wAncho
            wDatoOk = wDatoOk + "0"
        Next i

        wDatoOk = wDatoOk + wdato
    Else
        wDatoOk = wdato
    End If

    StrZero1 = wDatoOk
End Function

Public Function Rpad(texto As Variant, ByVal iMaxLen As Long) As String

End Function

Public Function LPad(texto As Variant, ByVal iMaxLen As Long) As String

End Function

Public Sub errores(sCodigo As Long)

    Dim oCode    As MESSAGECODE

    Dim oMessage As clsMessages

    Dim sMess    As String

    Dim iPos     As Integer

    Select Case sCodigo

        Case "9999"
            oCode = kMESSAGE_ERR_CODIGO_YA_REGISTRADO
            Set oMessage = New clsMessages
            oMessage.Codigo = oCode
            
            Call oMessage.ShowMesage(iLanguage)

        Case Else
            sMess = Err.Description
            iPos = InStr(1, sMess, "SERVER]", 1)

            If iPos > 0 Then
                sMess = Mid(sMess, iPos + 7)
            End If

            oCode = kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO
            Set oMessage = New clsMessages
            oMessage.Codigo = oCode
            oMessage.OptionalText = Err.Description
            
            Call oMessage.ShowMesage(iLanguage)

            'Aviso "Ha ocurrido un imprevisto !!!  " & Chr(13) & _
            'Chr(13) & "El mensaje de Error es : " & Err.Description & _
            'Chr(13) & "El Nro. de Error es : " & Err.Number, 1
    End Select

    Set oMessage = Nothing
End Sub

Public Function get_botones1(ByVal f As Form, _
                             ByVal Vcod_perfil As Variant, _
                             ByVal vcod_empresa As Variant, _
                             ByVal fname As Variant)

    Dim RS1 As ADODB.Recordset

    Set RS1 = New ADODB.Recordset

    Dim sQuery As String

    sQuery = "Sp_funciones3 '" & Vcod_perfil & "','" & vcod_empresa & "','" & fname & "'"

    If RTrim(cSEGURIDAD) = "" Then
        cSEGURIDAD = DSN_Seguridad
    End If

    RS1.ActiveConnection = cSEGURIDAD
    RS1.CursorType = adOpenStatic
    RS1.Open sQuery

    Dim Scad  As String

    Dim ICONT As Integer

    If Not (RS1.BOF And RS1.EOF) Then
        'For j = 1 To rs1.RecordCount
        ICONT = 1

        While Not RS1.EOF

            If ICONT = 1 Then
                Scad = LTrim(RTrim(RS1!nom_corto))
            Else
                Scad = Scad + "/" + LTrim(RTrim(RS1!nom_corto))
            End If

            RS1.MoveNext
            ICONT = ICONT + 1

        Wend

        'Boton_Enabled rs1!nom_corto, f
    End If

    get_botones1 = Scad
End Function

Public Sub IdiomaEtiquetas(ByVal oForm As Object)

    On Error GoTo hand

    Dim ctl As Control

    iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

    If iLanguage <> "1" Then
        oForm.Caption = oForm.Tag
    End If

    'For Each ctl In oForm.Controls
    '    If Not TypeOf ctl Is TextBox And Not TypeOf ctl Is FunctionsButtons.FunctButt _
    '     And Not TypeOf ctl Is Mantenimientos.MantFunc And Not TypeOf ctl Is ComboBox _
    '     And Not TypeOf ctl Is DataCombo _
    '     And Not TypeOf ctl Is DTPicker And Not TypeOf ctl Is MaskEdBox And Not TypeOf ctl Is LinkLabel _
    '     And Not TypeOf ctl Is MSFlexGridLib.MSFlexGrid And Not TypeOf ctl Is MDIExtend _
    '     And Not TypeOf ctl Is SSDataWidgets_B.ssDBCombo _
    '     And Not TypeOf ctl Is SSDataWidgets_B.SSDBDropDown _
    '     And Not TypeOf ctl Is SSDataWidgets_B.ssDBGrid _
    '     And Not TypeOf ctl Is Shape _
    '     And Not TypeOf ctl Is ListBox _
    '     And Not TypeOf ctl Is PictureBox _
    '     And Not TypeOf ctl Is Line _
    '     And Not TypeOf ctl Is Image _
    '     Then
    '
    '        If iLanguage <> "1" Then
    '            If RTrim(ctl.Tag) <> "" Then
    '                ctl.Caption = ctl.Tag
    '            End If
    '        End If
    '    Else
    '        If TypeOf ctl Is FunctionsButtons.FunctButt _
    '            Or TypeOf ctl Is Mantenimientos.MantFunc Then
    '            ctl.Language = iLanguage
    '        End If
    '    End If
    'Next ctl
    For Each ctl In oForm.Controls

        If TypeOf ctl Is Label Or TypeOf ctl Is Frame Or TypeOf ctl Is OptionButton Or TypeOf ctl Is CommandButton Then
            If iLanguage <> "1" Then
                If RTrim(ctl.Tag) <> "" Then
                    ctl.Caption = ctl.Tag
                End If
            End If

        Else

            If TypeOf ctl Is FunctionsButtons.FunctButt Or TypeOf ctl Is Mantenimientos.MantFunc Then
                ctl.Language = iLanguage
            End If
        End If

    Next ctl

    Exit Sub

hand:
    ErrorHandler Err, "IdiomaEtiquetas"
End Sub

Public Sub IdiomaEtiquetas1(ByVal oForm As Object)

    Dim ctl As Control

    'LoadConnectEmpresa ""
    iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

    If iLanguage <> "1" Then
        oForm.Caption = oForm.Tag
    End If

    For Each ctl In oForm.Controls

        If Not TypeOf ctl Is TextBox And Not TypeOf ctl Is FunctionsButtons.FunctButt And Not TypeOf ctl Is Mantenimientos.MantFunc And Not TypeOf ctl Is ComboBox And Not TypeOf ctl Is DataCombo And Not TypeOf ctl Is Image And Not TypeOf ctl Is CommonDialog And Not TypeOf ctl Is ImageList And Not TypeOf ctl Is Toolbar And Not TypeOf ctl Is StatusBar And Not TypeOf ctl Is Menu And Not TypeOf ctl Is Timer And Not TypeOf ctl Is MDIExtend And Not TypeOf ctl Is Line And Not TypeOf ctl Is SSActiveToolBars Then

            If iLanguage <> "1" Then
                ctl.Caption = ctl.Tag
            End If

        Else

            If TypeOf ctl Is FunctionsButtons.FunctButt Or TypeOf ctl Is Mantenimientos.MantFunc Then
                ctl.Language = iLanguage
            End If
        End If

    Next ctl

End Sub

Sub Centrar_form(ByRef Formulario As Form)

    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (6945 - Formulario.Height) / 2
End Sub

Public Sub ComboBoxToComboBox(ByRef lstOrigen As Object, _
                              ByRef lstDestino As Object, _
                              ByVal iModal As Integer)

    Dim i As Long

    Dim j As Long

    If iModal = 0 Then
        If lstOrigen.ListIndex <> -1 Then
            lstDestino.AddItem ""

            For i = 0 To 0
                lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(lstOrigen.ListIndex)
            Next

            lstOrigen.RemoveItem lstOrigen.ListIndex

            '    For j = 0 To lstOrigen.ListCount - 1
            '        If RTrim(lstOrigen.List(j)) <> "" Then
            '            If lstOrigen.Selected(j) = True Then
            '                lstDestino.AddItem ""
            '                For i = 0 To 0
            '                    lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(j)
            '                Next
            '            End If
            '
            '        End If
            '    Next
            '
            '    For j = lstOrigen.ListCount - 1 To 0 Step -1
            '        If lstOrigen.Selected(j) = True Then
            '            lstOrigen.RemoveItem j
            '        End If
            '    Next
    
        End If

    Else

        For j = 0 To lstOrigen.ListCount - 1

            If RTrim(lstOrigen.List(j)) <> "" Then
                lstDestino.AddItem ""

                For i = 0 To 0
                    lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(j)
                Next

            End If

        Next
    
        For j = lstOrigen.ListCount - 1 To 0 Step -1
            lstOrigen.RemoveItem j
        Next
    
    End If

End Sub

Public Function ExecuteCommandSQL(ByVal Connect As String, ByVal sql As String) As Long
  
    On Error GoTo errorx

    Dim oCn As Object

    Set oCn = CreateObject("ADODB.Connection")
    oCn.Open Connect
    oCn.CommandTimeout = 900
  
    oCn.Execute sql, ExecuteCommandSQL, 128
    oCn.Close
    Set oCn = Nothing

    Exit Function

errorx:

    ExecuteCommandSQL = -2
    Err.Raise Err.Number, "ExecuteCommandSQL", Err.Description
End Function

Function ExisteCampo(pCampo As String, _
                     pTabla As String, _
                     pValor As Variant, _
                     Conexion As String, _
                     Optional pEsStringValor As Boolean = True) As Boolean

    On Error GoTo hand

    If pEsStringValor Then
        If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = '" & pValor & "'", Conexion) > 0 Then
            ExisteCampo = True
        Else
            ExisteCampo = False
        End If

    Else

        If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = " & pValor, Conexion) > 0 Then
            ExisteCampo = True
        Else
            ExisteCampo = False
        End If
    End If

    Exit Function

hand:
    ErrorHandler Err, "ExisteCampo"
    ExisteCampo = False
End Function

Public Function VBsprintf(ByRef InString As String, ParamArray aInValues()) As String

    Dim OutString   As String

    Dim ThisChar    As String

    Dim IndexString As Integer

    Dim IndexValues As Integer

    Dim iNotchar    As Integer

    Dim vValor      As Variant

    Dim strCadena   As String

    OutString = ""
    IndexValues = 0

    For IndexString = 1 To Len(InString)
        ThisChar = Mid(InString, IndexString, 1)
 
        If ThisChar <> "$" Then
            OutString = OutString & ThisChar
        Else

            If VarType(aInValues(IndexValues)) = vbstring Then
                vValor = aInValues(IndexValues)

                If Len(vValor) >= 2 Then
                    If Mid(vValor, 1, 1) <> Chr(39) Then
                        vValor = NotChar(vValor)
                    End If
           
                    strCadena = Mid(vValor, 2, Len(vValor) - 2)

                    If InStr(strCadena, Chr(34)) Or InStr(strCadena, Chr(39)) Then
                        strCadena = NotChar(strCadena)
                        vValor = Chr(39) & strCadena & Chr(39)
                    End If
                End If

            Else
                vValor = CStr(aInValues(IndexValues))
                vValor = NotChar(vValor)
            End If
   
            OutString = OutString + vValor
            IndexValues = IndexValues + 1
        End If

    Next

    VBsprintf = OutString

End Function

Private Function NotChar(ByVal vValor As String) As String

    Dim i       As Integer

    Dim sReturn As String

    If InStr(vValor, Chr(34)) Or InStr(vValor, Chr(39)) Then

        For i = 1 To Len(vValor)

            If Asc(Mid(vValor, i, 1)) <> 39 And Asc(Mid(vValor, i, 1)) <> 34 Then
                sReturn = sReturn + Mid(vValor, i, 1)
            End If

        Next

    Else
        sReturn = vValor
    End If

    NotChar = sReturn
End Function

Public Sub BuscarComboD(MyCombo As Object, MyKey)

    On Error Resume Next

    MyCombo.ListIndex = -1

    If MyCombo.ListCount > 0 Then
        If RTrim(MyKey) <> "" Then
            MyCombo.value = MyKey
        End If
    End If

End Sub

Public Sub ComboBoxToComboBox1(oSource As Object, _
                               oTarget As Object, _
                               Optional bAll As Boolean = False)

    Dim i     As Integer

    Dim j     As Integer

    Dim ix    As Integer

    Dim iRows As Integer

    Dim iCols As Integer

    iRows = oSource.ListCount - 1
    iCols = oSource.ColumnCount - 1

    If oTarget.ColumnCount <= iCols Then
        oTarget.ColumnCount = iCols + 1
    End If

    oTarget.ColumnWidths = oSource.ColumnWidths
    ix = oTarget.ListCount
    'otarget.

    For i = 0 To iRows 'To Step -1

        If (oSource.Selected(i) = True) Or (bAll = True) Then
            oTarget.AddItem '""

            For j = 0 To iCols
                'oTarget.Column(j, ix) = oSource.Column(j, i)
                oTarget.Column(j, oTarget.ListCount - 1) = "" & oSource.Column(j, i)
            Next j

            ix = ix + 1
  
        End If

    Next i

    For i = iRows To 0 Step -1

        If (oSource.Selected(i) = True) Or (bAll = True) Then
            oSource.RemoveItem (i)
        End If

    Next

End Sub

Public Function ASearch(avArray As Variant, _
                        vSearchFor As Variant, _
                        iIndice As Integer, _
                        Optional base As Variant) As Integer
    
    Dim iIndex  As Integer

    Dim iMaxLen As Integer
    
    ' Valor de retorno si no se encuentra el elemento
    ASearch = -1
    
    iMaxLen = UBound(avArray, 2)
    
    ' Inicio de busqueda del elemento
    For iIndex = 0 To iMaxLen
    
        If avArray(iIndice, iIndex) = vSearchFor Then
        
            ASearch = iIndex

            Exit Function
        
        End If
        
    Next

End Function

Public Function ASearchNew(avArray As Variant, _
                           vSearchFor As Variant, _
                           iIndice As Integer, _
                           Optional base As Variant) As Integer
                        
    ' Control de Parametro opcional
    
    Dim iIndex  As Integer

    Dim iMaxLen As Integer
    
    ' Valor de retorno si no se encuentra el elemento
    ASearchNew = -1
    
    iMaxLen = UBound(avArray, 1)
    
    ' Inicio de busqueda del elemento
    For iIndex = 0 To iMaxLen
    
        If avArray(iIndex, iIndice) = vSearchFor Then
        
            ASearchNew = iIndex

            Exit Function
        
        End If
        
    Next

End Function

'Public Function EjecBatch(Fichero$) As Boolean
'    Dim valor As Long
'    Dim Comienzo As STARTUPINFO
'    Dim Proceso As PROCESS_INFORMATION
'    Comienzo.cb = Len(Comienzo)
'    valor = CreateProcessA(0&, Fichero$, 0&, 0&, 1&, &H20&, 0&, 0&, Comienzo, Proceso)
'    valor = WaitForSingleObject(Proceso.hProcess, -1&)
'    If valor = -1 Then
'        EjecBatch = False
'        MsgBox "El proceso " & Fichero$ & " no ha podido ser lanzado con éxito." & vbCrLf & "Asegúrese que el programa existe o el path es correcto.", 16, "Ejecución Batch"
'    Else
'        valor = CloseHandle(Proceso.hProcess)
'        EjecBatch = True
'    End If
'End Function

Public Function Redondear(dblnToR As Double, Optional intCntDec As Integer) As Double

    Dim dblPot As Double

    Dim dblF   As Double

    If dblnToR < 0 Then
        dblF = -0.5
    Else
        dblF = 0.5
    End If

    dblPot = 10 ^ intCntDec
    Redondear = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot
End Function

Public Sub SelectionText(cltSel As Object)
    cltSel.SelStart = 0
    cltSel.SelLength = Len(cltSel.Text)
End Sub

Public Sub EjecutaOpcionMenu(ByVal sNameOpcion As String, _
                             perfil As String, _
                             empresa As String)

    On Error GoTo EjecutaOpcionMenu

    Dim rutexe     As String

    Dim nomfor     As String

    Dim cod_padre  As String

    Dim des_opcion As String

    Dim tDllName   As String

    Dim sopcion    As String

    Dim lValDev    As Long

    Dim nivel      As Integer

    Dim Tipo       As String

    Dim icono      As String

    Dim ReadmeDoc  As String

    Dim oWordDoc   As Word.Document

    Dim oWordApp   As Word.Application

    Dim strSql     As String

    On Error GoTo EjecutaOpcionMenu

    Get_Datos_form sNameOpcion, rutexe, nomfor, nivel, Tipo, icono, cod_padre, des_opcion, ReadmeDoc

    sopcion = Tipo
    
    tDllName = rutexe
    
    If sopcion = "C" Or sopcion = "P" Or sopcion = "M" Or sopcion = "N" Then
        If sDllName <> tDllName Then
            If Not oFormObjDLL Is Nothing Then
                Set oFormObjDLL = Nothing
            End If

            If Not objFormDLL Is Nothing Then
                Set objFormDLL = Nothing
            End If

            sDllName = tDllName

            If sopcion <> "N" Then
                Set objFormDLL = CreateObject(sDllName & ".clsForm")
                'Else
                'Set objFormDLL = New InteropUserControlLibraryPrueba1.clsForm
                'Set objFormDLL = New clsform2     ' InteropUserControlLibraryPrueba1_Interop_InteropForm9    '.clsForm      'CreateObject(sDllName & "." & sDllName & "." & nomfor)
            End If
        End If
        
        If sopcion = "N" Then
            Set objFormDLL.Parent = MDIPrincipal
        End If
        
        Set oFormObjDLL = objFormDLL.GetForm(nomfor)

        If Not (oFormObjDLL Is Nothing) Then
            objFormDLL.Cod_Empresa = empresa
            objFormDLL.UserName = vusu
            objFormDLL.Cod_Perfil = perfil
            objFormDLL.Rutas = App.Path
            objFormDLL.Cod_Opcion = sNameOpcion
            objFormDLL.ConnectEmpresa = DSN_Empresa
            objFormDLL.ConnectSeguridad = DSN_Seguridad
            objFormDLL.Language = iLanguage
            
            'Set oFormObjDLL2 = objFormDLL.GetForm("FRM_TOOLBAR")
            
            'Dim iTemp As Integer
            'iTemp = 1
            'oFormObjDLL2.CambiarContenedor iTemp, oFormObjDLL

            If sopcion <> "N" Then
            
                objFormDLL.Parent = MDIPrincipal
             
            End If

            On Error GoTo EjecutaOpcionMenu

            If sopcion = "M" Or sopcion = "N" Then
                oFormObjDLL.Show
            Else
                oFormObjDLL.Show vbModal
                
            End If

           ' MDIPrincipal.RefreshWindowList
            Set oFormObjDLL = Nothing
        End If

    Else

        If sopcion = "D" And RTrim(ReadmeDoc) <> "" Then
            Set oWordApp = New Word.Application
            oWordApp.Documents.Open ReadmeDoc
            oWordApp.Visible = True
            Set oWordApp = Nothing
        End If
    End If

    strSql = "EXEC SEG_REGISTRA_ACCESO_ESTACION '" & vusu & "','" & ComputerName & "','" & vemp & "','','" & RTrim(sNameOpcion) & "'"
    Call ExecuteCommandSQL(DSN_Seguridad, strSql)

    Exit Sub

EjecutaOpcionMenu:
    ErrorHandler Err, "EjecutaOpcionMenu"
    Set oFormObjDLL = Nothing
End Sub

Public Function Get_Datos_form(ByVal sopcion As String, _
                               ByRef rutexe As String, _
                               ByRef nomfor As String, _
                               ByRef nivel As Integer, _
                               ByRef Tipo As String, _
                               ByRef icono As String, _
                               ByRef cod_padre As String, _
                               ByRef des_opcion As String, _
                               ByRef ReadmeDoc As String)

    Dim iCount As Integer

    Dim mRs    As ADODB.Recordset
    
    sQuery = "SELECT isnull(RUTEXE,''),isnull(nomfor,''),isnull(nivel,0),isnull(tipo,''),isnull(icono,''),isnull(cod_padre,''),isnull(des_opcion,'') , ISNULL(ReadmeDoc,'') FROM SEG_OPCIONES  WHERE COD_OPCION='" & sopcion & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = conn
    mRs.CursorType = adOpenStatic
    mRs.Open sQuery
    iCount = mRs.RecordCount

    If iCount > 0 Then
        rutexe = mRs(0)
        nomfor = mRs(1)
        nivel = mRs(2)
        Tipo = mRs(3)
        icono = mRs(4)
        cod_padre = mRs(5)
        des_opcion = mRs(6)
        ReadmeDoc = mRs(7)
    End If

    Set mRs = Nothing
End Function

Sub Informa(ByVal Mens As String, Optional ByVal amensaje As clsMessages)

    If Mens <> "" Then

        Dim rpta As Byte

        rpta = MsgBox(Mens, vbInformation, "Informa")

        Exit Sub

    End If

    amensaje.ShowMesage iLanguage
End Sub

Public Function ComputerName() As String

    Dim KeyName$

    Dim keylen&

    Dim iNull
            
    keylen& = 2000
    KeyName$ = String$(keylen, 0)
    
    GetcomputerName KeyName$, keylen&
    
    iNull = InStr(KeyName, Chr(0))
    ComputerName = Mid(KeyName$, 1, iNull - 1)
End Function

Public Sub Mensaje(ByVal oMESSAGECODE As MESSAGECODE)

    Dim amensaje As clsMessages

    Set amensaje = New clsMessages

    amensaje.Codigo = oMESSAGECODE
    amensaje.ShowMesage iLanguage

End Sub

Public Function GetDataSet(ByVal Connect As String, _
                           ByVal sql As String) As Object 'ADOR.Recordset

    On Error GoTo errorx

    Dim oRs As Object

    Dim oCn As Object
 
    Set oRs = CreateObject("ADODB.Recordset")
    Set oCn = CreateObject("ADODB.Connection")
    oCn.CursorLocation = 3
    oCn.Open Connect
    oCn.CommandTimeout = 900
    oRs.Open sql, oCn, 3, 4
    Set GetDataSet = oRs
    Set GetDataSet.ActiveConnection = oCn
    Set oRs.ActiveConnection = Nothing
    oCn.Close
    Set oCn = Nothing
 
    Exit Function

errorx:
    Err.Raise Err.Number, Err.Source, Err.Description
    MsgBox Err.Description
End Function

Public Function Refresh(ByRef rsData As Object, ByRef vBuffer As Variant) As Variant

    Dim n As Integer

    Dim i As Integer

    If Not rsData Is Nothing Then
        n = rsData.Fields.count - 1
        ReDim vBuffer(n, iMaxEnumField)

        For i = 0 To n
            vBuffer(i, IName) = rsData.Fields(i).Name
            vBuffer(i, iActualSize) = rsData.Fields(i).ActualSize
            vBuffer(i, iAttributes) = rsData.Fields(i).Attributes
            vBuffer(i, iDefinedSize) = rsData.Fields(i).DefinedSize
            vBuffer(i, iNumericScale) = rsData.Fields(i).NumericScale
            vBuffer(i, iOriginalValue) = rsData.Fields(i).value
            vBuffer(i, iPrecision) = rsData.Fields(i).Precision
            vBuffer(i, iType) = rsData.Fields(i).Type
            vBuffer(i, iUnderlyingValue) = rsData.RecordCount
            vBuffer(i, iValue) = rsData.Fields(i).value
        Next i

    End If

End Function

Public Sub Aviso(Mensaje As String, Tipo As Integer)

    Select Case Tipo

        Case 1
            MsgBox Mensaje, vbExclamation, "Aviso"

        Case 2
            MsgBox Mensaje, vbInformation + vbMsgBoxRight, "Mensaje"

        Case 3
            MsgBox Mensaje, vbCritical, "Error Grave"
    End Select

End Sub

Function FixData(wtexto As Variant, ofield As ADODB.FIELD)

    If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then
   
        Select Case ofield.Type

            Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
                wtexto = 0

            Case adBoolean
                wtexto = False

            Case adDate
                wtexto = Empty

            Case adChar, adVarChar
                wtexto = ""
        End Select

    End If

    FixData = wtexto
End Function

Public Function CargarRecordSetDesconectado(ByVal sSQl As String, _
                                            ByVal cCONNECT As String) As ADODB.Recordset

    Dim rsBD     As ADODB.Recordset

    Dim rsGridEx As ADODB.Recordset

    Dim ofield   As Object

    Dim oCon     As ADODB.Connection

    Set oCon = New ADODB.Connection
    
    oCon.CursorLocation = adUseClient
    oCon.Open cCONNECT
    oCon.CommandTimeout = 900
    
    Set rsBD = New ADODB.Recordset
    Set rsBD.ActiveConnection = oCon
     
    rsBD.CursorLocation = adUseClient
    rsBD.CursorType = adOpenStatic
    
    rsBD.Open sSQl

    Set rsGridEx = New ADODB.Recordset
    rsGridEx.CursorLocation = adUseClient
    Set rsGridEx.ActiveConnection = Nothing

    For Each ofield In rsBD.Fields

        rsGridEx.Fields.Append ofield.Name, ofield.Type, ofield.DefinedSize, adFldIsNullable
        rsGridEx.Fields(ofield.Name).NumericScale = rsBD.Fields(ofield.Name).NumericScale
        rsGridEx.Fields(ofield.Name).DefinedSize = rsBD.Fields(ofield.Name).DefinedSize
        rsGridEx.Fields(ofield.Name).Precision = rsBD.Fields(ofield.Name).Precision
    Next

    rsGridEx.Open
           
    If rsBD.RecordCount Then
        rsBD.MoveFirst

        Do While Not rsBD.EOF
            rsGridEx.AddNew

            For Each ofield In rsBD.Fields

                rsGridEx.Fields(ofield.Name).value = FixData(rsBD.Fields(ofield.Name).value, rsBD.Fields(ofield.Name))
            Next

            rsGridEx.Update
            rsBD.MoveNext
        Loop

    End If

    Set CargarRecordSetDesconectado = rsGridEx
    
End Function

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, _
                                 ByVal iFixsCols As Integer, _
                                 ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If
    
    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If
    
End Function

