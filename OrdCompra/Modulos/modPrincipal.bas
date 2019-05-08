Attribute VB_Name = "modPrincipal"
Option Explicit
'Option Base 0
Dim sQuery As String


Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Const CB_FINDSTRING = &H14C
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const BlockSize = 100000 'This size can be experimented with for
                         'performance and reliability.


Global Ruta_Logo_Empresa As String
Global Num_Ruc_Empresa As String
Global Direccion_Empresa As String
Global DSN_Empresa As String
Global DSN_Seguridad As String
Global Ruta0_Empresa As String
Global Fecha_Hora_Conexion As String

Public vusu As String
Public vemp As String
Public vper As String
Public vemp1 As String
Public vpas As String
Public vNomFor As String
Public vRuta As String

Public iLanguage As Integer
Public oMDIParent As Object

Public cn As ADODB.Connection
Public sDllName As String
Public oFormObjDLL As Object
Public objFormDLL As Object

' Constantes de la Aplicación
Public Const M_NUEVO As String = "NUEVO"
Public Const M_EDITAR As String = "EDITAR"
Public Const M_ELIMINAR As String = "ELIMINAR"

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
Public DetPedido As RegDetallePedido, cServidor As String

Public cConnect As String
Public cSEGURIDAD As String

Public bCargaConexion As Boolean

Public Const cCLASS_TG_PURORD As String = "Visuales.clsTG_PurOrd"
    
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Declare Function GetcomputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public Const kCHR_BOLD_IN As String = "E"
Public Const kCHR_BOLD_OUT As String = "F"

Public Const kKEY_SEPARATOR As String = "-+-"
'Estructura de cada elemento del arreglo para los mensajes
Public Type Message
    Tipo As TypeMsg
    Code As CodeMsg
    Description As String
    HelpID As Integer
    Tag As Variant
End Type
'Ubound del arreglo (mayor codemsg  -1)
Public Const kMESSAGE_COUNT = 123
Public aMessage(kMESSAGE_COUNT) As Message

'Constantes para Tipo de Mensajes
Public Enum TypeMsg
    kTYPEMSG_INFORMATION = 1
    kTYPEMSG_QUESTION = 2
    kTYPEMSG_WARNING = 3
    kTYPEMSG_ERROR = 4
    kTYPEMSG_FIELD = 5
End Enum

'Constantes publicas para Codigo de Mensajes
Public Enum CodeMsg
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

Enum FIELDx
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
Public Const kMESSAGE_PROVEEDOR_PASADOS As String = ""
Public Const kMESSAGE_PROVEEDOR_PENDIENTES As String = ""



Public Sub InitMessages()
    Dim iLenarray As Integer
    
    iLenarray = 0
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_NOTEMPTY
    aMessage(iLenarray).Description = "Es indispensable completar información"
    'aMessage(iLenarray).Description2 = "Information Incomplete"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
    iLenarray = 1
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_FOUND
    aMessage(iLenarray).Description = "Información ya se encuetra registrada"
    aMessage(iLenarray).Tag = 1
    
    iLenarray = 2
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_FILE_NOT_FOUND
    aMessage(iLenarray).Description = "La información no ha sido registrada"
'    aMessage(iLenarray).Description2 = "Information not valid"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
'    iLenarray = 3
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_WAIT
'    aMessage(iLenarray).Description = "Espere un momento."
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
    
    iLenarray = 4
    aMessage(iLenarray).Tipo = kTYPEMSG_WARNING
    aMessage(iLenarray).Code = kMESSAGE_WAR_ENABLED_DELETED
    aMessage(iLenarray).Description = "Confirma eliminación de información" 'Eliminación del Objeto
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
    iLenarray = 5
    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
    aMessage(iLenarray).Code = kMESSAGE_ASK_PRINT_FILE
    aMessage(iLenarray).Description = "Desea Imprimir Archivo"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 6
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_USERCONNECTFAIL
    aMessage(iLenarray).Description = "No se pudo Abrir la Conección " '
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 7
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_USERCONNECTFAIL
    aMessage(iLenarray).Description = "Empresa - Usuario o Contraseña no Permitida. Verifique Datos "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 8
    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
    aMessage(iLenarray).Code = kMESSAGE_ASK_EXIT_SYSTEM
    aMessage(iLenarray).Description = "Desea Salir del Sistema"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

'    iLenarray = 9
'    aMessage(iLenarray).Tipo = kTYPEMSG_FIELD
'    aMessage(iLenarray).Code = KMESSAGE_FLD_EXIT_RANGE
'    aMessage(iLenarray).Description = "Campo '$' : Es de tipo '$' Debe estar comprendido entre '$' y '$' "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1

    iLenarray = 10
    aMessage(iLenarray).Tipo = kTYPEMSG_WARNING
    aMessage(iLenarray).Code = kMESSAGE_WAR_CONFIR_CHANGES
    aMessage(iLenarray).Description = "Confirma los cambios efectuados?"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

'    iLenarray = 11
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INCORRECT_VALIDFIELD
'    aMessage(iLenarray).Description = "Campo No existe en Getproperties: "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1

    iLenarray = 12
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_CODIGO_YA_REGISTRADO
    aMessage(iLenarray).Description = "El Código ya ha sido registrado"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 13
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO
    aMessage(iLenarray).Description = "Ha ocurrido un imprevisto !!! "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 14
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_REGISTRO_TIENE_TRANSAC_RELACIONADAS
    aMessage(iLenarray).Description = "Operación no valida.  El Registro tiene transacciones relacionadas"
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 15
    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
    aMessage(iLenarray).Code = kMESSAGE_INF_FILE_PRINT_OK
    aMessage(iLenarray).Description = "Archivo ha sido Enviado a Impresión "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1

    iLenarray = 16
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_FILE_NOT_FOUND
    aMessage(iLenarray).Description = "Archivo No encontrado "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
    
'    iLenarray = 17
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_BUSY_SERVER
'    aMessage(iLenarray).Description = "Servidor está ocupado"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
        
    
'    iLenarray = 18
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_CANCEL_PROCESS
'    aMessage(iLenarray).Description = "Desea Cancelar el Proceso ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
    
    iLenarray = 19
    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
    aMessage(iLenarray).Code = kMESSAGE_INF_PROCESS_SATISFACTO
    aMessage(iLenarray).Description = "El proceso concluyó satisfactoriamente "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
    iLenarray = 20
    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
    aMessage(iLenarray).Code = kMESSAGE_ERR_PROCESS_INSATISFACT
    aMessage(iLenarray).Description = "El proceso No concluyó satisfactoriamente "
    aMessage(iLenarray).HelpID = 1
    aMessage(iLenarray).Tag = 1
    
'    iLenarray = 21
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INCORRECT_TAG_SET
'    aMessage(iLenarray).Description = "Propiedad Tag del control a Validar es Incorrecta"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1

    
'    iLenarray = 22
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_MISSING_TYPEOF
'    aMessage(iLenarray).Description = "Clase de Control en typeof no está definida en User Sub/Function"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 23
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_PROCESS
'    aMessage(iLenarray).Description = "Desea procesar ahora ... ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 24
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_HOUR_EQUAL
'    aMessage(iLenarray).Description = "Horas Inicio y Final deben de ser diferentes"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 25
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_HOUR_FAIL
'    aMessage(iLenarray).Description = "Hora Inicio no es válida"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 26
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_DATE_FAIL
'    aMessage(iLenarray).Description = "Fecha Fin debe ser mayor a Fecha Inicio "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 27
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_NO_INIT_SEARCH
'    aMessage(iLenarray).Description = "No se admite criterio de Búsqueda. Cantidad de caracteres requeridos : "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 28
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INVALID_OPER
'    aMessage(iLenarray).Description = "Operación no permitida cuando Estado es igual a "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 29
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_CONFIR_COPY
'    aMessage(iLenarray).Description = "Desea que se Efectue la Copia  ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 30
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INVALID_HOUR
'    aMessage(iLenarray).Description = "Hora invalida, por favor verifique. . . "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 31
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_CONFIR_PROCESS
'    aMessage(iLenarray).Description = "Confirma Proceso ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 32
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INVALID_SELECC
'    aMessage(iLenarray).Description = "Selección no válida. Vuelva a elegir"
'    aMessage(iLenarray).Description2 = "Invalid Selección. Choice Please "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 33
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_DATA_NOTFOUND
'    aMessage(iLenarray).Description = "No se encontro información "
'    aMessage(iLenarray).Description2 = "Data not Found"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 34
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_NEW_OPEESP
'    aMessage(iLenarray).Description = "No se encontró coincidencias . Desea crear Nueva"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 35
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INVALID_TARGET
'    aMessage(iLenarray).Description = "Destino es Invalido. Por favor Verifique . . ."
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 36
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_INTERNAL_ONLY
'    aMessage(iLenarray).Description = "Soló es valido entre Procesos Internos."
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 37
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_ERR_COMPUESTO
'    aMessage(iLenarray).Description = "Operación no permitida cuando Estado $ es $"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 38
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_SELEC_NOTFOUND
'    aMessage(iLenarray).Description = "No se encontro información.  Vuelva a elegir. "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 39
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_CONFIR_DEPURA
'    aMessage(iLenarray).Description = "Desea proceder a la Depuración . . ."
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 40
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_PUT_END
'    aMessage(iLenarray).Description = "Desea que se Adicione al final de la selección ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 41
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_GENERAL
'    aMessage(iLenarray).Description = ""
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 42
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_FILE_PRINT_OK
'    aMessage(iLenarray).Description = "Impresion ha concluido satisfactoriamente...!"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 43
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_OPERAC_NOT_IN_PROCESS
'    aMessage(iLenarray).Description = "PROCESO DE LA OPERACION NO SE ENCUENTRA VIGENTE"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 44
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_NOT_EXIST_LECTURA
'    aMessage(iLenarray).Description = "No Existe lectura para la fecha especificada"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 45
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_NOT_RIGHT_OPTION
'    aMessage(iLenarray).Description = "Usted no tiene Acceso a esta Transacción"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 46
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_APPLY_CHANGE_DETAILS
'    aMessage(iLenarray).Description = "Desea Aplicar Cambios a los detalles existentes ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 47
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_MAILING_FILE
'    aMessage(iLenarray).Description = "Desea enviar Impresión por email ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 48
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_TYPE_YN
'    aMessage(iLenarray).Description = "El campo debe ser S ó N "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 49
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_NEW_CODIGO
'    aMessage(iLenarray).Description = "!!! Nuevo Codigo Generado !!!"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 50
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_APPLY_CHANGE_ESTILO
'    aMessage(iLenarray).Description = "Desea Actualizar el Estilo"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 51
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_APPLY_CHANGE_ESTILO
'    aMessage(iLenarray).Description = "Desea aplicar los cambios a todo el Pedido"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'
'    iLenarray = 52
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_WAR_ENABLED_CHANGES_IMAGES
'    aMessage(iLenarray).Description = "Está seguro de Cambiar el Gráfico Actual. Este cambio es de caracter irreversible"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 53
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_DELETE_NOT_PROCESS
'    aMessage(iLenarray).Description = "Eliminación no procede"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 54
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_PRINT_CANTIDAD_SOLICITADA
'    aMessage(iLenarray).Description = "Incluyo en Impresión Cantidades Solicitadas (Consumo Neto) ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 55
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_UPDATE_ALL_TALLAS_MOVIM
'    aMessage(iLenarray).Description = "Desea Actualizar todos las Tallas de Este Item con la Nueva Cantidad ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 56
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_EXPORT_EXCEL
'    aMessage(iLenarray).Description = "Desea enviar Filas Seleccionadas a Microsoft Excel ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 57
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_TRANSFERENCIA
'    aMessage(iLenarray).Description = "Desea Transferir Nuevamente el Movimiento ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 58
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_TRANSDBF
'    aMessage(iLenarray).Description = "Movimiento Transferido, Desea Anular Transferencia ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 59
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_VALIDAOT
'    aMessage(iLenarray).Description = "Desactivar Validacion De Ingreso Al Almacen De Tela Cruda ?"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 60
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_CLIENTE
'    aMessage(iLenarray).Description = "Debe ingresar Abreviado de Cliente"
'    aMessage(iLenarray).Description2 = "Client Abr. doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 61
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_NOM_CLIENTE
'    aMessage(iLenarray).Description = "Debe ingresar Nombre de Cliente"
'    aMessage(iLenarray).Description2 = "Client Name doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'   iLenarray = 62
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_CLIENTE
'    aMessage(iLenarray).Description = "Debe ingresar Descripcion de Cliente "
'    aMessage(iLenarray).Description2 = "Client Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'   iLenarray = 63
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_DIVISION
'    aMessage(iLenarray).Description = "Debe ingresar Nombre de Division"
'    aMessage(iLenarray).Description2 = "Division Name doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'   iLenarray = 64
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_DIVISION
'    aMessage(iLenarray).Description = "Debe ingresar Codigo de Division "
'    aMessage(iLenarray).Description2 = "Division Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'
'   iLenarray = 65
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_COLOR
'    aMessage(iLenarray).Description = "Debe ingresar Nombre de Color"
'    aMessage(iLenarray).Description2 = "Colour Name doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'   iLenarray = 66
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_COLOR
'    aMessage(iLenarray).Description = "Debe ingresar Codigo de Color "
'    aMessage(iLenarray).Description2 = "Colour Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 67
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_ESTCLI
'    aMessage(iLenarray).Description = "Debe ingresar Código de Estilo"
'    aMessage(iLenarray).Description2 = "Style Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 68
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_NOM_ESTCLI
'    aMessage(iLenarray).Description = "Debe ingresar Nombre de Estilo"
'    aMessage(iLenarray).Description2 = "Style Name doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 69
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_TIP_ESTCLI
'    aMessage(iLenarray).Description = "Debe ingresar Tipo Prenda "
'    aMessage(iLenarray).Description2 = "Garment Type doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 70
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_NUEVO_ESTCLI
'    aMessage(iLenarray).Description = "Estilo Cliente NO EXISTE, desea crear uno Nuevo ? "
'    aMessage(iLenarray).Description2 = "Client Style NOT EXIST . Are you sure create new Style ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 71
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_NUEVO_PURORD
'    aMessage(iLenarray).Description = "Purchase Order NO EXISTE, desea crear uno Nuevo ? "
'    aMessage(iLenarray).Description2 = "Purchase Order NOT EXIST . Are you sure create new ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 72
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_COMI
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Comisionista"
'    aMessage(iLenarray).Description2 = "Commissionist Code not allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 73
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_COMI
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Comisionista"
'    aMessage(iLenarray).Description2 = "Commissionist Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 74
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_DESTINO
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Destino"
'    aMessage(iLenarray).Description2 = "Target Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 75
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_DESTINO
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Destino"
'    aMessage(iLenarray).Description2 = "Target Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 76
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 77
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_ABR_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Abreviatura Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Abr. doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 78
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_NOM_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Nombre Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Name doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 79
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 80
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DIR_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Dirección Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Address doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 81
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_TEL_FABRICA
'    aMessage(iLenarray).Description = "Debe ingresar Telefono 1 Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Phone 1 doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 82
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_MONEDA
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Moneda"
'    aMessage(iLenarray).Description2 = "Money Code doesn't allow null values "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 83
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_MONEDA
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Moneda"
'    aMessage(iLenarray).Description2 = "Money Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 84
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_ORGANIZ
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Organización"
'    aMessage(iLenarray).Description2 = "Organization Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 85
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_ORGANIZ
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Organización"
'    aMessage(iLenarray).Description2 = "Organization Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 86
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_PAGO
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Pago de Embarque"
'    aMessage(iLenarray).Description2 = "Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 87
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_PAGO
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Pago de Embarque"
'    aMessage(iLenarray).Description2 = "Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 88
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_FABRES
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Fabrica"
'    aMessage(iLenarray).Description2 = "Factory Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 89
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_CLIRES
'    aMessage(iLenarray).Description = "Debe ingresar Codigo Cliente"
'    aMessage(iLenarray).Description2 = "Client Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 90
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_RES_CLIENTE
'    aMessage(iLenarray).Description = "Debe ingresar Responsable Cliente"
'    aMessage(iLenarray).Description2 = "Response Client doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 91
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_TEMP
'    aMessage(iLenarray).Description = "Debe ingresar Código Temporada"
'    aMessage(iLenarray).Description2 = "Season Code doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 92
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_TEMP
'    aMessage(iLenarray).Description = "Debe ingresar Descripción Temporada"
'    aMessage(iLenarray).Description2 = "Season Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 93
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_DATA_SAVE
'    aMessage(iLenarray).Description = "Los cambios se realizaron satisfactoriametne"
'    aMessage(iLenarray).Description2 = "Changes Updated"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 94
'    aMessage(iLenarray).Tipo = kTYPEMSG_INFORMATION
'    aMessage(iLenarray).Code = KMESSAGE_INF_DATA_DELETE
'    aMessage(iLenarray).Description = "El registro fue eliminado"
'    aMessage(iLenarray).Description2 = "Row Deleted"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 95
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_PORC_CLIENTE
'    aMessage(iLenarray).Description = "El porcentaje debe ser numérico"
'    aMessage(iLenarray).Description2 = "Percent is numeric"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 96
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_TIPEMB
'    aMessage(iLenarray).Description = "Debe ingresar Codigo de Tipo de Embarque"
'    aMessage(iLenarray).Description2 = "Shipment Type Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 97
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_TIPEMB
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Tipo de Embarque"
'    aMessage(iLenarray).Description2 = "Shipment Type Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 98
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_TIPPRE
'    aMessage(iLenarray).Description = "Debe ingresar Código de Tipo de Prenda"
'    aMessage(iLenarray).Description2 = "Garment Type Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 99
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_TIPPRE
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Tipo de Prenda"
'    aMessage(iLenarray).Description2 = "Garment Type Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 100
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_UM
'    aMessage(iLenarray).Description = "Debe ingresar Código de Unidad de Medida"
'    aMessage(iLenarray).Description2 = "U.M. Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 101
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_UM
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Unidad de Medida"
'    aMessage(iLenarray).Description2 = "U.M. Description doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 102
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_ANO_MES
'    aMessage(iLenarray).Description = "No existen Datos"
'    aMessage(iLenarray).Description2 = "There are not valid data"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 103
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_SERIE
'    aMessage(iLenarray).Description = "Debe ingresar la Serie de la Factura"
'    aMessage(iLenarray).Description2 = "Invoice Serie doesn't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 104
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_FACTURA
'    aMessage(iLenarray).Description = "Fecha de Facturacion Invalida"
'    aMessage(iLenarray).Description2 = "Invoice Date is not valid"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 105
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_DELETE_PURORD
'    aMessage(iLenarray).Description = "Confirma eliminación de P.O."
'    aMessage(iLenarray).Description2 = "Are you sure Delete Purchase Order and your dependencies ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 106
'    aMessage(iLenarray).Tipo = kTYPEMSG_QUESTION
'    aMessage(iLenarray).Code = KMESSAGE_ASK_DELETE_LOTEST
'    aMessage(iLenarray).Description = "Confirma eliminación de Lote EStilo"
'    aMessage(iLenarray).Description2 = "Are you sure delete Lote Style and your dependencies ? "
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 107
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_MATPRI
'    aMessage(iLenarray).Description = "Debe ingresar Código de Mat.Prima"
'    aMessage(iLenarray).Description2 = "Mat.Prim. Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 108
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_COMPEST
'    aMessage(iLenarray).Description = "Debe ingresar Código de Comp.Estilo"
'    aMessage(iLenarray).Description2 = "Style Comp. Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 109
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_COMPEST
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Comp.Estilo"
'    aMessage(iLenarray).Description2 = "Style Comp. Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 110
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_GRUTAL
'    aMessage(iLenarray).Description = "Debe ingresar Código de Grupo Tallas"
'    aMessage(iLenarray).Description2 = "Size Group Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 111
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_GRUTAL
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Grupo Tallas"
'    aMessage(iLenarray).Description2 = "Size Group Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 112
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_FAMHIL
'    aMessage(iLenarray).Description = "Debe ingresar Código de Familia Hilado"
'    aMessage(iLenarray).Description2 = "Hil. Family Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 113
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_FAMHIL
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Familia Hilado"
'    aMessage(iLenarray).Description2 = "Hil. Family Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 114
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_HILADO
'    aMessage(iLenarray).Description = "Debe ingresar Código de Hilado"
'    aMessage(iLenarray).Description2 = "Hil. Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 115
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_HILADO
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Hilado"
'    aMessage(iLenarray).Description2 = "Hil. Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 116
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_MATPRI
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Mat. Prima"
'    aMessage(iLenarray).Description2 = "Mat.Prima Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 117
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_PZAEST
'    aMessage(iLenarray).Description = "Debe ingresar Código de Pza. Estilo"
'    aMessage(iLenarray).Description2 = "Piece Style Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 118
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_PZAEST
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Pza. Estilo"
'    aMessage(iLenarray).Description2 = "Piece Style Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 119
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_TALLA
'    aMessage(iLenarray).Description = "Debe ingresar Código de Talla"
'    aMessage(iLenarray).Description2 = "Size Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 120
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_IND_TALLA
'    aMessage(iLenarray).Description = "Debe ingresar Orden Talla"
'    aMessage(iLenarray).Description2 = "Size Order doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 121
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_COD_TIPCOMP
'    aMessage(iLenarray).Description = "Debe ingresar Código de Tipo de Componente"
'    aMessage(iLenarray).Description2 = "Comp. Type Code doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1
'
'    iLenarray = 122
'    aMessage(iLenarray).Tipo = kTYPEMSG_ERROR
'    aMessage(iLenarray).Code = KMESSAGE_ERR_VALIDA_DES_TIPCOMP
'    aMessage(iLenarray).Description = "Debe ingresar Descripción de Tipo de Componente"
'    aMessage(iLenarray).Description2 = "Comp. Type Description doesn 't allow null values"
'    aMessage(iLenarray).HelpID = 1
'    aMessage(iLenarray).Tag = 1

End Sub
Public Sub FormSet(ByRef FormMe As Form)
    Dim oControl As Control
    Dim oDiccionario As Object
    Dim vbuff As Variant
    Dim sUserActions As String
    Set oDiccionario = Nothing
    Centrar_form FormMe
    'IdiomaEtiquetas FormMe
End Sub

'Public Function ControlSet(ByVal clsColect As Object, ByRef oControl As Variant, ByRef sFieldName As String) As Variant
'    Dim ofield As ClieLibrary.clsField
'    Dim oMensaje As clsMensaje
'    Dim oCode As CodeMsg
'    Dim aMess(4) As Variant
'
'    ControlSet = ""
'
'    If Not (clsColect.Item(sFieldName) Is Nothing) Then
'        Set ofield = clsColect.Item(sFieldName)
'
'        If TypeOf oControl Is TextBox Then
'            oControl.MaxLength = Int(ofield.Length)
'            If oControl.Text = "" Then
'                If Not IsNull(ofield.Default) Then
'                    If VarType(ofield.Default) = vbString And (ofield.Default <> "" Or ofield.Default <> " ") Then
'                        oControl.Text = ofield.Default
'                    End If
'                End If
'
'            End If
'            If oControl.Text <> "" Then
'                ControlSet = StrZero(Val(oControl.Text), oControl.MaxLength)
'            Else
'                ControlSet = oControl.Text
'            End If
'        ElseIf UCase(Left(oControl.Name, 3)) = "INP" Then
'        'ElseIf TypeOf oControl Is InputLine Then
'            oControl.Text = ofield.Default
'            Dim sLenx As String, iPosDec As Integer, iDecPart As Integer
'
'            sLenx = CStr(ofield.Length)
'            iPosDec = InStr(sLenx, ".")
'
'            If iPosDec > 0 Then
'                iDecPart = Mid(sLenx, iPosDec + 1)
'                oControl.DecimalNumber = iDecPart
'            End If
'            oControl.Text = ofield.Default
'
'            'if ofield.
'        ElseIf TypeOf oControl Is ComboBox Then
'            If oControl.ListIndex >= 0 Then
'                ControlSet = StrZero(oControl.ItemData(oControl.ListIndex), ofield.Length)
'            End If
''        ElseIf TypeOf oControl Is MSForms.ComboBox Then
''            If oControl.ListIndex >= 0 Then
''                ControlSet = oControl.Value
''            End If
'        Else
'
'            Set oMensaje = New clsMensaje
'            oCode = KMESSAGE_ERR_MISSING_TYPEOF
'
'            oMensaje.Codigo = oCode
'            oMensaje.AttribName = sFieldName
'            'Call LoadMessage(aMess, CInt(oCode))
'            Call oMensaje.ShowMsg(aMess)
'
'        End If
'
'        Exit Function
'
'    Else
'        oCode = KMESSAGE_ERR_INCORRECT_VALIDFIELD
'        Set oMensaje = New clsMensaje
'        oMensaje.Codigo = CInt(oCode)
'        oMensaje.AttribName = sFieldName
'        'Call LoadMessage(aMess, CInt(oCode))
'        Call oMensaje.ShowMsg(aMess)
'
'    End If
'
'End Function


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

Public Sub LoadConnectEmpresa(ByVal vnewvalue As String)
    
    If Not bCargaConexion Then
    cConnect = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=facontex;Data Source=192.168.1.10"

   

    End If
End Sub

Public Sub LoadConnectSeguridad(ByVal vnewvalue As String)
    If Not bCargaConexion Then
    cSEGURIDAD = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=Seguridad;Data Source=192.168.1.10"
    
            
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
    sData = FixNulos(sData, vbString)
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

Public Function NetDate() As Date
Shell "NET TIME \\SERVERNT /SET /YES", vbHide

NetDate = Now
End Function

Public Sub MailingFile2(ByRef SendFile As String, ParamArray Recipients() As Variant)

    Dim obj As Object
    Dim omail As Object
    Dim i As Integer
    
    Set obj = CreateObject("Outlook.Application")
    Set omail = obj.CreateItem(0)
    omail.Subject = " "
    If Not IsEmpty(Recipients) Then
        For i = 0 To UBound(Recipients)
            omail.Recipients.Add (Recipients(i))
        Next
    End If
    omail.Attachments.Add (SendFile)
    omail.Display
    
    Set omail = Nothing
    Set obj = Nothing
End Sub

Public Function Mapi_SendMail(sFileName As String, ParamArray Recipients() As Variant)
    Dim oForm As Object
    Dim sRecipients As String
    'Set oForm = New frmSendMail
    Load oForm
    If Not oForm.bError Then
        oForm.FileName = sFileName
        oForm.Recipients = Recipients
        oForm.MAPIMessages1.MsgSubject = "" 'sSubject
        oForm.Load2
    End If
    Unload oForm
    Set oForm = Nothing
End Function


Public Sub MailingFile(ByRef SendFile As String, ParamArray Recipients() As Variant)
On Error GoTo hand
    Dim obj As Object
    Dim omail As Object
    Dim i As Integer
    
    Set obj = CreateObject("Outlook.Application")
    Set omail = obj.CreateItem(0)
    omail.Subject = " "
    If Not IsEmpty(Recipients) Then
        For i = 0 To UBound(Recipients)
            omail.Recipients.Add (Recipients(i))
        Next
    End If
    omail.Attachments.Add (SendFile)
    omail.Display
    
    Set omail = Nothing
    Set obj = Nothing
Exit Sub
hand:
ErrorHandler Err, "MailingFile"
End Sub
Public Sub MailingFileP(ByRef SendFile As String, ParamArray Recipients() As Variant)

    Dim obj As Object
    Dim omail As Object
    Dim i As Integer
    
    Set obj = CreateObject("Outlook.Application")
    Set omail = obj.CreateItem(0)
    omail.Subject = "Items Stock Criticos"

    If Not IsEmpty(Recipients) Then
        For i = 0 To UBound(Recipients)
            omail.Recipients.Add (Recipients(i))
        Next
    End If
    omail.Attachments.Add (SendFile)
    omail.Send
    Set omail = Nothing
    Set obj = Nothing
End Sub

Public Function SystemDirectory() As String
    Dim KeyName$
    Dim keylen&
    Dim iNull
            
    keylen& = 2000
    KeyName$ = String$(keylen, 0)
    
    GetSystemDirectory KeyName$, keylen&
    
    iNull = InStr(KeyName, Chr(0))
    SystemDirectory = Mid(KeyName$, 1, iNull - 1) + "\"
    'GetcomputerName keyname$, keylen&
End Function

Public Sub RBSToExcel(ByRef rbsData As LibraryVB.clsRecords)
On Error Resume Next
    Dim i As Integer
    Dim sTitle As String
    Dim irow As Long
    Dim iRowAll As Long
    Dim iColAll As Long
    Dim iColumn As Long
    Dim iCharColumn As Long
    
    Dim sRange As String
    Dim oExcel As Object
    If Not IsEmpty(rbsData.Buffer) Then
        Set oExcel = CreateObject("Excel.Application") ' New Excel.Application
                
        oExcel.Workbooks.Add
        oExcel.Sheets(1).Name = "AllData"
        
        iCharColumn = 64
        iColumn = 1
        For iColAll = 1 To rbsData.Count
            With oExcel
             sRange = Trim(Str(iColumn))
             .Sheets("AllData").Range(Chr(iCharColumn + iColAll) & sRange) = "" & rbsData(iColAll).Name
            End With
        Next
        iColumn = 1
        For iRowAll = 0 To rbsData.RecordCount - 1 ' .Rows - 1
            iCharColumn = 64
            iColumn = iColumn + 1
            For iColAll = 1 To rbsData.Count
                With oExcel
                 sRange = Trim(Str(iColumn))
                 .Sheets("AllData").Range(Chr(iCharColumn + iColAll) & sRange) = "" & rbsData(iColAll).Value
                End With
            Next
            rbsData.MoveNext
        Next iRowAll
        oExcel.Visible = True
        Set oExcel = Nothing
    Else
        'Mensaje KMESSAGE_INF_DATA_NOTFOUND
    End If
End Sub


Public Function ToUpper(KeyAscii As Integer) As String

    ToUpper = Asc(UCase(Chr(KeyAscii)))
End Function

Public Function RestoreRowSSDBGrid(ByRef grid As Object, ByVal irow As Variant, Optional ByVal iRows As Variant)
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



Public Sub RBSToSSDBGrid(ByRef oData As Object, ByRef pBuff As Variant, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim rsBuff As LibraryVB.clsRecords
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim X%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long

 ssDBGrid.FieldSeparator = "~"
 Set rsBuff = New LibraryVB.clsRecords
 Set rsBuff.refObject = oData

 rsBuff.Buffer = pBuff
 ssDBGrid.Redraw = False
 nCols = rsBuff.Count

 ic = ssDBGrid.Cols
 If ssDBGrid.Cols < nCols Then
    For i = nCols To ic + 1 Step -1
       ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
    Next i
 End If

 For Y = 0 To ssDBGrid.Cols - 1
   If ssDBGrid.Columns(Y).DataType = 5 Or ssDBGrid.Columns(Y).DataType = 6 Or ssDBGrid.Columns(Y).DataType = 9 Then
      ssDBGrid.Columns(Y).TagVariant = 0
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
      temp = temp & FixNulos(rsBuff(iContador + 1), vbString)
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
            ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(rsBuff(iContador + 1), vbDouble)
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
 Set rsBuff.refObject = Nothing
 Set rsBuff = Nothing

End Sub

Public Sub RBSToSSDBGridOC(ByRef oData As Object, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim rsBuff As LibraryVB.clsRecords
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim X%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long

 ssDBGrid.FieldSeparator = "~"
 Set rsBuff = New LibraryVB.clsRecords
 Set rsBuff.refObject = oData

 'rsBuff.Buffer = pBuff
 ssDBGrid.Redraw = False
 nCols = rsBuff.Count

 ic = ssDBGrid.Cols
 
 If ssDBGrid.Cols < nCols Then
    For i = nCols To ic + 1 Step -1
       ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
    Next i
 End If

 For Y = 0 To ssDBGrid.Cols - 1
   If ssDBGrid.Columns(Y).DataType = 5 Or ssDBGrid.Columns(Y).DataType = 6 Or ssDBGrid.Columns(Y).DataType = 9 Then
      ssDBGrid.Columns(Y).TagVariant = 0
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
      temp = temp & FixNulos(rsBuff(iContador + 1), vbString)
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
 Set rsBuff.refObject = Nothing
 Set rsBuff = Nothing

End Sub

Public Sub ADODBToSSDBGrid(ByVal rsBuff As ADODB.Recordset, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim X%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long
 
 ssDBGrid.FieldSeparator = "~"
 'Set rsBuff = New RBS.clsRecordSet
 'Set rsBuff.refObject = oData

 'rsBuff.Buffer = pBuff
 ssDBGrid.Redraw = False
 
 'nCols = RsBuff.Count
 nCols = rsBuff.Fields.Count

' ic = ssDBGrid.Cols
' If ssDBGrid.Cols < nCols Then
'    For i = nCols To ic + 1 Step -1
'       ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
'       ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
'       ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
'    Next i
' End If
'
' For y = 0 To ssDBGrid.Cols - 1
'   If ssDBGrid.Columns(y).DataType = 5 Or ssDBGrid.Columns(y).DataType = 6 Or ssDBGrid.Columns(y).DataType = 9 Then
'      ssDBGrid.Columns(y).TagVariant = 0
'   End If
' Next

 NVEZ = True


 X = 0
 Do While Not rsBuff.EOF
   temp = ""
   For iContador = 0 To nCols - 1
      ssDBGrid.Columns(iContador).Locked = True
      ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
      ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
      temp = temp & FixNulos(rsBuff(iContador).Value, vbString)
      If iContador < nCols - 1 Then
         temp = temp & "~"
      End If

      If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Alignment = 1
      End If

      'ssDbgrid.Columns(iContador).DataType = 5
      If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBGrid.TagVariant, vbLong) Then
        If Val(FixNulos(rsBuff(iContador).Value, vbDouble)) > 0 Then
            ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(rsBuff(iContador).Value, vbDouble)
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
 'RsBuff.Close
 'Set RsBuff = Nothing

End Sub

Public Sub ADODBToSSDBGridOC(ByVal rsBuff As ADODB.Recordset, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim X%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long
 
 ssDBGrid.FieldSeparator = "~"
 ssDBGrid.Redraw = False
 nCols = rsBuff.Fields.Count

 NVEZ = True


 X = 0
 Do While Not rsBuff.EOF
   temp = ""
   For iContador = 0 To nCols - 1
      'ssDBGrid.Columns(iContador).Locked = True
      'ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
      'ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
      temp = temp & FixNulos(rsBuff(iContador).Value, vbString)
      If iContador < nCols - 1 Then
         temp = temp & "~"
      End If

      'If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
      '      ssDBGrid.Columns(iContador).DataType = 5
      '      ssDBGrid.Columns(iContador).Alignment = 1
      'End If

      'ssDbgrid.Columns(iContador).DataType = 5
      'If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBGrid.TagVariant, vbLong) Then
      '  If Val(FixNulos(RsBuff(iContador).Value, vbDouble)) > 0 Then
      '      ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(RsBuff(iContador).Value, vbDouble)
      '  End If
      'End If
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
 'RsBuff.Close
 'Set RsBuff = Nothing

End Sub

Public Sub SSDBGridTOTALES(ByRef ssDBGrid As Object)  'SSDataWidgets_B.SSDBGrid)
On Error Resume Next
Dim iContador As Long
Dim temp As String
Dim X%
Dim Y%


 ssDBGrid.Redraw = False
 temp = ""
 For Y = 0 To ssDBGrid.Cols - 1
    temp = temp & FixNulos(ssDBGrid.Columns(Y).TagVariant, vbString)
    If Y < ssDBGrid.Cols - 1 Then
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
    
    ssDBGrid.Col = 0
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
Public Function Ancho_Columnas(ByVal fform As Form, ByVal dcontainer As Object, ByVal scadena As String)
Dim xPos As Integer
Dim xPos1 As Integer
Dim i As Integer

xPos = 1
xPos1 = 1
i = 0
Dim A As Integer
 While InStr(xPos1, scadena, ",") > 0
   xPos1 = InStr(xPos, scadena, ",") + 1
   dcontainer.Columns(i).Width = (CInt(Mid(scadena, xPos, xPos1 - xPos - 1)) * 100) + 50
   xPos = xPos1
   i = i + 1
 Wend
End Function
Public Sub DActivaControles(ByVal fform As Form, ByVal TipOpe As Variant, ByVal Scontroles As String)
Dim xEnabled As Boolean
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


Public Sub ReCarga_DBCombos(ByVal fform As Form, ByRef Rs As ADODB.Recordset)
With fform
.dbcboCategoria.ListField = "NOMBRE"
.dbcboCategoria.BoundColumn = "CODCAT"
.dbcboCategoria.BoundText = Rs!CATEGORIA

.dbcboUnidad.ListField = "NOMBRE"
.dbcboUnidad.BoundColumn = "CODUNI"
.dbcboUnidad.BoundText = Rs!UNIDAD
End With
End Sub
Public Function Maximo(ByVal stabla As String, _
ByVal sCampo As String, ByVal scondi As String, ByVal conn As ADODB.Connection, ByVal sTipo As String, ByVal ilargo As Integer)
Dim RS1 As ADODB.Recordset
Dim sQuery As String
Dim A As Variant
Dim B As Variant

Set RS1 = New ADODB.Recordset

RS1.ActiveConnection = cn
RS1.CursorType = adOpenStatic
If scondi = "" Then
 scondi = "1<2"
End If
If sTipo = "S" Then
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
Dim i As Long
Ceros = scadena
If iLen < 2 Then Exit Function
For i = 1 To iLen - 1
Ceros = schar & Ceros
Next i
End Function

Public Function Maximo1(ByVal stabla As String, _
ByVal sCampo As String, ByVal scondi As String, ByVal conn As ADODB.Connection, ByVal scampo1 As String, ByVal ilargo As Integer)
Dim RS1 As ADODB.Recordset
Dim sQuery As String
Dim A As String
Dim B As String
Dim c As String

Set RS1 = New ADODB.Recordset
RS1.ActiveConnection = cn
RS1.CursorType = adOpenStatic
If scondi = "" Then
scondi = "1<2"
End If
sQuery = "select  max(" & scampo1 & "),max(" & sCampo & ") from " & stabla & " where " & scondi
RS1.Open sQuery
    A = LTrim(RTrim(RS1(0)))
    B = ilargo - Len(A)
    c = LTrim(RTrim(Str(RS1(1) + 1)))
    If IsNull(A) Then
    A = A & Ceros("1", ilargo, "0")
    Else
    A = A & LTrim(RTrim(Ceros(c, B, "0")))
    End If
Maximo1 = A
Set RS1 = Nothing
End Function
Public Sub Carga_Categorias(ByVal fform As Form, ByVal Datag As Object, ByRef Rs As ADODB.Recordset)
Dim sQuery As String
sQuery = "SELECT COD_MOTATR AS CODIGO,DES_MOTATR AS DESCRIPCION FROM TG_MOTATR"
'Dim rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cn
Rs.CursorType = adOpenStatic
Rs.Open sQuery
With fform
Set Datag.DataSource = Rs
Set .txtIdCategoria.DataSource = Rs
.txtIdCategoria.DataField = "CODIGO"
Set .txtNombre.DataSource = Rs
.txtNombre.DataField = "DESCRIPCION"
End With
'Set rs = Nothing
End Sub

 

Function StrZero(nDato As Variant, nZeros As Integer)
   Dim wdato As String, wAncho As Integer, wDatoOk As String
   Dim tdato As Variant
   Dim i As Integer
   If TypeName(nDato) = "String" Then
    If nDato = "" Then
     StrZero = ""
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
   StrZero = wDatoOk
End Function

Public Function Rpad(Texto As Variant, ByVal iMaxLen As Long) As String

End Function


Public Function LPad(Texto As Variant, ByVal iMaxLen As Long) As String
    
End Function

Public Function LPadr(InString As String, _
                        NumChar As Integer, _
                        Optional Char As Variant) As String
                        
    Dim WithThisChar As String
    Dim StringChar As String
    Dim iIndex As Integer
    
    StringChar = ""
    WithThisChar = IIf(IsMissing(Char), Space$(1), Char)
    
    For iIndex = 1 To NumChar - Len(InString)
        StringChar = StringChar + WithThisChar
    Next
    
    LPadr = StringChar + InString

End Function



Public Sub errores(sCodigo As Long)
Dim oCode As CodeMsg
Dim oMessage As clsMessages
Dim aMess(4) As Variant
Dim sMess As String
Dim iPos As Integer

    Select Case sCodigo
        Case "9999"
            oCode = kMESSAGE_ERR_CODIGO_YA_REGISTRADO
            Set oMessage = New clsMessages
            oMessage.codigo = oCode
            Call LoadMessage(aMess, oCode)
            Call oMessage.ShowMesage(iLanguage)
            'Aviso "El Código ya ha sido registrado.  ", 1

'        Case -2147217900, -2147211505
'            oCode = KMESSAGE_ERR_REGISTRO_TIENE_TRANSAC_RELACIONADAS
'            Set omessage = New clsMensaje
'            omessage.Codigo = oCode
'            Call LoadMessage(amess, oCode)
'            Call omessage.ShowMsg(amess)

            'Aviso "No se puede efectuar la operación debido a que el registro ha sido asignado a otras Tablas", 1
        Case Else
            sMess = Err.Description
            iPos = InStr(1, sMess, "SERVER]", 1)
            If iPos > 0 Then
                sMess = Mid(sMess, iPos + 7)
            End If
            oCode = kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO
            Set oMessage = New clsMessages
            oMessage.codigo = oCode
            'oMessage.AddText = Chr(13) & " El mensaje de Error es : " & Err.Number
            'oMessage.AttribDescripLarga = Chr(13) & sMess ' Err.Description
            Call LoadMessage(aMess, oCode)
            Call oMessage.ShowMesage(iLanguage)

            'Aviso "Ha ocurrido un imprevisto !!!  " & Chr(13) & _
            'Chr(13) & "El mensaje de Error es : " & Err.Description & _
            'Chr(13) & "El Nro. de Error es : " & Err.Number, 1
    End Select

Set oMessage = Nothing
End Sub



Public Sub LoadMessage(ByRef aMess As Variant, ByVal iIndex As Integer)
aMess(0) = aMessage(iIndex).Tipo
aMess(1) = aMessage(iIndex).Code
If iLanguage = 1 Then
        aMess(2) = aMessage(iIndex).Description
    Else
        aMess(2) = aMessage(iIndex).Description
    End If
aMess(3) = aMessage(iIndex).HelpID
aMess(4) = aMessage(iIndex).Tag
End Sub






Public Function get_botones1(ByVal f As Form, ByVal Vcod_perfil As Variant, ByVal vcod_empresa As Variant, ByVal fname As Variant)
Dim RS1 As ADODB.Recordset
Set RS1 = New ADODB.Recordset
Dim sQuery As String
sQuery = "Sp_funciones3 '" & Vcod_perfil & "','" & vcod_empresa & "','" & fname & "'"
RS1.ActiveConnection = cSEGURIDAD
RS1.CursorType = adOpenStatic
RS1.Open sQuery
Dim Scad As String
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
For Each ctl In oForm.Controls
    If Not TypeOf ctl Is TextBox And Not TypeOf ctl Is FunctionsButtons.FunctButt _
     And Not TypeOf ctl Is Mantenimientos.MantFunc And Not TypeOf ctl Is ComboBox _
     And Not TypeOf ctl Is DataCombo _
     And Not TypeOf ctl Is DTPicker And Not TypeOf ctl Is MaskEdBox And Not TypeOf ctl Is LinkLabel Then
        If iLanguage <> "1" Then
            ctl.Caption = ctl.Tag
        End If
    Else
        If TypeOf ctl Is FunctionsButtons.FunctButt _
            Or TypeOf ctl Is Mantenimientos.MantFunc Then
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
If iLanguage <> "1" Then
  oForm.Caption = oForm.Tag
End If
For Each ctl In oForm.Controls
    If Not TypeOf ctl Is TextBox And Not TypeOf ctl Is FunctionsButtons.FunctButt _
     And Not TypeOf ctl Is Mantenimientos.MantFunc And Not TypeOf ctl Is ComboBox _
      And Not TypeOf ctl Is DataCombo And Not TypeOf ctl Is Image Then
        If iLanguage <> "1" Then
            ctl.Caption = ctl.Tag
        End If
    Else
        If TypeOf ctl Is FunctionsButtons.FunctButt _
            Or TypeOf ctl Is Mantenimientos.MantFunc Then
            ctl.Language = iLanguage
        End If
    End If
Next ctl
End Sub

Sub Centrar_form(ByRef Formulario As Form)
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (6945 - Formulario.Height) / 2
End Sub



Public Sub ComboBoxToComboBox(ByRef lstOrigen As Object, ByRef lstDestino As Object, ByVal iModal As Integer)
'iModal 0 Pasa item actual
'       1 Pasa todos los items
Dim i As Long

If iModal = 0 Then
    If lstOrigen.ListIndex <> -1 Then
        lstDestino.AddItem
        For i = 0 To lstOrigen.ColumnCount - 1
            
            lstDestino.Column(i, lstDestino.ListCount - 1) = lstOrigen.Column(i, lstOrigen.ListIndex)
        Next
        lstOrigen.RemoveItem lstOrigen.ListIndex
    End If
Else

End If
End Sub


Public Function GetRecordset(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object ' New ADODB.RecordSet '
  Dim objAdoCn As Object ' New ADODB.Connection '
  
 
  Set objADORs = CreateObject("ADODB.Recordset") 'New ADODB.RecordSet '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  objADORs.Open SQL, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset = objADORs
  Set GetRecordset.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 
Exit Function
ehGetRecordset:
  Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description
End Function



Public Function Refresh(ByRef rsData As Object, ByRef vBuffer As Variant) As Variant

Dim n As Integer
Dim i As Integer

If Not rsData Is Nothing Then
 n = rsData.Fields.Count - 1
 ReDim vBuffer(n, iMaxEnumField)
 For i = 0 To n
   vBuffer(i, IName) = rsData.Fields(i).Name
   vBuffer(i, iActualSize) = rsData.Fields(i).ActualSize
   vBuffer(i, iAttributes) = rsData.Fields(i).Attributes
   vBuffer(i, iDefinedSize) = rsData.Fields(i).DefinedSize
   vBuffer(i, iNumericScale) = rsData.Fields(i).NumericScale
   'vbuffer(i, iOriginalValue) = rsData.Fields(i).OriginalValue
   vBuffer(i, iOriginalValue) = rsData.Fields(i).Value
   vBuffer(i, iPrecision) = rsData.Fields(i).Precision
   vBuffer(i, iType) = rsData.Fields(i).Type
   vBuffer(i, iUnderlyingValue) = rsData.RecordCount 'rsData.Fields(i).UnderlyingValue
   vBuffer(i, iValue) = rsData.Fields(i).Value
 Next i
End If
End Function

Public Function ExecuteSQL(ByVal Connect As String, ByVal SQL As String) As Long
  'this function executes and SQL string and returns the number of records affected
  On Error GoTo ehExecuteSQL
  Dim objAdoCn As Object ' New ADODB.Connection

  Set objAdoCn = CreateObject("ADODB.Connection")    'ADO must be registered locally ' New ADODB.Connection  '
  objAdoCn.Open Connect                 'open connection
  objAdoCn.CommandTimeout = 900
  
  objAdoCn.Execute SQL, ExecuteSQL, 128  'recordsetAffected is returned
  objAdoCn.Close
  Set objAdoCn = Nothing

Exit Function
ehExecuteSQL:
 'MsgBox Err.Description
  'if transactions is not committed, it will be rolled back
  ExecuteSQL = -2                         '-2 indicates error condition
  Err.Raise Err.Number, "ExecuteSQL", Err.Description
End Function



Public Function VBsprintf(ByRef InString As String, ParamArray aInValues()) As String
Dim OutString As String
Dim ThisChar As String
Dim IndexString As Integer
Dim IndexValues As Integer
Dim iNotchar As Integer
Dim vValor As Variant
Dim strCadena As String

OutString = ""
IndexValues = 0


For IndexString = 1 To Len(InString)
 ThisChar = Mid(InString, IndexString, 1)
 
 
' If Asc(ThisChar) = 39 Then
'    MsgBox "llego "
' End If
 If ThisChar <> "$" Then
    OutString = OutString & ThisChar
 Else
   If VarType(aInValues(IndexValues)) = vbString Then
        vValor = aInValues(IndexValues)
        If Len(vValor) >= 2 Then
            If Mid(vValor, 1, 1) <> Chr(39) Then
                vValor = NotChar(vValor)
            End If
           '09/02/2000 2:08 pm
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
Dim i As Integer
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

Public Sub Mensaje(ByVal oCodeMsg As CodeMsg)
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages

amensaje.codigo = oCodeMsg
LoadMessage aMess, amensaje.codigo
amensaje.ShowMesage (iLanguage)

End Sub

Public Sub BuscarComboD(MyCombo As Object, MyKey)
    On Error Resume Next
    MyCombo.ListIndex = -1
    If MyCombo.ListCount > 0 Then
        If RTrim(MyKey) <> "" Then
            MyCombo.Value = MyKey
        End If
    End If
End Sub


Public Sub ComboBoxToComboBox1(oSource As Object, oTarget As Object, Optional bAll As Boolean = False)
Dim i As Integer
Dim j As Integer
Dim ix As Integer
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
                        vSearchFor As Variant, iIndice As Integer, _
                        Optional base As Variant) As Integer
                        
    ' Control de Parametro opcional
    
    
    Dim iIndex As Integer
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
                        vSearchFor As Variant, iIndice As Integer, _
                        Optional base As Variant) As Integer
                        
    ' Control de Parametro opcional
    
    
    Dim iIndex As Integer
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
       Dim dblF As Double
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




Public Sub EjecutaOpcion(ByVal sNameOpcion As String, perfil As String, empresa As String)  'CurrentNodeKey As String)
On Error GoTo EjecutaOpcion
    Dim tDllName As String
    Dim sopcion As String
    Dim lValDev As Long
    Dim rutexe As String
    Dim nomfor As String
    Dim nivel As Integer
    Dim Tipo As String
    Dim icono As String
    Dim cod_padre As String
    Dim des_opcion As String
    On Error GoTo EjecutaOpcion
    Get_Datos_form sNameOpcion, rutexe, nomfor, nivel, Tipo, icono, cod_padre, des_opcion

    sopcion = Tipo 'GetSubString(CurrentNodeKey, 6)
    
    tDllName = rutexe ' Trim(GetSubString(CurrentNodeKey, 3))
     
     If sopcion = "C" Or sopcion = "P" Then
        If sDllName <> tDllName Then
          If Not oFormObjDLL Is Nothing Then
            Set oFormObjDLL = Nothing
          End If

          If Not objFormDLL Is Nothing Then
            Set objFormDLL = Nothing
          End If
          sDllName = tDllName
          Set objFormDLL = CreateObject(sDllName & ".clsForm")
        End If

        Set oFormObjDLL = objFormDLL.GetForm(nomfor) 'Trim(GetSubString(CurrentNodeKey, 4)))
        If Not (oFormObjDLL Is Nothing) Then
            objFormDLL.Cod_Empresa = empresa
            objFormDLL.UserName = vusu
            objFormDLL.Cod_Perfil = perfil
            objFormDLL.Rutas = App.Path
            'objFormDLL.Cod_opcion
            objFormDLL.ConnectEmpresa = DSN_Empresa
            objFormDLL.ConnectSeguridad = DSN_Seguridad
            objFormDLL.Language = iLanguage
    On Error GoTo EjecutaOpcion
            'objFormDLL.NomEmpresa
            'objFormDLL.NomAplicacion
            'objFormDLL.NomOpcion
            oFormObjDLL.Show vbModal
            
            Set oFormObjDLL = Nothing
        End If
    Else
    End If
     Exit Sub
EjecutaOpcion:
    ErrorHandler Err, "EjecutaOpcion"
    Set oFormObjDLL = Nothing
    'Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function Get_Datos_form(ByVal sopcion As String, ByRef rutexe As String, ByRef nomfor As String, ByRef nivel As Integer, ByRef Tipo As String, ByRef icono As String, ByRef cod_padre As String, ByRef des_opcion As String)
    Dim iCount As Integer
    Dim mRs As ADODB.Recordset
    
    sQuery = "SELECT isnull(RUTEXE,''),isnull(nomfor,''),isnull(nivel,0),isnull(tipo,''),isnull(icono,''),isnull(cod_padre,''),isnull(des_opcion,'') FROM SEG_OPCIONES  WHERE COD_OPCION='" & sopcion & "'"
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
    End If
    Set mRs = Nothing
End Function

Sub Informa(ByVal Mens As String, Optional ByVal amensaje As clsMessages)
If Mens <> "" Then
    Dim rpta As Byte
    rpta = MsgBox(Mens, vbInformation, "Informa")
    Exit Sub
End If
Dim aMess(4)
LoadMessage aMess, amensaje.codigo
amensaje.ShowMesage (iLanguage)
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

