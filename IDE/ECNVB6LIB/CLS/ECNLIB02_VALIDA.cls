VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB02_VALIDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'************************************************************************************************************************************************************
' AUTOR       : EDGAR I. CARDENAS NAVARRO
' FECHA       : 09/05/2013
' NOMBRE      : ECNLIB03
' DESCRIPCION : CLASE CONTENEDORA DE LOS PROCEDIMIENTOS Y FUNCIONES DE VALIDACIONES GENERALES
'************************************************************************************************************************************************************

Public Function RsEsValidoParaLectura(ByRef xRs As ADODB.Recordset) As Boolean
    On Error Resume Next
    RsEsValidoParaLectura = True
    If xRs Is Nothing Then RsEsValidoParaLectura = False
    If xRs.BOF And xRs.EOF Then RsEsValidoParaLectura = False
    If xRs.RecordCount = 0 Then RsEsValidoParaLectura = False
End Function
