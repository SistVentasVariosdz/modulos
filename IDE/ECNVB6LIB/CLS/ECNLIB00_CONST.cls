VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB00_CONST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************************************************************
' AUTOR       : EDGAR I. CARDENAS NAVARRO
' FECHA       : 04/07/2013
' NOMBRE      : ECNLIB00_CONST
' DESCRIPCION : LIBRERIA DE CONSTANTES GENERALES A UTILIZAR EN SISTEMAS ECN
'************************************************************************************************************************************************************
Option Explicit

Private Const Ks_VAL_WIN_MIN           As String = "0"
Private Const Ks_VAL_WIN_MAX           As String = "1"

Private Const Ks_VAL_CHECK             As String = "1"
Private Const Ks_VAL_UNCHK             As String = "0"

Private Const Ks_VAL_NULL              As String = "<NULL>"

Private Const Ks_CARESP_WEB_RESTAURADO As String = "2" '=> Con Letra Webdings
Private Const Ks_CARESP_WEB_MINIMIZADO As String = "0" '=> Con Letra Webdings
Private Const Ks_CARESP_WEB_IZQUIERDA  As String = "3" '=> Con Letra Webdings
Private Const Ks_CARESP_WEB_DERECHA    As String = "4" '=> Con Letra Webdings
Private Const Ks_CARESP_WEB_ARRIBA     As String = "5" '=> Con Letra Webdings
Private Const Ks_CARESP_WEB_ABAJO      As String = "6" '=> Con Letra Webdings

Private Const Ks_CARESP_CURFIXED       As String = "®" '=> Con Letra Windings
Private Const Ks_CARESP_INSERTAR       As String = "<" '=> Con Letra Windings
Private Const Ks_CARESP_MODIFICA       As String = "!" '=> Con Letra Windings

Private Const Ks_CARESP_OPT_CHECKED    As String = "¤" '=> Con Letra Windings
Private Const Ks_CARESP_OPT_UNCHECK    As String = "¡" '=> Con Letra Windings

Private Const Ks_CARESP_CHK_CHECKED_00 As String = "a" '=> Con Letra Webdings
Private Const Ks_CARESP_CHK_CHECKED_01 As String = "ü" '=> Con Letra Windings

Property Get VAL_WIN_MIN() As String
    VAL_WIN_MIN = Ks_VAL_WIN_MIN
End Property

Property Get VAL_WIN_MAX() As String
    VAL_WIN_MAX = Ks_VAL_WIN_MAX
End Property

Property Get VAL_CHECK() As String
    VAL_CHECK = Ks_VAL_CHECK
End Property

Property Get VAL_UNCHK() As String
    VAL_UNCHK = Ks_VAL_UNCHK
End Property

Property Get VAL_NULL() As String
    VAL_NULL = Ks_VAL_NULL
End Property

Property Get CARESP_WEB_RESTAURADO() As String
    CARESP_WEB_RESTAURADO = Ks_CARESP_WEB_RESTAURADO
End Property

Property Get CARESP_WEB_MINIMIZADO() As String
    CARESP_WEB_MINIMIZADO = Ks_CARESP_WEB_MINIMIZADO
End Property

Property Get CARESP_WEB_IZQUIERDA() As String
    CARESP_WEB_IZQUIERDA = Ks_CARESP_WEB_IZQUIERDA
End Property

Property Get CARESP_WEB_DERECHA() As String
    CARESP_WEB_DERECHA = Ks_CARESP_WEB_DERECHA
End Property

Property Get CARESP_WEB_ARRIBA() As String
    CARESP_WEB_ARRIBA = Ks_CARESP_WEB_ARRIBA
End Property

Property Get CARESP_WEB_ABAJO() As String
    CARESP_WEB_ABAJO = Ks_CARESP_WEB_ABAJO
End Property

Property Get CARESP_CURFIXED() As String
    CARESP_CURFIXED = Ks_CARESP_CURFIXED
End Property

Property Get CARESP_INSERTAR() As String
    CARESP_INSERTAR = Ks_CARESP_INSERTAR
End Property

Property Get CARESP_MODIFICA() As String
    CARESP_MODIFICA = Ks_CARESP_MODIFICA
End Property

Property Get CARESP_OPT_CHECKED() As String
    CARESP_OPT_CHECKED = Ks_CARESP_OPT_CHECKED
End Property

Property Get CARESP_OPT_UNCHECK() As String
    CARESP_OPT_UNCHECK = Ks_CARESP_OPT_UNCHECK
End Property

Property Get CARESP_CHK_CHECKED_00() As String
    CARESP_CHK_CHECKED_00 = Ks_CARESP_CHK_CHECKED_00
End Property

Property Get CARESP_CHK_CHECKED_01() As String
    CARESP_CHK_CHECKED_01 = Ks_CARESP_CHK_CHECKED_01
End Property
