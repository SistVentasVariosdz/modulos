VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls004_ECNNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Show()
    Load frm004_ECNNET
    With frm004_ECNNET
        .Show
    End With
    Set frm004_ECNNET = Nothing
End Sub

Public Sub ShowPrompt()
    Load frm004_ECNNET
    With frm004_ECNNET
        .Show 1
    End With
    Set frm004_ECNNET = Nothing
End Sub

Public Function GetForm() As Object
    Set GetForm = frm004_ECNNET
End Function

Public Property Get WIN_RESULT() As GE_WIN_RESULT
    WIN_RESULT = GO_004_ENU_OPC_WIN_RESULT
End Property

Public Property Get WINAPPPATH() As String
    WINAPPPATH = GO_004_RUTA_ECN_APP
End Property

Public Property Let FORM_PARENT(ByVal xForm As Object)
    Set GO_004_FORM_PARENT = xForm
End Property

Public Property Get FORM_PARENT() As Object
    FORM_PARENT = GO_004_FORM_PARENT
End Property

Public Property Get COMPUTER_SEL() As String
    COMPUTER_SEL = GO_004_COMPUTER_SEL
End Property

