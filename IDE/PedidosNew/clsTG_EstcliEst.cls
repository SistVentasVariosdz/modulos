VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_EstclIEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const TRANSACTION_SERVER = False
Option Base 0
Option Explicit

Private mvarConnect As String

Private mName       As String

Private mRs         As Object

Dim vBuffProp()     As Variant

Public Function Retrieve()
End Function

Public Property Let ConexionString(ByVal vData As String)
    mvarConnect = vData
End Property

Public Property Get ConexionString() As String
    ConexionString = mvarConnect
End Property

Public Function RowsDataSet() As Variant

    If Not mRs.EOF Then
        Call Refresh(mRs, vBuffProp)
        RowsDataSet = mRs.GetRows()
    Else
        mRs.Close
        Set mRs = Nothing
        RowsDataSet = Empty
    End If

End Function

Public Function rsGetProperties() As Variant
    rsGetProperties = vBuffProp
End Function

Public Function ViewEStCliEst(ByVal mCod_Cliente As String, _
                              ByVal mCod_TemCli As String, _
                              Optional ByVal strCodClaPO As String = "") As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    'SSQL = "exec SM_TG_EstCliEst_ViewxCliente '$' ,'$'"
    'SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_TemCli)
    
    sSQl = "exec SM_TG_EstCliEst_ViewxCliente_PO '$' , '$', '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_TemCli, strCodClaPO)
    
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewEStCliEst = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewEStCliEst  Method", Err.Description
End Function

Public Function Update() As Boolean

    On Error GoTo UpdateErr

    Dim sSQl As String

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    ExecuteCommandSQL mvarConnect, sSQl
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetComplete
    #End If
    Update = True

    Exit Function

UpdateErr:
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetAbort
    #End If
    Update = False
    Err.Raise Err.Number, "Update Method", Err.Description
End Function

Public Function Delete() As Boolean

    On Error GoTo DeleteErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim sSQl As String

    sSQl = "exec "
    'SSQL = VBsprintf(SSQL, )
    ExecuteCommandSQL mvarConnect, sSQl
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetComplete
    #End If
    Delete = True

    Exit Function

DeleteErr:
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetAbort
    #End If

    Delete = False
    Err.Raise Err.Number, "Delete Method", Err.Description
End Function

Public Function Add() As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec "
    'SSQL = VBsprintf(SSQL, )
    ExecuteCommandSQL mvarConnect, sSQl
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetComplete
    #End If
    Add = True

    Exit Function

AddErr:
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetAbort
    #End If
    Add = False
    Err.Raise Err.Number, "Add Method", Err.Description
End Function

Public Function Valid(ByVal mAno As String, ByVal mCod_Hermano As String) As Boolean

    On Error GoTo ValidErr

    Dim sSQl As String

    sSQl = "exec "
    'SSQL = VBsprintf(SSQL, )
    Set mRs = GetDataSet(mvarConnect, sSQl)

    If mRs.EOF Then
        Valid = True
    Else
        Valid = False
    End If

    Exit Function

ValidErr:
    Valid = False
    Err.Raise Err.Number, "Valid Method", Err.Description
End Function

Private Sub Class_Initialize()
    mName = "TG_EstCliEst"
End Sub

