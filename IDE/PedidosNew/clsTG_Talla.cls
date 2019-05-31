VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_Talla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const TRANSACTION_SERVER = False

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

Public Function ViewxTalla(ByVal sCod_Cliente As String, _
                           ByVal sCod_TemCli As String, _
                           ByVal sCod_EstCliLot As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_TallaViewAll '$' , '$' ,'$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_TemCli, sCod_EstCliLot)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewxTalla = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewxTalla Method", Err.Description
End Function

Public Function Add(ByVal sCod_Talla As String, ByVal sCod_GruTal As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_Talla '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Talla, sCod_GruTal)
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

