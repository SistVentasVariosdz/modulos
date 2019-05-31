VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_Fabrica"
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

Public Function ViewFabricas(ByVal mAbr_Fabrica As String, _
                             ByVal mCod_Usuario As String, _
                             ByVal mCod_Cliente As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_Fabrica_VieShort '$' , '$' ,'$'"
    sSQl = VBsprintf(sSQl, mAbr_Fabrica, mCod_Usuario, mCod_Cliente)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewFabricas = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewFabricas Method", Err.Description
End Function

Private Sub Class_Initialize()
    mName = "TG_Fabrica"
End Sub

