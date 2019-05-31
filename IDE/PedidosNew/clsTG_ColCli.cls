VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_ColCli"
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

Public Function ViewxClie(ByVal mCod_Cliente As String, _
                          ByVal mCod_TemCli As String, _
                          ByVal mCod_EstCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_ColCliViewxClie '$' , '$' , '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_TemCli, mCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewxClie = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewxClie Method", Err.Description
End Function

Public Function ViewTipCol() As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_TipCol "
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewTipCol = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTipCol Method", Err.Description
End Function

Public Function ViewClasePO(ByVal mCod_ClasePO As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec TG_MUESTRA_CLASE_PO_PERMITIDAS '$'"
    sSQl = VBsprintf(sSQl, mCod_ClasePO)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewClasePO = RowsDataSet()
 
    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewClasePO Method", Err.Description
End Function

Public Function Add(ByVal sCod_Cliente As String, _
                    ByVal scod_colcli As String, _
                    ByVal sDes_ColCli As String, _
                    ByVal sCod_TipCol As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_ColCli '$','$','$','$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, scod_colcli, sDes_ColCli, sCod_TipCol)
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
    Err.Raise Err.Number, "AddEStCli Method", Err.Description
End Function

Public Function ViewTG_ColCli(ByVal mCod_Cliente As String, _
                              ByVal mCod_ColCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_ColCli '$' ,'$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_ColCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewTG_ColCli = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTG_ColCli Method", Err.Description
End Function

Public Function ViewTG_EstCliCol(ByVal mCod_Cliente As String, _
                                 ByVal mCod_TemCli As String, _
                                 ByVal mCod_EstCli As String, _
                                 ByVal mCod_ColCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_EstCliCol_View '$' ,'$' ,'$','$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_TemCli, mCod_EstCli, mCod_ColCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewTG_EstCliCol = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTG_EstCliCol Method", Err.Description
End Function

Public Function AddEStCliCol(ByVal sCod_Cliente As String, _
                             ByVal sCod_TemCli, _
                             ByVal sCod_EstCli As String, _
                             ByVal scod_colcli As String, _
                             ByVal sNom_ColCli As String, _
                             ByVal sCod_EstPro As String, _
                             ByVal sCod_Present As String, _
                             ByVal sDat_UltMod As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_EstCliCol '$','$','$','$','$','$','$','$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_TemCli, sCod_EstCli, scod_colcli, sNom_ColCli, sCod_EstPro, sCod_Present, sDat_UltMod)
    ExecuteCommandSQL mvarConnect, sSQl
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetComplete
    #End If
    AddEStCliCol = True

    Exit Function

AddErr:
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetAbort
    #End If
    AddEStCliCol = False
    Err.Raise Err.Number, "AddEStCliCol Method", Err.Description
End Function

Public Function ViewAllPresent_EstPro(ByVal mCod_EstPro As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_ES_EstProPre_ViewAll  '$'"
    sSQl = VBsprintf(sSQl, mCod_EstPro)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewAllPresent_EstPro = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTG_EstCliCol Method", Err.Description
End Function

Public Function AddEstProPre(ByVal mCod_EstPro As String, _
                             ByVal mDes_Present As String) As Boolean

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SI_ES_EstProPre  '$' , '$'"
    sSQl = VBsprintf(sSQl, mCod_EstPro, mDes_Present)
    ExecuteCommandSQL mvarConnect, sSQl
    AddEstProPre = True

    Exit Function

ViewAllErr:
    AddEstProPre = True
    Err.Raise Err.Number, "AddEstProPre Method", Err.Description
End Function

