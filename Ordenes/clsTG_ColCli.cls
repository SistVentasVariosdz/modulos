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
Private mName As String
Private mRs As Object
Dim vBuffProp() As Variant

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

Public Function ViewxClie(ByVal mCod_Cliente As String, ByVal mCod_TemCli As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_ColCliViewxClie '$' , '$' , '$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_TemCli, mCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewxClie = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewxClie Method", Err.Description
End Function

Public Function ViewTipCol() As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_TipCol "
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewTipCol = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewTipCol Method", Err.Description
End Function

Public Function ViewClasePO(ByVal mCod_ClasePO As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec TG_MUESTRA_CLASE_PO_PERMITIDAS '$'"
 SSQL = VBsprintf(SSQL, mCod_ClasePO)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewClasePO = RowsDataSet()
 
 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewClasePO Method", Err.Description
End Function


Public Function Add(ByVal sCod_Cliente As String, ByVal sCod_ColCli As String, ByVal sDes_ColCli As String, ByVal sCod_TipCol As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec SI_TG_ColCli '$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_ColCli, sDes_ColCli, sCod_TipCol)
 ExecuteCommandSQL mvarConnect, SSQL
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


Public Function ViewTG_ColCli(ByVal mCod_Cliente As String, ByVal mCod_ColCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_ColCli '$' ,'$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_ColCli)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewTG_ColCli = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewTG_ColCli Method", Err.Description
End Function

Public Function ViewTG_EstCliCol(ByVal mCod_Cliente As String, ByVal mCod_TemCli As String, ByVal mCod_EstCli As String, ByVal mCod_ColCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_EstCliCol_View '$' ,'$' ,'$','$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_TemCli, mCod_EstCli, mCod_ColCli)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewTG_EstCliCol = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewTG_EstCliCol Method", Err.Description
End Function



Public Function AddEStCliCol(ByVal sCod_Cliente As String, ByVal sCod_TemCli, ByVal sCod_EStCli As String, ByVal sCod_ColCli As String, ByVal sNom_ColCli As String, ByVal sCod_Estpro As String, ByVal sCod_Present As String, ByVal sDat_UltMod As String) As Boolean
On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec SI_TG_EstCliCol '$','$','$','$','$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_TemCli, sCod_EStCli, sCod_ColCli, sNom_ColCli, sCod_Estpro, sCod_Present, sDat_UltMod)
 ExecuteCommandSQL mvarConnect, SSQL
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
Dim SSQL As String
 SSQL = "exec SM_ES_EstProPre_ViewAll  '$'"
 SSQL = VBsprintf(SSQL, mCod_EstPro)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewAllPresent_EstPro = RowsDataSet()

 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewTG_EstCliCol Method", Err.Description
End Function


Public Function AddEstProPre(ByVal mCod_EstPro As String, ByVal mDes_Present As String) As Boolean
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SI_ES_EstProPre  '$' , '$'"
 SSQL = VBsprintf(SSQL, mCod_EstPro, mDes_Present)
 ExecuteCommandSQL mvarConnect, SSQL
 AddEstProPre = True

 Exit Function
ViewAllErr:
  AddEstProPre = True
  Err.Raise Err.Number, "AddEstProPre Method", Err.Description
End Function

