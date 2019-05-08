VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_PurOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const TRANSACTION_SERVER = False
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

Public Function ViewDetalle(ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_TemCli As String, ByVal mCod_EstCli As String, ByVal mCod_OrdPro As String, ByVal mCod_Usuario As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_PurOrd_ViewDetalle '$','$','$','$','$','$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_PurOrd, mCod_TemCli, mCod_EstCli, mCod_Usuario, mCod_OrdPro)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewDetalle = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "View All Method", Err.Description
End Function

Public Function ViewEstilos(ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_LotEst_ViewDetalle_NEW '$','$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_PurOrd)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewEstilos = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "View All Method", Err.Description
End Function

Public Function Update() As Boolean
On Error GoTo UpdateErr
Dim SSQL As String
#If TRANSACTION_SERVER = True Then
 Dim ctxObject As ObjectContext
 Set ctxObject = GetObjectContext()
#End If


 ExecuteCommandSQL mvarConnect, SSQL
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

Dim SSQL As String
SSQL = "exec "
 'SSQL = VBsprintf(SSQL, )
 ExecuteCommandSQL mvarConnect, SSQL
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

 Dim rs As Object
Dim SSQL As String
SSQL = "exec "
 'SSQL = VBsprintf(SSQL, )
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
 Err.Raise Err.Number, "Add Method", Err.Description
End Function

Public Function Valid(ByVal mAno As String, ByVal mCod_Hermano As String) As Boolean
On Error GoTo ValidErr

Dim SSQL As String
SSQL = "exec "
 'SSQL = VBsprintf(SSQL, )
 Set mRs = GetDataSet(mvarConnect, SSQL)

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
    mName = "TG_PurOrd"
End Sub


Public Function ViewAllClaPurOrd() As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_ClaPurOrd "
 Set mRs = GetDataSet(mvarConnect, SSQL)
 ViewAllClaPurOrd = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewAllClaPurOrd Method", Err.Description
End Function



Public Function ViewGruTal(ByVal sCod_GruTal As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_ES_Tallas '$'"
  SSQL = VBsprintf(SSQL, sCod_GruTal)
 Set mRs = GetDataSet(mvarConnect, SSQL)

 ViewGruTal = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewGruTal Method", Err.Description
End Function


Public Function ViewTipPre(ByVal sCod_TipPre As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_TipPre '$'"
 SSQL = VBsprintf(SSQL, sCod_TipPre)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 
 ViewTipPre = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewTipPre Method", Err.Description
End Function

Public Function ViewEstPropio(ByVal sCod_Estpro As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_ES_EstPro '$'"
 SSQL = VBsprintf(SSQL, sCod_Estpro)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 
 ViewEstPropio = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewEstPropio Method", Err.Description
End Function



Public Function AddEstPro(ByVal sDes_EStPro As String, ByVal sCod_TipPre As String, ByVal sCod_GruTal As String) As String
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec SI_ES_EStPro '$','$','$'"
 SSQL = VBsprintf(SSQL, sDes_EStPro, sCod_TipPre, sCod_GruTal)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 If Not mRs.EOF And Not mRs.EOF Then
    AddEstPro = mRs!Cod_EstPro
 End If
 
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddEstPro = False
 Err.Raise Err.Number, "Add Method", Err.Description
End Function

Public Function GenerarOP(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sFlg_Modo As String, ByVal sUsuario As String) As Boolean
 On Error GoTo GenerarOPErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec UP_GENERA_ORDENES_PO '$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, sFlg_Modo, sUsuario)
 ExecuteCommandSQL mvarConnect, SSQL
 
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 GenerarOP = True
 Exit Function
GenerarOPErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 GenerarOP = False
 Err.Raise Err.Number, "GenerarOP Method", Err.Description
End Function


Public Function ViewOPS(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec sm_lotestpro_view '$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 
 ViewOPS = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewOPS Method", Err.Description
End Function


Public Function AsignarOp(ByVal sCod_Fabrica As String, ByVal sCod_Ordpro As String) As Boolean
 On Error GoTo GenerarOPErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec UP_ACTUALIZA_TOTAL_ORDENES_PO '$','$'"
 SSQL = VBsprintf(SSQL, sCod_Fabrica, sCod_Ordpro)
 ExecuteCommandSQL mvarConnect, SSQL
 
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AsignarOp = True
 Exit Function
GenerarOPErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AsignarOp = False
 Err.Raise Err.Number, "GenerarOP Method", Err.Description
End Function



Public Function DesasignarOp(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_Ordpro As String) As Boolean
 On Error GoTo GenerarOPErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec UP_DESASIGNA_LOTESTPRO '$','$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_Ordpro)
 ExecuteCommandSQL mvarConnect, SSQL
 
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 DesasignarOp = True
 Exit Function
GenerarOPErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 DesasignarOp = False
 Err.Raise Err.Number, "GenerarOP Method", Err.Description
End Function

Public Function AyudaAsignaOPS(ByVal sCod_Fabrica As String, ByVal sCod_Cliente As String, ByVal sCod_Estpro As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_ES_ORDPRO_VIEWASG '$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Fabrica, sCod_Cliente, sCod_Estpro)
 Set mRs = GetDataSet(mvarConnect, SSQL)
 
 AyudaAsignaOPS = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "AyudaAsignaOPS Method", Err.Description
End Function
Public Function AsignarLotesEstpro(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_Estpro As String, ByVal sCod_Ordpro As String) As Boolean
 On Error GoTo GenerarOPErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
 Dim SSQL As String
 SSQL = "exec UP_ASIGNA_LOTESTPRO '$','$','$','$','$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_Estpro, sCod_Ordpro)
 ExecuteCommandSQL mvarConnect, SSQL
 
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AsignarLotesEstpro = True
 Exit Function
GenerarOPErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AsignarLotesEstpro = False
 Err.Raise Err.Number, "AsignarLotesEstpro Method", Err.Description
End Function
