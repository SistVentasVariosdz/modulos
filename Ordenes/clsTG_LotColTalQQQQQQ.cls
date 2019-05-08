VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_LotColTal"
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

Public Property Let Connect(ByVal vData As String)
    mvarConnect = vData
End Property

Public Property Get Connect() As String
    Connect = mvarConnect
End Property

Public Function GetRows() As Variant
If Not mRs.EOF Then
 Call Refresh(mRs, vBuffProp)
 GetRows = mRs.GetRows()
Else
 mRs.Close
 Set mRs = Nothing
 GetRows = Empty
End If
End Function

Public Function rsGetProperties() As Variant
  rsGetProperties = vBuffProp
End Function

Public Function ViewMatriz(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_EstCli As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String

 SSQL = "exec SG_GeneraMatriz '$' ,'$' ,'$' , '$'  , '$', '$' ,'$'"
 SSQL = VBsprintf(SSQL, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_EstCli, mCod_Fabrica, mCod_Destino, mFec_Despacho)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewMatriz = GetRows()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function


Public Function ViewMatrizKey(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String

 SSQL = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewMatrizKey = GetRows()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function

Public Function ViewMatrizKeyUpdate(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String

 SSQL = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$' , $"
 SSQL = VBsprintf(SSQL, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli, 0)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewMatrizKeyUpdate = GetRows()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatrizKeyUpdate Method", Err.Description
End Function

Public Function ViewVectorColorKey(ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String

 SSQL = "exec SG_GeneraVectorKey '$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewVectorColorKey = GetRows()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewVectorColorKey Method", Err.Description
End Function


Public Function ViewMatrizDif(ByVal iSecuencia As Integer, ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String

 SSQL = "exec SG_GeneraMatrizDif $, '$' ,'$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, iSecuencia, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewMatrizDif = GetRows()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function



Public Function LimpiaMatrizKeyEnTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String)
On Error GoTo LimpiaMatrizKeyEnTemporalErr
Dim SSQL As String

 SSQL = "exec SG_LimpiaMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 ExecuteSQL mvarConnect, SSQL
 


 Exit Function
LimpiaMatrizKeyEnTemporalErr:
  Err.Raise Err.Number, "LimpiaMatrizKeyEnTemporal Method", Err.Description
End Function


Public Function SaveToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mPrecio As Double, ByVal mWithKey As Integer, ByVal mCod_DivPRe As String)
On Error GoTo SaveToTemporalErr
Dim SSQL As String

 SSQL = "exec SI_TM_LotColTal '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' ,$ , $ , '$'"
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mWithKey, mCod_DivPRe)
 ExecuteSQL mvarConnect, SSQL
 


 Exit Function
SaveToTemporalErr:
  Err.Raise Err.Number, "SaveToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mNum_PreReq As Long)
On Error GoTo SaveCantidadesToTemporalErr
Dim SSQL As String

 SSQL = "exec SI_TM_LotColTalCantidadesRequeridas '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $ "
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mNum_PreReq)
 ExecuteSQL mvarConnect, SSQL
 


 Exit Function
SaveCantidadesToTemporalErr:
  Err.Raise Err.Number, "SaveCantidadesToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal2(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mNum_PreDesp As Long, ByVal iSecuencia As Integer)
On Error GoTo SaveCantidadesToTemporalErr
Dim SSQL As String

 SSQL = "exec SI_Actualiza_LotCotTal '$','$','$','$','$','$','$',$,$"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mNum_PreDesp, iSecuencia)
 ExecuteSQL mvarConnect, SSQL

 Exit Function
SaveCantidadesToTemporalErr:
  Err.Raise Err.Number, "SaveCantidadesToTemporal2 Method", Err.Description
End Function

Public Function SavePreciosToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mPrecio As Double, ByVal mCod_DivPRe As String)
On Error GoTo SavePreciosToTemporalErr
Dim SSQL As String

 SSQL = "exec SI_TM_LotColTalPrecios '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $  , '$'"
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mCod_DivPRe)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
SavePreciosToTemporalErr:
  Err.Raise Err.Number, "SavePreciosToTemporal Method", Err.Description
End Function


Public Function GenerarInformacion(ByVal sModalUpdate As String, ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_ClaPurOrd As String, ByVal sFec_DespachoAct As String, ByVal sCod_PagEmb As String, ByVal sCod_Fabrica As String, ByVal sCod_Destino As String, ByVal sCod_Embarque As String, ByVal sCod_Moneda As String, ByVal sCod_DivCli As String, ByVal sCod_TemCli As String, ByVal sFlg_Carta As String, ByVal sCod_BAnco As String, ByVal dPor_Slush As Double, ByVal sDes_General As String, ByVal sDes_Despacho As String, ByVal dPor_Comision As Double, ByVal scod_LotPurOrd As String, ByVal sCod_FabricaLot As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal sCod_EStCli As String, ByVal dPor_ComisionLot As Double, ByVal dPrecioLOT As Double, ByVal sCod_UsuUltAct As String, ByVal sCod_EstUltAct As String, ByVal sFlg_OrdenREorden As String, ByVal sFlg_Regular As String, ByVal sFec_DespachoOriLOT As String, _
ByVal dPrecioClienteLot As Double, ByVal sFechaIngAlmClie As String)
'ModalUpdate , Cod_Usuario, Cod_Cliente, Cod_PurOrd, Cod_ClaPurOrd, Fec_DespachoAct, Cod_PagEmb, Cod_Fabrica, Cod_Destino, Cod_Embarque, Cod_Moneda, Cod_DivCli, Cod_TemCli, Flg_Carta, Cod_Banco, Por_Slush, Des_General, Des_Despacho, Por_Comision, Cod_LotPurOrd, Cod_FabricaLot, Cod_DestinoLot, Fec_DespachoActLot, Cod_EstCli, Por_ComisionLot, PrecioLot, Cod_UsuUltAct, Cod_EstUltAct
On Error GoTo GenerarInformacionErr
Dim SSQL As String

 SSQL = "exec SM_TG_LotColTalPrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $, '$', '$', '$', '$', '$', $, $, '$', '$','$','$' ,'$', $ , '$'"
 SSQL = VBsprintf(SSQL, sModalUpdate, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, scod_LotPurOrd, sCod_FabricaLot, sCod_DestinoLOT, sFec_DespachoActLOT, sCod_EStCli, dPor_ComisionLot, dPrecioLOT, sCod_UsuUltAct, sCod_EstUltAct, sFlg_OrdenREorden, sFlg_Regular, sFec_DespachoOriLOT, dPrecioClienteLot, sFechaIngAlmClie)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
GenerarInformacionErr:
  Err.Raise Err.Number, "GenerarInformacion Method", Err.Description
End Function

Public Function UpdateInformacion(ByVal sModalUpdate As String, ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal mCod_ClaPurOrd As String, ByVal mCod_Moneda As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal dPor_ComisionLot As Double, ByVal sCod_UsuUltAct As String, ByVal sCod_EstUltAct As String, ByVal mPrecioCliente As Double, ByVal mFecIngAlm As String)
On Error GoTo GenerarInformacionErr
Dim SSQL As String

 SSQL = "exec SM_TG_LotColTalUpdate  '$', '$', '$', '$', '$' , '$', '$', '$', '$', $, '$', '$' ,  $ , '$'"
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, mCod_ClaPurOrd, mCod_Moneda, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_ComisionLot, sCod_UsuUltAct, sCod_EstUltAct, mPrecioCliente, mFecIngAlm)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
GenerarInformacionErr:
  Err.Raise Err.Number, "GenerarInformacion Method", Err.Description
End Function

Public Function UpdatePOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_ClaPurOrd As String, ByVal sFec_DespachoAct As String, ByVal sCod_PagEmb As String, ByVal sCod_Fabrica As String, ByVal sCod_Destino As String, ByVal sCod_Embarque As String, ByVal sCod_Moneda As String, ByVal sCod_DivCli As String, ByVal sCod_TemCli As String, ByVal sFlg_Carta As String, ByVal sCod_BAnco As String, ByVal dPor_Slush As Double, ByVal sDes_General As String, ByVal sDes_Despacho As String, ByVal dPor_Comision As Double, ByVal sFlg_OrdenREorden As String, ByVal sFlg_Regular As String)

On Error GoTo UpdatePOCErr
Dim SSQL As String

 SSQL = "exec SM_TG_PurOrd_AddUpdatePrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $ ,'$' ,'$' "
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, sFlg_OrdenREorden, sFlg_Regular)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
UpdatePOCErr:
  Err.Raise Err.Number, "UpdatePOC Method", Err.Description
End Function


Public Function DeletePOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String)
On Error GoTo DeletePOCErr
Dim SSQL As String

 SSQL = "exec SM_TG_PurOrd_DeletePrincipal '$', '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
DeletePOCErr:
  Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function

Public Function DeleteLotPurOrd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd)
On Error GoTo DeletePOCErr
Dim SSQL As String

 SSQL = "exec SM_TG_LotPurOrd_DeletePrincipal '$', '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd)
 ExecuteSQL mvarConnect, SSQL


 Exit Function
DeletePOCErr:
  Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function



Public Function LoadDataPOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_Usuario As String) As Variant

On Error GoTo LoadDataPOCErr
Dim SSQL As String

 SSQL = "exec SM_TG_PurOrd_ViewDetalle '$', '$', '$', '$', '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, "", "", sCod_Usuario)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 LoadDataPOC = GetRows()



 Exit Function
LoadDataPOCErr:

  Err.Raise Err.Number, "LoadDataPOC Method", Err.Description
End Function


Public Function LoadDataLOTEST(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String) As Variant

On Error GoTo LoadDataLOTESTErr
Dim SSQL As String

 SSQL = "exec SM_TG_LOTEST_ViewDATA '$', '$', '$', '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 LoadDataLOTEST = GetRows()



 Exit Function
LoadDataLOTESTErr:

  Err.Raise Err.Number, "LoadDataLOTEST Method", Err.Description
End Function



Public Function AddDivPre(ByVal sCod_DivPRe As String, ByVal sDes_DivPRe As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_DivPre '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_DivPRe, sDes_DivPRe)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddDivPre = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddDivPre = False
 Err.Raise Err.Number, "AddDivPre Method", Err.Description
End Function




Public Function VAlidDivPre(ByVal mCod_DivPRe As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_DivPre_View '$' "
 SSQL = VBsprintf(SSQL, mCod_DivPRe)
 Set mRs = GetRecordset(mvarConnect, SSQL)

 If mRs.EOF Then
  VAlidDivPre = False
 Else
  VAlidDivPre = True
 End If

 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "VAlidDivPre Method", Err.Description
End Function


Public Function ViewDivPre(ByVal mCod_DivPRe As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SM_TG_DivPre_ViewShort '$' "
 SSQL = VBsprintf(SSQL, mCod_DivPRe)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 ViewDivPre = GetRows

 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewDivPre Method", Err.Description
End Function



Public Function VAlidFechaDespacho(ByVal mFecha As String) As Variant
On Error GoTo ViewAllErr
Dim SSQL As String
 SSQL = "exec SG_TG_RCtrlVAlidFEcha '$' "
 SSQL = VBsprintf(SSQL, mFecha)
 Set mRs = GetRecordset(mvarConnect, SSQL)
 If Not mRs.EOF Then
    VAlidFechaDespacho = FixNulos(mRs!Ret, vbInteger)
 End If
 Exit Function
ViewAllErr:
  VAlidFechaDespacho = -1
  Err.Raise Err.Number, "ValidFechaDespacho Method", Err.Description
End Function




Public Function UpdateDatGenPurORd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal dPor_Comision As Double, ByVal sCod_Usuario As String, ByVal sCod_EStacion As String, ByVal sFec_DespachoOriLOT As String, ByVal dPrecioLOT As Double, ByVal sFlg_Regular As String, ByVal dPre_Cliente As Double, ByVal sFec_RecCliLOT As String, ByVal sCod_DivPRe As String, ByVal sDes_General As String)
On Error GoTo UpdateDatGenPurORdErr
Dim SSQL As String

 SSQL = "exec SM_TG_LotEstUpdateDatGen  '$', '$', '$', '$','$' , '$' , $ , '$' ,'$' , '$',$,'$',$ ,'$' , '$','$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, sFec_DespachoOriLOT, dPrecioLOT, sFlg_Regular, dPre_Cliente, sFec_RecCliLOT, sCod_DivPRe, sDes_General)
 
 ExecuteSQL mvarConnect, SSQL

 Exit Function
UpdateDatGenPurORdErr:
  Err.Raise Err.Number, "UpdateDatGenPurORd Method", Err.Description

End Function


Public Function AddDestino(ByVal sCod_Destino As String, ByVal sDes_Destino As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_Destino '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_Destino, sDes_Destino)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddDestino = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddDestino = False
 Err.Raise Err.Number, "AddDestino Method", Err.Description
End Function

Public Function AddDivCli(ByVal sCod_Cliente As String, ByVal sCod_DivCli As String, ByVal sNom_DivCli As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_DivCli '$' , '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_DivCli, sNom_DivCli)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddDivCli = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddDivCli = False
 Err.Raise Err.Number, "AddDivCli Method", Err.Description
End Function


Public Function AddTemCli(ByVal sCod_Cliente As String, ByVal sCod_TemCli As String, ByVal sNom_TemCli As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_TemCli '$' , '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_TemCli, sNom_TemCli)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddTemCli = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddTemCli = False
 Err.Raise Err.Number, "AddTemCli Method", Err.Description
End Function




Public Function AddPagEmb(ByVal sCod_PagEmb As String, ByVal sDes_PagEmb As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_PagEmb '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_PagEmb, sDes_PagEmb)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddPagEmb = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddPagEmb = False
 Err.Raise Err.Number, "AddPagEmb Method", Err.Description
End Function


Public Function AddTipEmb(ByVal sCod_Embarque As String, ByVal sDes_Embarque As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_TipEmb '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_Embarque, sDes_Embarque)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddTipEmb = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddTipEmb = False
 Err.Raise Err.Number, "AddTipEmb Method", Err.Description
End Function


Public Function AddBanco(ByVal sCod_BAnco As String, ByVal sNom_Banco As String) As Boolean
 On Error GoTo AddErr
 #If TRANSACTION_SERVER = True Then
  Dim ctxObject As ObjectContext
  Set ctxObject = GetObjectContext()
 #End If

 Dim rs As Object
Dim SSQL As String
SSQL = "exec SI_TG_Banco '$' , '$'"
 SSQL = VBsprintf(SSQL, sCod_BAnco, sNom_Banco)
 ExecuteSQL mvarConnect, SSQL
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetComplete
 #End If
 AddBanco = True
 Exit Function
AddErr:
 #If TRANSACTION_SERVER = True Then
  ctxObject.SetAbort
 #End If
 AddBanco = False
 Err.Raise Err.Number, "AddBanco Method", Err.Description
End Function



Public Function EliminaNoSeleccionadosWithKey(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String)
On Error GoTo EliminaNoSeleccionadosWithKeyErr
Dim SSQL As String

 SSQL = "exec SG_EliminaNoSelecKey '$' ,'$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, sCod_Usuario, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 ExecuteSQL mvarConnect, SSQL

 Exit Function
EliminaNoSeleccionadosWithKeyErr:
  Err.Raise Err.Number, "EliminaNoSeleccionadosWithKeyErr Method", Err.Description
End Function



Public Function ViewColoresSELEC_Matriz(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String)
On Error GoTo ViewColoresSELEC_MatrizErr
Dim SSQL As String

 SSQL = "exec SG_MatrizColoresSelec '$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)

 ViewColoresSELEC_Matriz = GetRows()

 Exit Function
ViewColoresSELEC_MatrizErr:
  Err.Raise Err.Number, "ViewColoresSELEC_Matriz Method", Err.Description
End Function

Public Function ViewTAllasSELEC_Matriz(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String)
On Error GoTo ViewTAllasSELEC_MatrizErr
Dim SSQL As String

 SSQL = "exec SG_MatrizTallasSelec '$' ,'$' , '$'  , '$'"
 SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
 Set mRs = GetRecordset(mvarConnect, SSQL)

 ViewTAllasSELEC_Matriz = GetRows()

 Exit Function
ViewTAllasSELEC_MatrizErr:
  Err.Raise Err.Number, "ViewTAllasSELEC_Matriz Method", Err.Description
End Function




