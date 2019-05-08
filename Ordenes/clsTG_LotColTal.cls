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

Public Function ViewMatriz(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_EstCli As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mCod_TemCli As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String

 ssql = "exec SG_GeneraMatriz '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$'"
 ssql = VBsprintf(ssql, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_EstCli, mCod_Fabrica, mCod_Destino, mFec_Despacho, mCod_TemCli)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewMatriz = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function


Public Function ViewMatrizKey(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String, ByVal mCod_TemCli As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String

 ssql = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$' ,'$'"
 ssql = VBsprintf(ssql, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli, mCod_TemCli)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewMatrizKey = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function

Public Function ViewMatrizKeyUpdate(ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String

 ssql = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$' , $"
 ssql = VBsprintf(ssql, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli, 0)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewMatrizKeyUpdate = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatrizKeyUpdate Method", Err.Description
End Function

Public Function ViewVectorColorKey(ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String

 ssql = "exec SG_GeneraVectorKey '$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewVectorColorKey = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewVectorColorKey Method", Err.Description
End Function


Public Function ViewMatrizDif(ByVal iSecuencia As Integer, ByVal mCod_Usuario As String, ByVal mCod_Cliente As String, ByVal mCod_PurOrd As String, ByVal mCod_LotPurOrd As String, ByVal mCod_EstCli As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String

 ssql = "exec SG_GeneraMatrizDif $, '$' ,'$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, iSecuencia, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewMatrizDif = RowsDataSet()


 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function



Public Function LimpiaMatrizKeyEnTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo LimpiaMatrizKeyEnTemporalErr
Dim ssql As String

 ssql = "exec SG_LimpiaMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 ExecuteCommandSQL mvarConnect, ssql
 


 Exit Function
LimpiaMatrizKeyEnTemporalErr:
  Err.Raise Err.Number, "LimpiaMatrizKeyEnTemporal Method", Err.Description
End Function

'--------------------------
'Funcion creada por AHSP
'--------------------------
Public Function LimpiaTodoMatrizKeyEnTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo LimpiaMatrizKeyEnTemporalErr
Dim ssql As String

 ssql = "exec SG_LimpiaTodoMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 ExecuteCommandSQL mvarConnect, ssql
 


 Exit Function
LimpiaMatrizKeyEnTemporalErr:
  Err.Raise Err.Number, "LimpiaMatrizKeyEnTemporal Method", Err.Description
End Function

Public Function SaveToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mPrecio As Double, ByVal mWithKey As Integer, ByVal mCod_DivPRe As String)
On Error GoTo SaveToTemporalErr
Dim ssql As String

 ssql = "exec SI_TM_LotColTal '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' ,$ , $ , '$'"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mWithKey, mCod_DivPRe)
 ExecuteCommandSQL mvarConnect, ssql
 


 Exit Function
SaveToTemporalErr:
  Err.Raise Err.Number, "SaveToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mNum_PreReq As Long)
On Error GoTo SaveCantidadesToTemporalErr
Dim ssql As String

 ssql = "exec SI_TM_LotColTalCantidadesRequeridas '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $ "
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mNum_PreReq)
 ExecuteCommandSQL mvarConnect, ssql
 


 Exit Function
SaveCantidadesToTemporalErr:
  Err.Raise Err.Number, "SaveCantidadesToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal2(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mNum_PreDesp As Long, ByVal iSecuencia As Integer)
On Error GoTo SaveCantidadesToTemporalErr
Dim ssql As String

 ssql = "exec SI_Actualiza_LotCotTal '$','$','$','$','$','$','$',$,$"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mNum_PreDesp, iSecuencia)
 ExecuteCommandSQL mvarConnect, ssql

 Exit Function
SaveCantidadesToTemporalErr:
  Err.Raise Err.Number, "SaveCantidadesToTemporal2 Method", Err.Description
End Function

Public Function SavePreciosToTemporal(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_ColCli As String, ByVal sCod_Talla As String, ByVal mCod_Fabrica As String, ByVal mCod_Destino As String, ByVal mFec_Despacho As String, ByVal mPrecio As Double, ByVal mCod_DivPRe As String)
On Error GoTo SavePreciosToTemporalErr
Dim ssql As String

 ssql = "exec SI_TM_LotColTalPrecios '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $  , '$'"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_ColCli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mCod_DivPRe)
 ExecuteCommandSQL mvarConnect, ssql


 Exit Function
SavePreciosToTemporalErr:
  Err.Raise Err.Number, "SavePreciosToTemporal Method", Err.Description
End Function


Public Function GenerarInformacion(ByVal sModalUpdate As String, ByVal sCod_Usuario As String, _
ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_ClaPurOrd As String, _
ByVal sFec_DespachoAct As String, ByVal sCod_PagEmb As String, ByVal sCod_Fabrica As String, _
ByVal sCod_Destino As String, ByVal sCod_Embarque As String, ByVal sCod_Moneda As String, _
ByVal sCod_DivCli As String, ByVal sCod_TemCli As String, ByVal sFlg_Carta As String, _
ByVal sCod_BAnco As String, ByVal dPor_Slush As Double, ByVal sDes_General As String, _
ByVal sDes_Despacho As String, ByVal dPor_Comision As Double, ByVal sCod_LotPurOrd As String, _
ByVal sCod_FabricaLot As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, _
ByVal sCod_EstCli As String, ByVal dPor_ComisionLot As Double, ByVal dPrecioLOT As Double, _
ByVal sCod_UsuUltAct As String, ByVal sCod_EstUltAct As String, ByVal sFlg_OrdenREorden As String, _
ByVal sFlg_Regular As String, Por_AdicProd As Double, Pre_AdicProd As Integer, Num_PreCri As Integer, Cod_GrupoPro As String, sFlg_ComisionEnPorcentaje As String, dImp_Comision As Double, ByVal sFec_Emision As String, ByVal sFec_LlegadaPO) As String
Dim mRs As ADODB.Recordset

'ModalUpdate , Cod_Usuario, Cod_Cliente, Cod_PurOrd, Cod_ClaPurOrd, Fec_DespachoAct, Cod_PagEmb, Cod_Fabrica, Cod_Destino, Cod_Embarque, Cod_Moneda, Cod_DivCli, Cod_TemCli, Flg_Carta, Cod_Banco, Por_Slush, Des_General, Des_Despacho, Por_Comision, Cod_LotPurOrd, Cod_FabricaLot, Cod_DestinoLot, Fec_DespachoActLot, Cod_EstCli, Por_ComisionLot, PrecioLot, Cod_UsuUltAct, Cod_EstUltAct
On Error GoTo GenerarInformacionErr
Dim ssql As String

 ssql = "exec SM_TG_LotColTalPrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $, '$', '$', '$', '$', '$', $, $, '$', '$','$','$','$','$','$','$','$',$,'$','$'"
 ssql = VBsprintf(ssql, sModalUpdate, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, sCod_LotPurOrd, sCod_FabricaLot, sCod_DestinoLOT, sFec_DespachoActLOT, sCod_EstCli, dPor_ComisionLot, dPrecioLOT, sCod_UsuUltAct, sCod_EstUltAct, sFlg_OrdenREorden, sFlg_Regular, _
 Por_AdicProd, Pre_AdicProd, Num_PreCri, Cod_GrupoPro, sFlg_ComisionEnPorcentaje, dImp_Comision, sFec_Emision, sFec_LlegadaPO)
 
 Set mRs = GetDataSet(mvarConnect, ssql)
 If Not mRs Is Nothing Then
    If Not mRs.EOF Then
        GenerarInformacion = FixNulos(mRs!cod_purord, vbstring)
    End If
    mRs.Close
 End If
 Set mRs = Nothing
 Exit Function
GenerarInformacionErr:
  Err.Raise Err.Number, "GenerarInformacion Method", Err.Description
  GenerarInformacion = "-1"
  
  If Not mRs Is Nothing Then
    mRs.Close
    Set mRs = Nothing
  End If
End Function

Public Function UpdateInformacion(ByVal sModalUpdate As String, ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal mCod_ClaPurOrd As String, ByVal mCod_Moneda As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal dPor_ComisionLot As Double, ByVal sCod_UsuUltAct As String, ByVal sCod_EstUltAct As String, sFlg_ComisionEnPorcentaje As String, dImp_Comision As Double)
On Error GoTo GenerarInformacionErr
Dim ssql As String

 ssql = "exec SM_TG_LotColTalUpdate  '$', '$', '$', '$', '$' , '$', '$', '$', '$', $, '$', '$','$',$"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, mCod_ClaPurOrd, mCod_Moneda, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_ComisionLot, sCod_UsuUltAct, sCod_EstUltAct, sFlg_ComisionEnPorcentaje, dImp_Comision)
 ExecuteCommandSQL mvarConnect, ssql


 Exit Function
GenerarInformacionErr:
  Err.Raise Err.Number, "GenerarInformacion Method", Err.Description
End Function

Public Function UpdatePOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_ClaPurOrd As String, ByVal sFec_DespachoAct As String, ByVal sCod_PagEmb As String, ByVal sCod_Fabrica As String, ByVal sCod_Destino As String, ByVal sCod_Embarque As String, ByVal sCod_Moneda As String, ByVal sCod_DivCli As String, ByVal sCod_TemCli As String, ByVal sFlg_Carta As String, ByVal sCod_BAnco As String, ByVal dPor_Slush As Double, ByVal sDes_General As String, ByVal sDes_Despacho As String, ByVal dPor_Comision As Double, ByVal sFlg_OrdenREorden As String, ByVal sFlg_Regular As String, Por_AdicProd As Double, Pre_AdicProd As Integer, Num_PreCri As Integer, Cod_GrupoPro As String, Cod_Usuario As String, Fec_Emision As String, Fec_LlegadaPO As String)

On Error GoTo UpdatePOCErr
Dim ssql As String

 ssql = "exec SM_TG_PurOrd_AddUpdatePrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $ ,'$' ,'$' ,'$','$','$','$','$','$','$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, sFlg_OrdenREorden, sFlg_Regular, Por_AdicProd, Pre_AdicProd, Num_PreCri, Cod_GrupoPro, Cod_Usuario, Fec_Emision, Fec_LlegadaPO)
 
 ExecuteCommandSQL mvarConnect, ssql


 Exit Function
UpdatePOCErr:
  Err.Raise Err.Number, "UpdatePOC Method", Err.Description
End Function


Public Function DeletePOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String)
On Error GoTo DeletePOCErr
Dim ssql As String

 ssql = "exec SM_TG_PurOrd_DeletePrincipal '$', '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd)
 ExecuteCommandSQL mvarConnect, ssql


 Exit Function
DeletePOCErr:
  Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function

Public Function DeleteLotPurOrd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd)
On Error GoTo DeletePOCErr
Dim ssql As String

 ssql = "exec SM_TG_LotPurOrd_DeletePrincipal '$', '$' , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd)
 ExecuteCommandSQL mvarConnect, ssql


 Exit Function
DeletePOCErr:
  Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function



Public Function LoadDataPOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_Usuario As String) As Variant

On Error GoTo LoadDataPOCErr
Dim ssql As String

 ssql = "exec SM_TG_PurOrd_ViewDetalle '$', '$', '$', '$', '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, "", "", sCod_Usuario)
 Set mRs = GetDataSet(mvarConnect, ssql)
 LoadDataPOC = RowsDataSet()



 Exit Function
LoadDataPOCErr:

  Err.Raise Err.Number, "LoadDataPOC Method", Err.Description
End Function

Public Function LoadDataPOC_mp(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String) As Variant

On Error GoTo LoadDataPOCErr
Dim ssql As String

 ssql = "exec SM_DATOS_ADICIONALES '$','$', '$', '$', '$','$','$','$','$'"
 ssql = VBsprintf(ssql, "", sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, "", "", "", "")
 Set mRs = GetDataSet(mvarConnect, ssql)
 LoadDataPOC_mp = RowsDataSet()



 Exit Function
LoadDataPOCErr:

  Err.Raise Err.Number, "LoadDataPOC Method", Err.Description
End Function


Public Function LoadDataLOTEST(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String) As Variant

On Error GoTo LoadDataLOTESTErr
Dim ssql As String

 ssql = "exec SM_TG_LOTEST_ViewDATA '$', '$', '$', '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)
 LoadDataLOTEST = RowsDataSet()



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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_DivPre '$' , '$'"
 ssql = VBsprintf(ssql, sCod_DivPRe, sDes_DivPRe)
 ExecuteCommandSQL mvarConnect, ssql
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
Dim ssql As String
 ssql = "exec SM_TG_DivPre_View '$' "
 ssql = VBsprintf(ssql, mCod_DivPRe)
 Set mRs = GetDataSet(mvarConnect, ssql)

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
Dim ssql As String
 ssql = "exec SM_TG_DivPre_ViewShort '$' "
 ssql = VBsprintf(ssql, mCod_DivPRe)
 Set mRs = GetDataSet(mvarConnect, ssql)
 ViewDivPre = RowsDataSet

 Exit Function
ViewAllErr:
  Err.Raise Err.Number, "ViewDivPre Method", Err.Description
End Function



Public Function VAlidFechaDespacho(ByVal mFecha As String) As Variant
On Error GoTo ViewAllErr
Dim ssql As String
 ssql = "exec SG_TG_RCtrlVAlidFEcha '$' "
 ssql = VBsprintf(ssql, mFecha)
 Set mRs = GetDataSet(mvarConnect, ssql)
 If Not mRs.EOF Then
    VAlidFechaDespacho = FixNulos(mRs!Ret, vbInteger)
 End If
 Exit Function
ViewAllErr:
  VAlidFechaDespacho = -1
  Err.Raise Err.Number, "ValidFechaDespacho Method", Err.Description
End Function


Public Function UpdateDatGenPurORd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_DestinoLOT As String, ByVal dPor_Comision As Double, ByVal sCod_Usuario As String, ByVal sCod_EStacion As String, _
                                   ByVal sFec_DespachoOriLOT As String, ByVal dPrecioLOT As Double, ByVal sFlg_Regular As String, ByVal dPre_Cliente As Double, _
                                   ByVal sFec_RecCliLOT As String, ByVal sCod_DivPRe As String, ByVal sDes_General As String, sFlg_ComisionEnPorcentaje As String, dImp_Comision As Double, _
                                   ByVal dPrecioCosto As Double, _
                                   ByVal dblPreCost_Tela As Double, _
                                   ByVal dblPreCost_Avios As Double, _
                                   ByVal dblPreCost_MOD As Double, _
                                   ByVal dblPreCost_FOB As Double, _
                                   ByVal dblPreCost_Financieros As Double, _
                                   ByVal dblPreCost_Fijos As Double, _
                                   ByVal dFecAceptacionDelCliente As Date)
On Error GoTo UpdateDatGenPurORdErr
Dim ssql As String


 'sSql = "exec SM_TG_LotEstUpdateDatGen  '$', '$', '$', '$','$' , '$' , $ , '$' ,'$' , '$',$,'$',$ ,'$' , '$','$','$','$',$"
 'sSql = VBsprintf(sSql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, sFec_DespachoOriLOT, dPrecioLOT, sFlg_Regular, dPre_Cliente, sFec_RecCliLOT, sCod_DivPRe, sDes_General, vusu, sFlg_ComisionEnPorcentaje, dImp_Comision)
 
 ssql = "exec SM_TG_LotEstUpdateDatGen_NEW  '$', '$', '$', '$', '$', $, '$', '$', '$', $, '$', $, '$', '$', '$', '$', '$', $, '$', $, $, $, $, $, $, $, '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_DestinoLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, _
                        sFec_DespachoOriLOT, dPrecioLOT, sFlg_Regular, dPre_Cliente, _
                        sFec_RecCliLOT, sCod_DivPRe, sDes_General, vusu, sFlg_ComisionEnPorcentaje, dImp_Comision, "S", _
                        dPrecioCosto, _
                        dblPreCost_Tela, _
                        dblPreCost_Avios, _
                        dblPreCost_MOD, _
                        dblPreCost_FOB, _
                        dblPreCost_Financieros, _
                        dblPreCost_Fijos, _
                        CStr(Format(dFecAceptacionDelCliente, "dd/MM/yyyy HH:mm")) _
                )
                  
 
 ExecuteCommandSQL mvarConnect, ssql

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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_Destino '$' , '$'"
 ssql = VBsprintf(ssql, sCod_Destino, sDes_Destino)
 ExecuteCommandSQL mvarConnect, ssql
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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_DivCli '$' , '$' , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_DivCli, sNom_DivCli)
 ExecuteCommandSQL mvarConnect, ssql
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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_TemCli '$' , '$' , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_TemCli, sNom_TemCli)
 ExecuteCommandSQL mvarConnect, ssql
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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_PagEmb '$' , '$'"
 ssql = VBsprintf(ssql, sCod_PagEmb, sDes_PagEmb)
 ExecuteCommandSQL mvarConnect, ssql
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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_TipEmb '$' , '$'"
 ssql = VBsprintf(ssql, sCod_Embarque, sDes_Embarque)
 ExecuteCommandSQL mvarConnect, ssql
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

 Dim Rs As Object
Dim ssql As String
ssql = "exec SI_TG_Banco '$' , '$'"
 ssql = VBsprintf(ssql, sCod_BAnco, sNom_Banco)
 ExecuteCommandSQL mvarConnect, ssql
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



Public Function EliminaNoSeleccionadosWithKey(ByVal sCod_Usuario As String, ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo EliminaNoSeleccionadosWithKeyErr
Dim ssql As String

 ssql = "exec SG_EliminaNoSelecKey '$' ,'$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 ExecuteCommandSQL mvarConnect, ssql

 Exit Function
EliminaNoSeleccionadosWithKeyErr:
  Err.Raise Err.Number, "EliminaNoSeleccionadosWithKeyErr Method", Err.Description
End Function

Public Function ViewColoresSELEC_Matriz(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo ViewColoresSELEC_MatrizErr
Dim ssql As String

 ssql = "exec SG_MatrizColoresSelec '$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)

 ViewColoresSELEC_Matriz = RowsDataSet()

 Exit Function
ViewColoresSELEC_MatrizErr:
  Err.Raise Err.Number, "ViewColoresSELEC_Matriz Method", Err.Description
End Function

Public Function ViewTAllasSELEC_Matriz(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo ViewTAllasSELEC_MatrizErr
Dim ssql As String

 ssql = "exec SG_MatrizTallasSelec '$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)

 ViewTAllasSELEC_Matriz = RowsDataSet()

 Exit Function
ViewTAllasSELEC_MatrizErr:
  Err.Raise Err.Number, "ViewTAllasSELEC_Matriz Method", Err.Description
End Function



Public Function InhabilitaModifCantidades(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String)
On Error GoTo InhabilitaModifCantidadesErr
Dim ssql As String

 ssql = "exec SG_VIEWTG_LotColTal_Destinos_Empaque  '$' ,'$' , '$'  , '$'"
 ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
 Set mRs = GetDataSet(mvarConnect, ssql)

 InhabilitaModifCantidades = RowsDataSet()

 Exit Function
InhabilitaModifCantidadesErr:
  Err.Raise Err.Number, "InhabilitaModifCantidades Method", Err.Description
End Function

