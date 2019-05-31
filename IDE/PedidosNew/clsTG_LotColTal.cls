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

Public Function ViewMatriz(ByVal mCod_Usuario As String, _
                           ByVal mCod_Cliente As String, _
                           ByVal mCod_PurOrd As String, _
                           ByVal mCod_EstCli As String, _
                           ByVal mCod_Fabrica As String, _
                           ByVal mCod_Destino As String, _
                           ByVal mFec_Despacho As String, _
                           ByVal mCod_TemCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_GeneraMatriz '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$'"
    sSQl = VBsprintf(sSQl, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_EstCli, mCod_Fabrica, mCod_Destino, mFec_Despacho, mCod_TemCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewMatriz = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function

Public Function ViewMatrizKey(ByVal mCod_Usuario As String, _
                              ByVal mCod_Cliente As String, _
                              ByVal mCod_PurOrd As String, _
                              ByVal mCod_LotPurOrd As String, _
                              ByVal mCod_EstCli As String, _
                              ByVal mCod_TemCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$' ,'$'"
    sSQl = VBsprintf(sSQl, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli, mCod_TemCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewMatrizKey = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function

Public Function ViewMatrizKeyUpdate(ByVal mCod_Usuario As String, _
                                    ByVal mCod_Cliente As String, _
                                    ByVal mCod_PurOrd As String, _
                                    ByVal mCod_LotPurOrd As String, _
                                    ByVal mCod_EstCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_GeneraMatrizKey '$' ,'$' ,'$' , '$'  , '$' , $"
    sSQl = VBsprintf(sSQl, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli, 0)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewMatrizKeyUpdate = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewMatrizKeyUpdate Method", Err.Description
End Function

Public Function ViewVectorColorKey(ByVal mCod_Cliente As String, _
                                   ByVal mCod_PurOrd As String, _
                                   ByVal mCod_LotPurOrd As String, _
                                   ByVal mCod_EstCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_GeneraVectorKey '$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewVectorColorKey = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewVectorColorKey Method", Err.Description
End Function

Public Function ViewMatrizDif(ByVal iSecuencia As Integer, _
                              ByVal mCod_Usuario As String, _
                              ByVal mCod_Cliente As String, _
                              ByVal mCod_PurOrd As String, _
                              ByVal mCod_LotPurOrd As String, _
                              ByVal mCod_EstCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_GeneraMatrizDif $, '$' ,'$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, iSecuencia, mCod_Usuario, mCod_Cliente, mCod_PurOrd, mCod_LotPurOrd, mCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewMatrizDif = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewMatriz Method", Err.Description
End Function

Public Function LimpiaMatrizKeyEnTemporal(ByVal sCod_Usuario As String, _
                                          ByVal sCod_Cliente As String, _
                                          ByVal sCod_PurOrd As String, _
                                          ByVal sCod_LotPurOrd As String, _
                                          ByVal sCod_EstCli As String)

    On Error GoTo LimpiaMatrizKeyEnTemporalErr

    Dim sSQl As String

    sSQl = "exec SG_LimpiaMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

LimpiaMatrizKeyEnTemporalErr:
    Err.Raise Err.Number, "LimpiaMatrizKeyEnTemporal Method", Err.Description
End Function

'--------------------------
'Funcion creada por AHSP
'--------------------------
Public Function LimpiaTodoMatrizKeyEnTemporal(ByVal sCod_Usuario As String, _
                                              ByVal sCod_Cliente As String, _
                                              ByVal sCod_PurOrd As String, _
                                              ByVal sCod_LotPurOrd As String, _
                                              ByVal sCod_EstCli As String)

    On Error GoTo LimpiaMatrizKeyEnTemporalErr

    Dim sSQl As String

    sSQl = "exec SG_LimpiaTodoMatrizKey '$' ,'$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

LimpiaMatrizKeyEnTemporalErr:
    Err.Raise Err.Number, "LimpiaMatrizKeyEnTemporal Method", Err.Description
End Function

Public Function SaveToTemporal(ByVal sCod_Usuario As String, _
                               ByVal sCod_Cliente As String, _
                               ByVal sCod_PurOrd As String, _
                               ByVal sCod_LotPurOrd As String, _
                               ByVal sCod_EstCli As String, _
                               ByVal scod_colcli As String, _
                               ByVal sCod_Talla As String, _
                               ByVal mCod_Fabrica As String, _
                               ByVal mCod_Destino As String, _
                               ByVal mFec_Despacho As String, _
                               ByVal mPrecio As Double, _
                               ByVal mWithKey As Integer, _
                               ByVal mCod_DivPRe As String)

    On Error GoTo SaveToTemporalErr

    Dim sSQl As String

    sSQl = "exec SI_TM_LotColTal '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' ,$ , $ , '$'"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, scod_colcli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mWithKey, mCod_DivPRe)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

SaveToTemporalErr:
    Err.Raise Err.Number, "SaveToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal(ByVal sCod_Usuario As String, _
                                         ByVal sCod_Cliente As String, _
                                         ByVal sCod_PurOrd As String, _
                                         ByVal sCod_LotPurOrd As String, _
                                         ByVal sCod_EstCli As String, _
                                         ByVal scod_colcli As String, _
                                         ByVal sCod_Talla As String, _
                                         ByVal mCod_Fabrica As String, _
                                         ByVal mCod_Destino As String, _
                                         ByVal mFec_Despacho As String, _
                                         ByVal mNum_PreReq As Long)

    On Error GoTo SaveCantidadesToTemporalErr

    Dim sSQl As String

    sSQl = "exec SI_TM_LotColTalCantidadesRequeridas '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $ "
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, scod_colcli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mNum_PreReq)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

SaveCantidadesToTemporalErr:
    Err.Raise Err.Number, "SaveCantidadesToTemporal Method", Err.Description
End Function

Public Function SaveCantidadesToTemporal2(ByVal sCod_Cliente As String, _
                                          ByVal sCod_PurOrd As String, _
                                          ByVal sCod_LotPurOrd As String, _
                                          ByVal sCod_EstCli As String, _
                                          ByVal scod_colcli As String, _
                                          ByVal sCod_Talla As String, _
                                          ByVal mCod_Fabrica As String, _
                                          ByVal mNum_PreDesp As Long, _
                                          ByVal iSecuencia As Integer)

    On Error GoTo SaveCantidadesToTemporalErr

    Dim sSQl As String

    sSQl = "exec SI_Actualiza_LotCotTal '$','$','$','$','$','$','$',$,$"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, scod_colcli, sCod_Talla, mCod_Fabrica, mNum_PreDesp, iSecuencia)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

SaveCantidadesToTemporalErr:
    Err.Raise Err.Number, "SaveCantidadesToTemporal2 Method", Err.Description
End Function

Public Function SavePreciosToTemporal(ByVal sCod_Usuario As String, _
                                      ByVal sCod_Cliente As String, _
                                      ByVal sCod_PurOrd As String, _
                                      ByVal sCod_LotPurOrd As String, _
                                      ByVal sCod_EstCli As String, _
                                      ByVal scod_colcli As String, _
                                      ByVal sCod_Talla As String, _
                                      ByVal mCod_Fabrica As String, _
                                      ByVal mCod_Destino As String, _
                                      ByVal mFec_Despacho As String, _
                                      ByVal mPrecio As Double, _
                                      ByVal mCod_DivPRe As String)

    On Error GoTo SavePreciosToTemporalErr

    Dim sSQl As String

    sSQl = "exec SI_TM_LotColTalPrecios '$' ,'$' ,'$' , '$'  , '$', '$' ,'$','$','$','$' , $  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, scod_colcli, sCod_Talla, mCod_Fabrica, mCod_Destino, mFec_Despacho, mPrecio, mCod_DivPRe)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

SavePreciosToTemporalErr:
    Err.Raise Err.Number, "SavePreciosToTemporal Method", Err.Description
End Function

Public Function GenerarInformacion(ByVal sModalUpdate As String, _
                                   ByVal sCod_Usuario As String, _
                                   ByVal sCod_Cliente As String, _
                                   ByVal sCod_PurOrd As String, _
                                   ByVal sCod_ClaPurOrd As String, _
                                   ByVal sFec_DespachoAct As String, _
                                   ByVal sCod_PagEmb As String, _
                                   ByVal sCod_Fabrica As String, _
                                   ByVal sCod_Destino As String, _
                                   ByVal sCod_Embarque As String, _
                                   ByVal sCod_Moneda As String, _
                                   ByVal sCod_DivCli As String, _
                                   ByVal sCod_TemCli As String, _
                                   ByVal sFlg_Carta As String, _
                                   ByVal sCod_BAnco As String, _
                                   ByVal dPor_Slush As Double, _
                                   ByVal sDes_General As String, _
                                   ByVal sDes_Despacho As String, _
                                   ByVal dPor_Comision As Double, _
                                   ByVal sCod_LotPurOrd As String, _
                                   ByVal sCod_FabricaLot As String, _
                                   ByVal sCod_DestinoLOT As String, _
                                   ByVal sFec_DespachoActLOT As String, _
                                   ByVal sCod_EstCli As String, _
                                   ByVal dPor_ComisionLot As Double, ByVal dPrecioLOT As Double, ByVal sCod_UsuUltAct As String, ByVal sCod_EstUltAct As String, ByVal sFlg_OrdenREorden As String, ByVal sFlg_Regular As String, Por_AdicProd As Double, Pre_AdicProd As Integer, Num_PreCri As Integer, Cod_GrupoPro As String, sFlg_ComisionEnPorcentaje As String, dImp_Comision As Double, ByVal sFec_Emision As String, ByVal sFec_LlegadaPO) As String

    Dim mRs As ADODB.Recordset

    'ModalUpdate , Cod_Usuario, Cod_Cliente, Cod_PurOrd, Cod_ClaPurOrd, Fec_DespachoAct, Cod_PagEmb, Cod_Fabrica, Cod_Destino, Cod_Embarque, Cod_Moneda, Cod_DivCli, Cod_TemCli, Flg_Carta, Cod_Banco, Por_Slush, Des_General, Des_Despacho, Por_Comision, Cod_LotPurOrd, Cod_FabricaLot, Cod_DestinoLot, Fec_DespachoActLot, Cod_EstCli, Por_ComisionLot, PrecioLot, Cod_UsuUltAct, Cod_EstUltAct
    On Error GoTo GenerarInformacionErr

    Dim sSQl As String

    sSQl = "exec SM_TG_LotColTalPrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $, '$', '$', '$', '$', '$', $, $, '$', '$','$','$','$','$','$','$','$',$,'$','$'"
    sSQl = VBsprintf(sSQl, sModalUpdate, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, sCod_LotPurOrd, sCod_FabricaLot, sCod_DestinoLOT, sFec_DespachoActLOT, sCod_EstCli, dPor_ComisionLot, dPrecioLOT, sCod_UsuUltAct, sCod_EstUltAct, sFlg_OrdenREorden, sFlg_Regular, Por_AdicProd, Pre_AdicProd, Num_PreCri, Cod_GrupoPro, sFlg_ComisionEnPorcentaje, dImp_Comision, sFec_Emision, sFec_LlegadaPO)
 
    Set mRs = GetDataSet(mvarConnect, sSQl)

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

Public Function UpdateInformacion(ByVal sModalUpdate As String, _
                                  ByVal sCod_Usuario As String, _
                                  ByVal sCod_Cliente As String, _
                                  ByVal sCod_PurOrd As String, _
                                  ByVal sCod_LotPurOrd As String, _
                                  ByVal sCod_EstCli As String, _
                                  ByVal mCod_ClaPurOrd As String, _
                                  ByVal mCod_Moneda As String, _
                                  ByVal sCod_DestinoLOT As String, _
                                  ByVal sFec_DespachoActLOT As String, _
                                  ByVal dPor_ComisionLot As Double, _
                                  ByVal sCod_UsuUltAct As String, _
                                  ByVal sCod_EstUltAct As String, _
                                  sFlg_ComisionEnPorcentaje As String, _
                                  dImp_Comision As Double)

    On Error GoTo GenerarInformacionErr

    Dim sSQl As String

    sSQl = "exec SM_TG_LotColTalUpdate  '$', '$', '$', '$', '$' , '$', '$', '$', '$', $, '$', '$','$',$"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, mCod_ClaPurOrd, mCod_Moneda, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_ComisionLot, sCod_UsuUltAct, sCod_EstUltAct, sFlg_ComisionEnPorcentaje, dImp_Comision)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

GenerarInformacionErr:
    Err.Raise Err.Number, "GenerarInformacion Method", Err.Description
End Function

Public Function UpdatePOC(ByVal sCod_Cliente As String, _
                          ByVal sCod_PurOrd As String, _
                          ByVal sCod_ClaPurOrd As String, _
                          ByVal sFec_DespachoAct As String, _
                          ByVal sCod_PagEmb As String, _
                          ByVal sCod_Fabrica As String, _
                          ByVal sCod_Destino As String, _
                          ByVal sCod_Embarque As String, _
                          ByVal sCod_Moneda As String, _
                          ByVal sCod_DivCli As String, _
                          ByVal sCod_TemCli As String, _
                          ByVal sFlg_Carta As String, _
                          ByVal sCod_BAnco As String, _
                          ByVal dPor_Slush As Double, _
                          ByVal sDes_General As String, _
                          ByVal sDes_Despacho As String, _
                          ByVal dPor_Comision As Double, _
                          ByVal sFlg_OrdenREorden As String, _
                          ByVal sFlg_Regular As String, _
                          Por_AdicProd As Double, _
                          Pre_AdicProd As Integer, _
                          Num_PreCri As Integer, _
                          Cod_GrupoPro As String, _
                          Cod_Usuario As String, _
                          Fec_Emision As String, Fec_LlegadaPO As String, Fec_ProyectadaPO As String)

    On Error GoTo UpdatePOCErr

    Dim sSQl As String

    sSQl = "exec SM_TG_PurOrd_AddUpdatePrincipal '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', '$', $, '$', '$', $ ,'$' ,'$' ,'$','$','$','$','$','$','$','$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_ClaPurOrd, sFec_DespachoAct, sCod_PagEmb, sCod_Fabrica, sCod_Destino, sCod_Embarque, sCod_Moneda, sCod_DivCli, sCod_TemCli, sFlg_Carta, sCod_BAnco, dPor_Slush, sDes_General, sDes_Despacho, dPor_Comision, sFlg_OrdenREorden, sFlg_Regular, Por_AdicProd, Pre_AdicProd, Num_PreCri, Cod_GrupoPro, Cod_Usuario, Fec_Emision, Fec_LlegadaPO, Fec_ProyectadaPO)
 
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

UpdatePOCErr:
    Err.Raise Err.Number, "UpdatePOC Method", Err.Description
End Function

Public Function DeletePOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String)

    On Error GoTo DeletePOCErr

    Dim sSQl As String

    sSQl = "exec SM_TG_PurOrd_DeletePrincipal '$', '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

DeletePOCErr:
    Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function

Public Function DeleteLotPurOrd(ByVal sCod_Cliente As String, _
                                ByVal sCod_PurOrd As String, _
                                ByVal sCod_LotPurOrd)

    On Error GoTo DeletePOCErr

    Dim sSQl As String

    sSQl = "exec SM_TG_LotPurOrd_DeletePrincipal '$', '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

DeletePOCErr:
    Err.Raise Err.Number, "DeletePOCErr Method", Err.Description
End Function

Public Function LoadDataPOC(ByVal sCod_Cliente As String, _
                            ByVal sCod_PurOrd As String, _
                            ByVal sCod_Usuario As String) As Variant

    On Error GoTo LoadDataPOCErr

    Dim sSQl As String

    sSQl = "exec SM_TG_PurOrd_ViewDetalle '$', '$', '$', '$', '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, "", "", sCod_Usuario)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    LoadDataPOC = RowsDataSet()

    Exit Function

LoadDataPOCErr:

    Err.Raise Err.Number, "LoadDataPOC Method", Err.Description
End Function

Public Function LoadDataPOC_mp(ByVal sCod_Cliente As String, _
                               ByVal sCod_PurOrd As String, _
                               ByVal sCod_LotPurOrd As String, _
                               ByVal sCod_EstCli As String) As Variant

    On Error GoTo LoadDataPOCErr

    Dim sSQl As String

    sSQl = "exec SM_DATOS_ADICIONALES '$','$', '$', '$', '$','$','$','$','$'"
    sSQl = VBsprintf(sSQl, "", sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, "", "", "", "")
    Set mRs = GetDataSet(mvarConnect, sSQl)
    LoadDataPOC_mp = RowsDataSet()

    Exit Function

LoadDataPOCErr:

    Err.Raise Err.Number, "LoadDataPOC Method", Err.Description
End Function

Public Function LoadDataLOTEST(ByVal sCod_Cliente As String, _
                               ByVal sCod_PurOrd As String, _
                               ByVal sCod_LotPurOrd As String, _
                               ByVal sCod_EstCli As String) As Variant

    On Error GoTo LoadDataLOTESTErr

    Dim sSQl As String

    sSQl = "exec SM_TG_LOTEST_ViewDATA '$', '$', '$', '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    LoadDataLOTEST = RowsDataSet()

    Exit Function

LoadDataLOTESTErr:

    Err.Raise Err.Number, "LoadDataLOTEST Method", Err.Description
End Function

Public Function AddDivPre(ByVal sCod_DivPRe As String, _
                          ByVal sDes_DivPRe As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_DivPre '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_DivPRe, sDes_DivPRe)
    ExecuteCommandSQL mvarConnect, sSQl
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

    Dim sSQl As String

    sSQl = "exec SM_TG_DivPre_View '$' "
    sSQl = VBsprintf(sSQl, mCod_DivPRe)
    Set mRs = GetDataSet(mvarConnect, sSQl)

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

    Dim sSQl As String

    sSQl = "exec SM_TG_DivPre_ViewShort '$' "
    sSQl = VBsprintf(sSQl, mCod_DivPRe)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewDivPre = RowsDataSet

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewDivPre Method", Err.Description
End Function

Public Function VAlidFechaDespacho(ByVal mFecha As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SG_TG_RCtrlVAlidFEcha '$' "
    sSQl = VBsprintf(sSQl, mFecha)
    Set mRs = GetDataSet(mvarConnect, sSQl)

    If Not mRs.EOF Then
        VAlidFechaDespacho = FixNulos(mRs!Ret, vbInteger)
    End If

    Exit Function

ViewAllErr:
    VAlidFechaDespacho = -1
    Err.Raise Err.Number, "ValidFechaDespacho Method", Err.Description
End Function

'Public Function UpdateDatGenPurORd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurORd As String, ByVal sCod_EstCli As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal dPor_Comision As Double, ByVal sCod_Usuario As String, ByVal sCod_EStacion As String, ByVal dPrecioLot As Double, ByVal sCod_DivPre As String)
'On Error GoTo UpdateDatGenPurORdErr
'Dim SSQL As String
'
' SSQL = "exec SM_TG_LotEstUpdateDatGen  '$', '$', '$', '$','$' , '$' , $ , '$' ,'$' , $, '$'"
' SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, sCod_LotPurORd, sCod_EstCli, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, dPrecioLot, sCod_DivPre)
' ExecuteCommandSQL mvarConnect, SSQL
'
'
' Exit Function
'UpdateDatGenPurORdErr:
'  Err.Raise Err.Number, "UpdateDatGenPurORd Method", Err.Description
'
'End Function
'Public Function UpdateDatGenPurORd(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal sCod_LotPurOrd As String, ByVal sCod_EstCli As String, ByVal sCod_DestinoLOT As String, ByVal sFec_DespachoActLOT As String, ByVal dPor_Comision As Double, ByVal sCod_Usuario As String, ByVal sCod_EStacion As String, ByVal sFec_DespachoOriLOT As String, ByVal dPrecioLOT As Double, ByVal sFlg_Regular As String, ByVal dPre_Cliente As Double, ByVal sFec_RecCliLOT As String, ByVal sCod_DivPRe As String, ByVal sDes_General As String, sFlg_ComisionEnPorcentaje As String, dImp_Comision As Double)
Public Function UpdateDatGenPurORd(ByVal sCod_Cliente As String, _
                                   ByVal sCod_PurOrd As String, _
                                   ByVal sCod_LotPurOrd As String, _
                                   ByVal sCod_EstCli As String, _
                                   ByVal sCod_DestinoLOT As String, _
                                   ByVal dPor_Comision As Double, _
                                   ByVal sCod_Usuario As String, _
                                   ByVal sCod_EStacion As String, _
                                   ByVal sFec_DespachoOriLOT As String, _
                                   ByVal dPrecioLOT As Double, _
                                   ByVal sFlg_Regular As String, _
                                   ByVal dPre_Cliente As Double, _
                                   ByVal sFec_RecCliLOT As String, _
                                   ByVal sCod_DivPRe As String, _
                                   ByVal sDes_General As String, _
                                   sFlg_ComisionEnPorcentaje As String, _
                                   dImp_Comision As Double, _
                                   dtip_Lead_Time As String)

    On Error GoTo UpdateDatGenPurORdErr

    Dim sSQl As String

    'sSql = "exec SM_TG_LotEstUpdateDatGen  '$', '$', '$', '$','$' , '$' , $ , '$' ,'$' , '$',$,'$',$ ,'$' , '$','$','$','$',$"
    sSQl = "exec SM_TG_LotEstUpdateDatGen_NEW  '$', '$', '$', '$','$' , $ , '$' ,'$' , '$',$,'$',$ ,'$' , '$','$','$','$',$,'$','$'"
    'sSql = VBsprintf(sSql, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_DestinoLOT, sFec_DespachoActLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, sFec_DespachoOriLOT, dPrecioLOT, sFlg_Regular, dPre_Cliente, sFec_RecCliLOT, sCod_DivPRe, sDes_General, vusu, sFlg_ComisionEnPorcentaje, dImp_Comision)
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sCod_DestinoLOT, dPor_Comision, sCod_Usuario, sCod_EStacion, sFec_DespachoOriLOT, dPrecioLOT, sFlg_Regular, dPre_Cliente, sFec_RecCliLOT, sCod_DivPRe, sDes_General, vusu, sFlg_ComisionEnPorcentaje, dImp_Comision, "S", dtip_Lead_Time)
 
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

UpdateDatGenPurORdErr:
    Err.Raise Err.Number, "UpdateDatGenPurORd Method", Err.Description

End Function

Public Function AddDestino(ByVal sCod_Destino As String, _
                           ByVal sDes_Destino As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_Destino '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Destino, sDes_Destino)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function AddDivCli(ByVal sCod_Cliente As String, _
                          ByVal sCod_DivCli As String, _
                          ByVal sNom_DivCli As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_DivCli '$' , '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_DivCli, sNom_DivCli)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function AddTemCli(ByVal sCod_Cliente As String, _
                          ByVal sCod_TemCli As String, _
                          ByVal sNom_TemCli As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_TemCli '$' , '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_TemCli, sNom_TemCli)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function AddPagEmb(ByVal sCod_PagEmb As String, _
                          ByVal sDes_PagEmb As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_PagEmb '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_PagEmb, sDes_PagEmb)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function AddTipEmb(ByVal sCod_Embarque As String, _
                          ByVal sDes_Embarque As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_TipEmb '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Embarque, sDes_Embarque)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function AddBanco(ByVal sCod_BAnco As String, _
                         ByVal sNom_Banco As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_Banco '$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_BAnco, sNom_Banco)
    ExecuteCommandSQL mvarConnect, sSQl
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

Public Function EliminaNoSeleccionadosWithKey(ByVal sCod_Usuario As String, _
                                              ByVal sCod_Cliente As String, _
                                              ByVal sCod_PurOrd As String, _
                                              ByVal sCod_LotPurOrd As String, _
                                              ByVal sCod_EstCli As String)

    On Error GoTo EliminaNoSeleccionadosWithKeyErr

    Dim sSQl As String

    sSQl = "exec SG_EliminaNoSelecKey '$' ,'$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Usuario, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    ExecuteCommandSQL mvarConnect, sSQl

    Exit Function

EliminaNoSeleccionadosWithKeyErr:
    Err.Raise Err.Number, "EliminaNoSeleccionadosWithKeyErr Method", Err.Description
End Function

Public Function ViewColoresSELEC_Matriz(ByVal sCod_Cliente As String, _
                                        ByVal sCod_PurOrd As String, _
                                        ByVal sCod_LotPurOrd As String, _
                                        ByVal sCod_EstCli As String)

    On Error GoTo ViewColoresSELEC_MatrizErr

    Dim sSQl As String

    sSQl = "exec SG_MatrizColoresSelec '$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)

    ViewColoresSELEC_Matriz = RowsDataSet()

    Exit Function

ViewColoresSELEC_MatrizErr:
    Err.Raise Err.Number, "ViewColoresSELEC_Matriz Method", Err.Description
End Function

Public Function ViewTAllasSELEC_Matriz(ByVal sCod_Cliente As String, _
                                       ByVal sCod_PurOrd As String, _
                                       ByVal sCod_LotPurOrd As String, _
                                       ByVal sCod_EstCli As String)

    On Error GoTo ViewTAllasSELEC_MatrizErr

    Dim sSQl As String

    sSQl = "exec SG_MatrizTallasSelec '$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)

    ViewTAllasSELEC_Matriz = RowsDataSet()

    Exit Function

ViewTAllasSELEC_MatrizErr:
    Err.Raise Err.Number, "ViewTAllasSELEC_Matriz Method", Err.Description
End Function

Public Function InhabilitaModifCantidades(ByVal sCod_Cliente As String, _
                                          ByVal sCod_PurOrd As String, _
                                          ByVal sCod_LotPurOrd As String, _
                                          ByVal sCod_EstCli As String)

    On Error GoTo InhabilitaModifCantidadesErr

    Dim sSQl As String

    sSQl = "exec SG_VIEWTG_LotColTal_Destinos_Empaque  '$' ,'$' , '$'  , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)

    InhabilitaModifCantidades = RowsDataSet()

    Exit Function

InhabilitaModifCantidadesErr:
    Err.Raise Err.Number, "InhabilitaModifCantidades Method", Err.Description
End Function

