VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTG_Cliente"
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

Public Function ViewClientes(ByVal mCod_Cliente As String, _
                             ByVal mCod_Usuario As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_Cliente_VieShort '$' , '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_Usuario)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewClientes = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewClientes Method", Err.Description
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
    mName = "TG_Cliente"
End Sub

Public Function ViewDivClientes(ByVal mCod_Cliente As String, _
                                ByVal mCod_DivCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_DivCli_ViewShort '$' , '$' "
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_DivCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewDivClientes = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewDivClientes Method", Err.Description
End Function

Public Function ViewTempClientes(ByVal mCod_Cliente As String, _
                                 ByVal mCod_TemCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_TemCli_ViewShort '$' ,'$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_TemCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewTempClientes = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTempClientes Method", Err.Description
End Function

Public Function ViewPagEmbarque(ByVal mCod_PagEmb As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_PagEmb_ViewShort '$' "
    sSQl = VBsprintf(sSQl, mCod_PagEmb)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewPagEmbarque = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewPagEmbarque Method", Err.Description
End Function

Public Function ViewTipEmb(ByVal mCod_TipEmb As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_TipEmb_ViewShort '$' "
    sSQl = VBsprintf(sSQl, mCod_TipEmb)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewTipEmb = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewTipEmb Method", Err.Description
End Function

Public Function ViewMoneda(ByVal mCod_Moneda As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_Moneda_ViewShort '$' "
    sSQl = VBsprintf(sSQl, mCod_Moneda)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewMoneda = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewMoneda Method", Err.Description
End Function

Public Function ViewBanco(ByVal mCod_Banco As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_Banco_ViewShort '$' "
    sSQl = VBsprintf(sSQl, mCod_Banco)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewBanco = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewBanco Method", Err.Description
End Function

Public Function ViewGrupo(ByVal sCod_Cliente As String, _
                          ByVal mCod_Grupo As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_ES_GRUPOPRO_ViewShort '$','$' "
    sSQl = VBsprintf(sSQl, sCod_Cliente, mCod_Grupo)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewGrupo = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewGrupo Method", Err.Description
End Function

Public Function AddEStCli(ByVal sCod_Cliente As String, _
                          ByVal sCod_TemCli As String, _
                          ByVal sCod_EstCli As String, _
                          ByVal sDes_EstCli As String, _
                          ByVal sTela As String, _
                          ByVal sCod_EstPro As String, _
                          ByVal sModo As String) As Boolean

    On Error GoTo AddErr

    #If TRANSACTION_SERVER = True Then

        Dim ctxObject As ObjectContext

        Set ctxObject = GetObjectContext()
    #End If

    Dim rs   As Object

    Dim sSQl As String

    sSQl = "exec SI_TG_EstCliShort '$','$','$','$','$','$' , '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_TemCli, sCod_EstCli, sDes_EstCli, sTela, sCod_EstPro, sModo)
    ExecuteCommandSQL mvarConnect, sSQl
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetComplete
    #End If
    AddEStCli = True

    Exit Function

AddErr:
    #If TRANSACTION_SERVER = True Then
        ctxObject.SetAbort
    #End If
    AddEStCli = False
    Err.Raise Err.Number, "AddEStCli Method", Err.Description
End Function

Public Function ViewEStCli(ByVal mCod_Cliente As String, _
                           ByVal mCod_TemCli As String) As Variant

    On Error GoTo ViewAllErr

    Dim sSQl As String

    sSQl = "exec SM_TG_EstCli_ViewxCliente '$' , '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_TemCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
    ViewEStCli = RowsDataSet()

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewEStCli  Method", Err.Description
End Function

Public Function ValidEstilos(ByVal mCod_Cliente As String, _
                             ByVal mCod_EstCli As String, _
                             ByVal mCod_TemCli As String, _
                             ByRef mCod_EstPro As String, _
                             ByRef mCod_GruTAl As String) As String

    On Error GoTo ViewAllErr

    Dim sSQl As String

    mCod_EstPro = ""
 
    sSQl = "exec SM_TG_EstCli_ViewShort '$' , '$' , '$'"
    sSQl = VBsprintf(sSQl, mCod_Cliente, mCod_EstCli, mCod_TemCli)
    Set mRs = GetDataSet(mvarConnect, sSQl)
 
    If mRs.EOF Then
        ValidEstilos = "NO EXISTE"
    Else

        If RTrim(mRs!cod_estcli) <> "" Then
            If FixNulos(mRs!Num_EstProRea, vbstring) <> 1 Then
                ValidEstilos = "TIENE MAS DE 1 ESTILO PROPIO"
                mCod_EstPro = mRs!Cod_EstPro
                mCod_GruTAl = mRs!Cod_GruTal
            Else
                ValidEstilos = "EXISTE"
                mCod_EstPro = mRs!Cod_EstPro
                mCod_GruTAl = mRs!Cod_GruTal
            End If

        ElseIf RTrim(mRs!cod_estcli) = "" And RTrim(mRs!Es_Un_EstiloExistente) = 0 Then
            ValidEstilos = "NO EXISTE"
            mCod_EstPro = ""
        ElseIf RTrim(mRs!cod_estcli) = "" And RTrim(mRs!Es_Un_EstiloExistente) > 0 Then
            ValidEstilos = "ESTILO EXISTE Y NO ASIG A TEMCLI"
        End If
    End If

    Exit Function

ViewAllErr:
    Err.Raise Err.Number, "ViewEstilos Method", Err.Description
End Function

