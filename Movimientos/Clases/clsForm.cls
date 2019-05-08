VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConnect As String

Public Property Get ConnectEmpresa() As String
    ConnectEmpresa = cConnect
End Property

Public Property Get Language() As Integer
    Language = iLanguage
End Property
Public Property Let Language(ByVal vnewvalue As Integer)
    iLanguage = vnewvalue
End Property

Public Property Get Parent() As Object
    Set Parent = oMDIParent
End Property

Public Property Let Parent(ByVal vnewvalue As Object)
    Set oMDIParent = vnewvalue
End Property

Public Property Let Rutas(ByVal mRutas As Variant)
    vRuta = mRutas
End Property
Public Property Get Rutas()
    Rutas = vRuta
End Property
Public Property Let Cod_Empresa(ByVal mCod_Empresa As String)
    'vemp = mCod_Empresa
    vemp1 = mCod_Empresa
    bCargaConexion = True
End Property
Public Property Let Cod_Perfil(ByVal mCod_Perfil As String)
    vper = mCod_Perfil
End Property
Public Property Let Cod_Opcion(ByVal mCod_Opcion As String)
    'gsCod_Opcion = mCod_Opcion
End Property
Public Property Let Key_Node(ByVal mKey_Node As String)
    'gsKey_Node = mKey_Node
End Property
Public Property Let Permisos(ByVal mPermisos As String)
    'gsPermisos = mPermisos
End Property
Public Property Let Nom_Empresa(ByVal vnewvalue As String)
    'gsNomEmpresa = vnewvalue
End Property
Public Property Let Nom_Aplicacion(ByVal vnewvalue As String)
    'gsNomAplicacion = vnewvalue
End Property
Public Property Let Nom_Opcion(ByVal vnewvalue As String)
    vNomFor = vnewvalue
End Property
Public Property Let UserName(ByVal vnewvalue As String)
    vusu = vnewvalue
End Property
Public Property Get UserName() As String
    UserName = vusu
End Property
Public Property Let ConnectEmpresa(ByVal vnewvalue As String)
    mConnect = vnewvalue
    'vemp = vnewvalue
    cConnect = mConnect
End Property
Public Property Let ConnectSeguridad(ByVal vnewvalue As String)
    cSEGURIDAD = vnewvalue
End Property
Public Function GetForm(sFormName) As Object
sFormName = Trim(UCase(sFormName))
Select Case sFormName
    Case "FRMMANTALMACEN"
        Set GetForm = New FrmMantAlmacen
    Case "FRMMANTTIPMOV"
        Set GetForm = New FrmMantTipMov
    Case "FRMCLAORDCOMP"
        Set GetForm = New frmClaOrdComp
    Case "FRMMANTMOVPERM"
        Set GetForm = New FrmMantMovPerm
    Case "FRMMOVALMACEN"
        Set GetForm = New FrmMovAlmacen
    Case "FRMCLAORDCOMP"
        Set GetForm = New frmClaOrdComp
    Case "FRMKARDEX"
        Set GetForm = New FrmKardex
    Case "FRMREP"
        Set GetForm = New FrmRep
    Case "FRMKARDEXTELACA"
        Set GetForm = New FrmKardexTelaca
    Case "FRMSTOCKFAM"
        Set GetForm = New FrmStockFam
        
    Case "FRMKARDEXHILCRU"
        Set GetForm = New FrmKardexHilCru
        
    Case "FRMKARDEXHILTEN"
        Set GetForm = New FrmKardexHilTen
        
    Case "FRMKARDEXTELCRU"
        Set GetForm = New FrmKardexTelCru
        
    Case "FRMMOVSTOCKS"
        Set GetForm = New frmMovStocks
        
    Case "FRMDATOSTECNICOS"
        Set GetForm = New frmDatosTecnicos

'    Case "FRMMOVSTOCKSGUIAS"
'        Set GetForm = New frmMovStocksGuias

    Case "FRMREQVSREALAVIOS"
        Set GetForm = New frmReqVsRealAvios
'    Case "FRMFLUJOPRODUCCIONDIARIA"
'        Set GetForm = New frmFlujoProduccionDiaria
    Case "FRMCONTROLPRODMENSUAL"
        Set GetForm = New frmControlProdMensual
    'Case "FRMSTOCKSSALDOS"
    '    Set GetForm = New frmStocksSaldos
    Case "FRMRPTSTKTELAS"
        Set GetForm = frmRptStkTelas
    Case "FRMRPTINGXCOMPROBANTE"
        Set GetForm = frmRptIngxComprobante
    Case "FRMACCCF_SEGALM"
        Set GetForm = frmAccCF_SEGALM
    Case "FRMACCLG_SEGALM"
        Set GetForm = frmAccLG_SEGALM
    Case "FRM_STOCKCRITICOS"
            Set GetForm = Frm_StockCriticos
    'Case "FRMRECLAMOSAVIOSPROD"
        'Set GetForm = New frmReclamosAviosProd
End Select
End Function
    
