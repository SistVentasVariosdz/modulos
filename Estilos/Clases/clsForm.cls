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
    vemp = mCod_Empresa
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
    cCONNECT = mConnect
End Property
Public Property Let ConnectSeguridad(ByVal vnewvalue As String)
    cSEGURIDAD = vnewvalue
End Property
Public Function GetForm(sFormName) As Object
sFormName = Trim(UCase(sFormName))
Select Case sFormName
    Case "FRMMANTESTALLA"
        Set GetForm = New frmMantESTalla
    Case "FRMMANTTALLADET"
        Set GetForm = New frmMantTallaDet
    Case "FRMMANTCOMHIL"
        Set GetForm = New frmMantComHil
    Case "FRMMANTCOMPEST"
        Set GetForm = New frmMantCompEst
    Case "FRMMANTFAMHIL"
        Set GetForm = New frmMantFamHil
    Case "FRMMANTHILTEL"
        Set GetForm = New frmMantHilTel
    Case "FRMMANTMATPRI"
        Set GetForm = New frmMantMatPri
    Case "FRMMANTPZAEST"
        Set GetForm = New frmMantPzaEst
    Case "FRMMANTTIPCOMP"
        Set GetForm = New frmMantTipComp
    Case "FRMMANITEMS"
        Set GetForm = New frmManItems
    Case "FRMMANTELAS"
        Set GetForm = New frmManTelas

    Case "FRMGRUPOPRO"
        Set GetForm = New frmGrupoPro
    Case "FRMPROCESOS"
        Set GetForm = New frmProcesos
    Case "FRMMUESTRATELASENVIADASDESARROLLO"
        Set GetForm = New frmMuestraTelasEnviadasDesarrollo
    Case "FRMMANTITEMSERVICIOS"
        Set GetForm = New frmMantItemServicios
    Case "FRMMANTFAMITEM"
        Set GetForm = New frmMantFamItem
    Case "FRMESTCLITEM"
        Set GetForm = New frmEstCliTem
        
End Select
End Function
Public Property Get Tag() As String
    Tag = sTag
End Property
Public Property Let Tag(ByVal vnewvalue As String)
    sTag = vnewvalue
End Property
