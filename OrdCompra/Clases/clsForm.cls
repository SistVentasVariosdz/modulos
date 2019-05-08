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
    cConnect = mConnect
End Property
Public Property Let ConnectSeguridad(ByVal vnewvalue As String)
    cSEGURIDAD = vnewvalue
End Property
Public Function GetForm(sFormName) As Object
sFormName = Trim(UCase(sFormName))
Select Case sFormName
    Case "FRMORDCOMP"
        Set GetForm = New frmOrdComp
    Case "FRMSERVTEN"
        Set GetForm = New frmStockServTenido
    Case "FRMREPORTEOCOMPRAITEMS"
     Set GetForm = New FrmReporteOCompraItems
   
End Select
End Function
