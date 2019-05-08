VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowES_OrdProReqItems_Prov 
   Caption         =   "Items-Proveedor"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   1890
      TabIndex        =   0
      Top             =   3135
      Width           =   1245
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   2925
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5159
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmShowES_OrdProReqItems_Prov.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowES_OrdProReqItems_Prov.frx":0352
      Column(2)       =   "frmShowES_OrdProReqItems_Prov.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowES_OrdProReqItems_Prov.frx":04BE
      FormatStyle(2)  =   "frmShowES_OrdProReqItems_Prov.frx":05F6
      FormatStyle(3)  =   "frmShowES_OrdProReqItems_Prov.frx":06A6
      FormatStyle(4)  =   "frmShowES_OrdProReqItems_Prov.frx":075A
      FormatStyle(5)  =   "frmShowES_OrdProReqItems_Prov.frx":0832
      FormatStyle(6)  =   "frmShowES_OrdProReqItems_Prov.frx":08EA
      FormatStyle(7)  =   "frmShowES_OrdProReqItems_Prov.frx":09CA
      FormatStyle(8)  =   "frmShowES_OrdProReqItems_Prov.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmShowES_OrdProReqItems_Prov.frx":12CE
      PrinterProperties=   "frmShowES_OrdProReqItems_Prov.frx":1620
   End
End
Attribute VB_Name = "frmShowES_OrdProReqItems_Prov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String
Public sCod_Item As String
Public sCod_Comb As String
Public sCod_Color As String
Public sCod_Talla As String
Public sCod_destino As String
Public sCod_EstCli As String
Public oParent As Object
Public sSQL As String

Public Function BUSCAR() As Boolean
On Error GoTo Errores
Dim vBookmark As Variant

sSQL = "SM_MUESTRA_COD_PROV  '$','$','$','$','$','$','$' "
sSQL = VBsprintf(sSQL, sCod_Almacen, sCod_Item, sCod_Comb, sCod_Color, sCod_Talla, sCod_destino, sCod_EstCli)

vBookmark = GridEX1.Row
GridEX1.ClearFields


Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Columns("COD_PROV").Width = 1500
GridEX1.Columns("can_stock").Width = 2500


GridEX1.Columns("cod_prov").Caption = "Cod.Proveedor"
GridEX1.Columns("can_stock").Caption = "Stock"

GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 1
Exit Function

Errores:
    Errores err.Number
End Function


Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    
    oParent.gexLista.Value(oParent.gexLista.Columns("COD_PROV").Index) = GridEX1.Value(GridEX1.Columns("COD_PROV").Index)
    
End Sub

