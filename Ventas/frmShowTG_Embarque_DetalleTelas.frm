VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTG_Embarque_DetalleTelas 
   Caption         =   "Detalle Embarque Telas"
   ClientHeight    =   4044
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10428
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4044
   ScaleWidth      =   10428
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9045
      _ExtentX        =   15960
      _ExtentY        =   6795
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   288
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowTG_Embarque_DetalleTelas.frx":0000
      Column(2)       =   "frmShowTG_Embarque_DetalleTelas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowTG_Embarque_DetalleTelas.frx":016C
      FormatStyle(2)  =   "frmShowTG_Embarque_DetalleTelas.frx":02A4
      FormatStyle(3)  =   "frmShowTG_Embarque_DetalleTelas.frx":0354
      FormatStyle(4)  =   "frmShowTG_Embarque_DetalleTelas.frx":0408
      FormatStyle(5)  =   "frmShowTG_Embarque_DetalleTelas.frx":04E0
      FormatStyle(6)  =   "frmShowTG_Embarque_DetalleTelas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmShowTG_Embarque_DetalleTelas.frx":0678
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   2310
      Left            =   9120
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   4085
      Custom          =   $"frmShowTG_Embarque_DetalleTelas.frx":0850
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowTG_Embarque_DetalleTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lNum_Embarque As Long

Public Function BUSCAR() As Boolean
On Error GoTo errores
Dim ssql As String
Dim vBookmark As Variant

ssql = "TG_Embarques_Telas_Muestra '$'"
ssql = VBsprintf(ssql, lNum_Embarque)
  
vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True
GridEX1.FrozenColumns = 3

GridEX1.Columns("clave").Width = 0

Exit Function

errores:
    errores err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim fnuevo As Boolean
Dim clave As String


Select Case ActionName
  Case "ADICIONAR"
            Dim frmEmbTeI As New frmTG_Embarque_Telas
            frmEmbTeI.lNum_Embarque = lNum_Embarque
            Set frmEmbTeI.oParent = Me
            frmEmbTeI.Saccion = "I"
            frmEmbTeI.Show vbModal
            Set frmEmbTeI = Nothing
            
  Case "MODIFICAR"
            Dim frmEmbTeU As New frmTG_Embarque_Telas
            
            frmEmbTeU.lNum_Embarque = lNum_Embarque
            frmEmbTeU.lSec_Embarque = GridEX1.Value(GridEX1.Columns("Sec_Embarque").Index)
            frmEmbTeU.txtCodTela = GridEX1.Value(GridEX1.Columns("cod_tela").Index)
            frmEmbTeU.txtDesTela = GridEX1.Value(GridEX1.Columns("des_tela").Index)
            frmEmbTeU.txtCodComb = GridEX1.Value(GridEX1.Columns("cod_comb").Index)
            frmEmbTeU.txtDesComb = GridEX1.Value(GridEX1.Columns("des_comb").Index)
            frmEmbTeU.txtCodColor = GridEX1.Value(GridEX1.Columns("cod_color").Index)
            frmEmbTeU.txtDesColor = GridEX1.Value(GridEX1.Columns("des_color").Index)
            frmEmbTeU.txtCodUniMedida = GridEX1.Value(GridEX1.Columns("Uni_Med").Index)
            frmEmbTeU.txtDesUniMedida = GridEX1.Value(GridEX1.Columns("Des_UniMed").Index)
            
            frmEmbTeU.txtPeso_Bruto_Prog = GridEX1.Value(GridEX1.Columns("Peso_Bruto_Prog").Index)
            frmEmbTeU.txtPeso_Neto_Prog = GridEX1.Value(GridEX1.Columns("Peso_Neto_Prog").Index)
            frmEmbTeU.txtRollosProg = GridEX1.Value(GridEX1.Columns("Rollos_Prog").Index)
            frmEmbTeU.txtUbicajeProg = GridEX1.Value(GridEX1.Columns("Cubicaje_Prog").Index)
            frmEmbTeU.txtKgsProg = GridEX1.Value(GridEX1.Columns("Kgs_Prog").Index)
            frmEmbTeU.txtUnidadesProg = GridEX1.Value(GridEX1.Columns("Unidades_Prog").Index)
                                    
            
            Set frmEmbTeU.oParent = Me
            frmEmbTeU.Saccion = "U"
            
            clave = GridEX1.Value(GridEX1.Columns("clave").Index)
            
            frmEmbTeU.Show vbModal
            fnuevo = GridEX1.Find(GridEX1.Columns("clave").Index, jgexGreaterThanOrEqualTo, clave)
            
            Set frmEmbTeU = Nothing
  Case "ELIMINAR"
            Eliminar_Datos
            BUSCAR
  Case "SALIR"
            Unload Me
End Select
End Sub

Private Sub Eliminar_Datos()
On Error GoTo errx
Dim ssql As String

ssql = "TG_Embarque_Telas_man '$',$,$,'$','$','$','$',$,$,$,$,$,$"
  
ssql = VBsprintf(ssql, "D", lNum_Embarque, GridEX1.Value(GridEX1.Columns("Sec_Embarque").Index), "", "", "", "", 0, 0, 0, 0, 0, 0)
  


ExecuteCommandSQL cCONNECT, ssql

MsgBox "Los datos fueron procesados con éxito.", vbInformation, "Mensaje"
  

Exit Sub
errx:
    errores err.Number
End Sub
