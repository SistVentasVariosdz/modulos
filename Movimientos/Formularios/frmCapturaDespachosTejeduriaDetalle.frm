VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCapturaDespachosTejeduriaDetalle 
   Caption         =   "Detalle de Movimiento"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   3945
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6959
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmCapturaDespachosTejeduriaDetalle.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmCapturaDespachosTejeduriaDetalle.frx":0352
      Column(2)       =   "frmCapturaDespachosTejeduriaDetalle.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCapturaDespachosTejeduriaDetalle.frx":04BE
      FormatStyle(2)  =   "frmCapturaDespachosTejeduriaDetalle.frx":05F6
      FormatStyle(3)  =   "frmCapturaDespachosTejeduriaDetalle.frx":06A6
      FormatStyle(4)  =   "frmCapturaDespachosTejeduriaDetalle.frx":075A
      FormatStyle(5)  =   "frmCapturaDespachosTejeduriaDetalle.frx":0832
      FormatStyle(6)  =   "frmCapturaDespachosTejeduriaDetalle.frx":08EA
      FormatStyle(7)  =   "frmCapturaDespachosTejeduriaDetalle.frx":09CA
      FormatStyle(8)  =   "frmCapturaDespachosTejeduriaDetalle.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmCapturaDespachosTejeduriaDetalle.frx":12CE
      PrinterProperties=   "frmCapturaDespachosTejeduriaDetalle.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   7650
      TabIndex        =   1
      Top             =   4185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1005
      Custom          =   "0~0~ACEPTAR~Verdadero~Verdadero~&Aceptar~0~0~4~~0~Verdadero~Falso~&Aceptar~"
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmCapturaDespachosTejeduriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String
Public sNum_MovStk As String


Public Function BUSCAR() As Boolean
On Error GoTo Errores
Dim sSQL As String
Dim vBookmark As Variant

sSQL = "INTERFASE_TEJ_CONFECCIONES_VER_DESPACHOS_POR_LEER_DETALLE '$' , '$'"
sSQL = VBsprintf(sSQL, sCod_Almacen, sNum_MovStk)

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 2

Exit Function

Errores:
    err.Raise err.Number, err.Source, err.Description
End Function


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Unload Me
End Sub

Private Sub GridEX1_DblClick()
    Dim i As Integer
    For i = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(i).Caption & Chr(34) & ").width = " & CStr(GridEX1.Columns(i).Width)
    Next
End Sub

