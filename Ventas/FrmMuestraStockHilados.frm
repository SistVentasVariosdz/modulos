VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraStockHilados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Hilados Hilanderia"
   ClientHeight    =   7590
   ClientLeft      =   30
   ClientTop       =   1065
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11895
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      Begin VB.OptionButton optDescripcion 
         Caption         =   "&Descripcion"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCodHilAnt 
         Caption         =   "Codigo Hilado &Ant"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCodHilado 
         Caption         =   "Codigo &Hilado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtBus 
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   225
         Width           =   6975
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6120
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   10795
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "FrmMuestraStockHilados.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmMuestraStockHilados.frx":0352
      Column(2)       =   "FrmMuestraStockHilados.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmMuestraStockHilados.frx":04BE
      FormatStyle(2)  =   "FrmMuestraStockHilados.frx":05F6
      FormatStyle(3)  =   "FrmMuestraStockHilados.frx":06A6
      FormatStyle(4)  =   "FrmMuestraStockHilados.frx":075A
      FormatStyle(5)  =   "FrmMuestraStockHilados.frx":0832
      FormatStyle(6)  =   "FrmMuestraStockHilados.frx":08EA
      FormatStyle(7)  =   "FrmMuestraStockHilados.frx":09CA
      FormatStyle(8)  =   "FrmMuestraStockHilados.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "FrmMuestraStockHilados.frx":12CE
      PrinterProperties=   "FrmMuestraStockHilados.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3960
      TabIndex        =   0
      Top             =   6960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"FrmMuestraStockHilados.frx":17F8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   6240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMuestraStockHilados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipoAdd As String
Dim sCod_Almacen, sDes_Almacen As String
 
Private Sub CONFIGURA_GRID()
On Error GoTo hand

GridEX1.Columns("conchilc").Width = 1095
GridEX1.Columns("conchilc").Caption = "Cod Hilado"
GridEX1.Columns("conccorc").Width = 1050
GridEX1.Columns("conccorc").Caption = "Cod Art"
GridEX1.Columns("contconc").Width = 4905
GridEX1.Columns("conctejc").Width = 1065
GridEX1.Columns("Pre_Hilo").Width = 735
GridEX1.Columns("conctejc").Caption = "Cod Nuevo"
GridEX1.Columns("contconc").Caption = "Desripcion"
GridEX1.Columns("Kilos").Width = 1125
GridEX1.Columns("Kilos").Caption = "Kilos"
GridEX1.Columns("CAJAS").Width = 810
GridEX1.Columns("CAJAS").Caption = "Cajas"
GridEX1.Columns("BOLSAS").Width = 930
GridEX1.Columns("BOLSAS").Caption = "Bolsas"
GridEX1.Columns("OTROS").Width = 825
GridEX1.Columns("OTROS").Caption = "Otros"
GridEX1.Columns("Conos").Width = 750
GridEX1.Columns("Conos").Caption = "Conos"

Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "BUSCAR"

End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  If Not oParent Is Nothing Then oParent.DropWindowList Me.Tag
End Sub
 
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
 
Dim sFecha As Date, StrSql As String
On Error GoTo hand
 
Select Case ActionName
Case "IMPRIMIR"
  REPORTE
Case "BUSCAR"
  BUSCAR
Case "SALIR"
Unload Me
End Select
 
Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, ActionName
End Sub
 
Sub BUSCAR()
 
Dim sFecha As Date, StrSql As String
On Error GoTo hand
 
  VB.Screen.MousePointer = vbHourglass
  StrSql = "HILADO_2004..stockdiarioshilados 'X'"
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
  CONFIGURA_GRID
  VB.Screen.MousePointer = 0
Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "BUSCAR"
End Sub
 
Sub REPORTE()
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, rs As Recordset, lvRuta As String

Set rs = New ADODB.Recordset

  If GridEX1.RowCount = 0 Then Exit Sub

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\Stock_Hilado_Pre.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", GridEX1.ADORecordset, lvRuta
    Set oo = Nothing
    
  
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Stock de Hilados " & Err.Description, vbCritical, "Impresion"
End Sub

Private Sub optCodHilado_Click()
  txtBus.Text = ""
  txtBus.MaxLength = 10
  txtBus.SetFocus
End Sub

Private Sub optCodHilAnt_Click()
  txtBus.Text = ""
  txtBus.MaxLength = 9
  txtBus.SetFocus
End Sub

Private Sub optDescripcion_Click()
  txtBus.Text = ""
  txtBus.MaxLength = 0
  txtBus.SetFocus
End Sub

Private Sub txtBus_Change()

If optCodHilado Then
  Call GridEX1.Find(GridEX1.Columns("Conctejc").Index, jgexContains, txtBus)
End If

If optCodHilAnt Then
  Call GridEX1.Find(GridEX1.Columns("conchilc").Index, jgexContains, txtBus)
End If

If optDescripcion Then
  Call GridEX1.Find(GridEX1.Columns("contconc").Index, jgexContains, txtBus)
End If

End Sub
